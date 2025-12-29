

/**
 * Attendance Missions Service for Cosmic Event Manager
 * Version 7.9.7
 *
 * @fileoverview Manages the Attendance_Missions sheet, calculating total
 * attendance mission points for each player based on mission column values.
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Sheet name for attendance missions tracking
 * @const {string}
 */
var ATTENDANCE_MISSIONS_SHEET_NAME = 'Attendance_Missions';

/**
 * Header name for the total points column
 * @const {string}
 */
var ATTENDANCE_POINTS_HEADER = 'Attendance Missions Points';

/**
 * List of mission column headers that contribute to the Attendance Missions Points total.
 * These columns are summed for each player row.
 *
 * Note: "Casual Commander Events" and "Transitional Commander Events" are both
 * Commander attendance buckets - they represent different tiers of Commander play
 * but are weighted equally in the sum.
 *
 * @const {string[]}
 */
var MISSION_COLUMNS_TO_SUM = [
  'First Contact',
  'Stellar Explorer',
  'Deck Diver',
  'Lunar Loyalty',
  'Meteor Shower',
  'Sealed Voyager',
  'Draft Navigator',
  'Stellar Scholar',
  // Commander attendance buckets (both count equally):
  'Casual Commander Events',       // Commander bucket #1
  'Transitional Commander Events', // Commander bucket #2
  'cEDH Events',
  'Limited Events',
  'Academy Events',
  'Outreach Events',
  'Free Play Events',
  'Interstellar Strategist',
  'Black Hole Survivor'
];

/**
 * Expected header order after the points column is properly positioned.
 * Used for validation and reordering if necessary.
 * @const {string[]}
 */
var EXPECTED_HEADER_ORDER = [
  'PreferredName',
  'Attendance Missions Points',
  'First Contact',
  'Stellar Explorer',
  'Deck Diver',
  'Lunar Loyalty',
  'Meteor Shower',
  'Sealed Voyager',
  'Draft Navigator',
  'Stellar Scholar',
  'Casual Commander Events',
  'Transitional Commander Events',
  'cEDH Events',
  'Limited Events',
  'Academy Events',
  'Outreach Events',
  'Free Play Events',
  'Interstellar Strategist',
  'Black Hole Survivor',
  'Total Events Attended'
];

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Recalculates the Attendance Missions Points for all players in the
 * Attendance_Missions sheet.
 *
 * This function:
 * 1. Opens the active spreadsheet and gets the Attendance_Missions sheet
 * 2. Ensures the "Attendance Missions Points" column exists in position B
 * 3. Builds a header-to-column-index map for dynamic column lookup
 * 4. Iterates through all player rows (row 2 onward)
 * 5. Sums the numeric values of all mission columns for each player
 * 6. Writes the totals back to column B using batch operations
 *
 * @throws {Error} If the Attendance_Missions sheet is missing
 * @throws {Error} If any required mission headers are missing
 */
function recalculateAttendanceMissionPoints() {
  Logger.log('Starting recalculateAttendanceMissionPoints()');

  // Get the active spreadsheet and target sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ATTENDANCE_MISSIONS_SHEET_NAME);

  // Validate sheet exists
  if (!sheet) {
    var errorMsg = 'Sheet "' + ATTENDANCE_MISSIONS_SHEET_NAME + '" not found. ' +
                   'Please run Build/Repair to create required sheets.';
    Logger.log('ERROR: ' + errorMsg);
    throw new Error(errorMsg);
  }

  Logger.log('Found sheet: ' + ATTENDANCE_MISSIONS_SHEET_NAME);

  // Ensure the Attendance Missions Points column is in the correct position (B)
  ensurePointsColumnInPosition_(sheet);

  // Build header map (must be done AFTER ensuring column position)
  var headerMap = getAttendanceHeaderMap_(sheet);
  Logger.log('Header map built with ' + Object.keys(headerMap).length + ' columns');

  // Get all data (for batch read)
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    Logger.log('No player data rows found (only header row exists)');
    return;
  }

  // Read all data at once (batch read for performance)
  var dataRange = sheet.getRange(1, 1, lastRow, lastCol);
  var allData = dataRange.getValues();

  Logger.log('Processing ' + (lastRow - 1) + ' player rows');

  // Calculate points for each player row
  var pointsColumn = []; // Will hold [points] for each row

  // First element is header (already set)
  pointsColumn.push([ATTENDANCE_POINTS_HEADER]);

  // Process each player row (starting from row 2, index 1)
  for (var rowIndex = 1; rowIndex < allData.length; rowIndex++) {
    var rowData = allData[rowIndex];
    var playerName = rowData[0]; // Column A is PreferredName

    // Skip empty rows (no player name)
    if (!playerName || String(playerName).trim() === '') {
      pointsColumn.push(['']);
      continue;
    }

    // Sum all mission columns for this player
    var totalPoints = sumMissionColumns_(rowData, headerMap);
    pointsColumn.push([totalPoints]);

    Logger.log('Row ' + (rowIndex + 1) + ' (' + playerName + '): ' + totalPoints + ' points');
  }

  // Write all points back to column B (batch write for performance)
  // Column B is index 2 (1-based)
  var pointsRange = sheet.getRange(1, 2, pointsColumn.length, 1);
  pointsRange.setValues(pointsColumn);

  Logger.log('Successfully updated ' + (pointsColumn.length - 1) + ' player point totals');
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Builds a map of header names to their 0-based column indexes.
 * Reads row 1 of the sheet and creates a lookup object.
 *
 * @param {Sheet} sheet - The Google Sheets Sheet object
 * @return {Object} Map of header name (string) -> column index (0-based number)
 * @throws {Error} If any required mission headers are missing
 * @private
 */
function getAttendanceHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    throw new Error('Sheet has no columns. Please verify sheet structure.');
  }

  // Read header row (row 1)
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues()[0];

  // Build the map
  var headerMap = {};
  for (var i = 0; i < headers.length; i++) {
    var headerName = String(headers[i]).trim();
    if (headerName) {
      headerMap[headerName] = i;
    }
  }

  // Validate that all required mission columns exist
  var missingHeaders = [];
  for (var j = 0; j < MISSION_COLUMNS_TO_SUM.length; j++) {
    var requiredHeader = MISSION_COLUMNS_TO_SUM[j];
    if (headerMap[requiredHeader] === undefined) {
      missingHeaders.push(requiredHeader);
    }
  }

  if (missingHeaders.length > 0) {
    var errorMsg = 'Missing required mission headers: ' + missingHeaders.join(', ') + '\n' +
                   'Please verify the Attendance_Missions sheet has all required columns.';
    Logger.log('ERROR: ' + errorMsg);
    throw new Error(errorMsg);
  }

  return headerMap;
}

/**
 * Ensures the "Attendance Missions Points" column exists in position B (column 2).
 *
 * - If the column doesn't exist, inserts a new column B and sets the header
 * - If the column exists but is in the wrong position, moves it to column B
 *
 * @param {Sheet} sheet - The Google Sheets Sheet object
 * @private
 */
function ensurePointsColumnInPosition_(sheet) {
  var lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    Logger.log('Sheet is empty, cannot ensure column position');
    return;
  }

  // Read current headers
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues()[0];

  // Find current position of the points column (if it exists)
  var currentPosition = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === ATTENDANCE_POINTS_HEADER) {
      currentPosition = i;
      break;
    }
  }

  Logger.log('Current position of "' + ATTENDANCE_POINTS_HEADER + '": ' +
             (currentPosition === -1 ? 'NOT FOUND' : 'Column ' + (currentPosition + 1)));

  // Target position is column B (index 1, or column number 2)
  var targetPosition = 1; // 0-based index

  if (currentPosition === -1) {
    // Column doesn't exist - insert new column at position B
    Logger.log('Inserting new column B for "' + ATTENDANCE_POINTS_HEADER + '"');
    sheet.insertColumnAfter(1); // Insert after column A (becomes new column B)
    sheet.getRange(1, 2).setValue(ATTENDANCE_POINTS_HEADER);

  } else if (currentPosition !== targetPosition) {
    // Column exists but is in wrong position - need to move it
    Logger.log('Moving "' + ATTENDANCE_POINTS_HEADER + '" from column ' +
               (currentPosition + 1) + ' to column 2');

    // Strategy: Delete the existing column and insert at correct position
    // First, save all data from the points column
    var lastRow = sheet.getLastRow();
    var pointsData = [];

    if (lastRow > 0) {
      pointsData = sheet.getRange(1, currentPosition + 1, lastRow, 1).getValues();
    }

    // Delete the column from its current position
    sheet.deleteColumn(currentPosition + 1);

    // Insert new column at position B (after A)
    sheet.insertColumnAfter(1);

    // Restore the data
    if (pointsData.length > 0) {
      sheet.getRange(1, 2, pointsData.length, 1).setValues(pointsData);
    } else {
      sheet.getRange(1, 2).setValue(ATTENDANCE_POINTS_HEADER);
    }
  } else {
    Logger.log('"' + ATTENDANCE_POINTS_HEADER + '" is already in correct position (column B)');
  }
}

/**
 * Sums the numeric values of all mission columns for a single row.
 *
 * This function iterates through MISSION_COLUMNS_TO_SUM and adds up the
 * numeric values from each column. Non-numeric values and blanks are treated as 0.
 *
 * Note: Both "Casual Commander Events" and "Transitional Commander Events" are
 * included in the sum as Commander attendance buckets. They are weighted equally
 * (no special multiplier) but represent different tiers of Commander play.
 *
 * @param {Array} rowData - Array of cell values for the row (from getValues())
 * @param {Object} headerMap - Map of header names to 0-based column indexes
 * @return {number} Total points for this row
 * @private
 */
function sumMissionColumns_(rowData, headerMap) {
  var total = 0;

  for (var i = 0; i < MISSION_COLUMNS_TO_SUM.length; i++) {
    var columnName = MISSION_COLUMNS_TO_SUM[i];
    var columnIndex = headerMap[columnName];

    // Skip if column not found (shouldn't happen if validation passed)
    if (columnIndex === undefined) {
      Logger.log('WARNING: Column "' + columnName + '" not found in header map');
      continue;
    }

    var cellValue = rowData[columnIndex];
    var numericValue = safeParseNumber_(cellValue);

    total += numericValue;
  }

  return total;
}

/**
 * Safely parses a cell value to a number.
 * Returns 0 for blanks, empty strings, and non-numeric values.
 *
 * @param {*} value - Cell value to parse
 * @return {number} Numeric value or 0
 * @private
 */
function safeParseNumber_(value) {
  // Handle null, undefined, empty string
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  // Convert to number
  var num = Number(value);

  // Return 0 if NaN (non-numeric)
  if (isNaN(num)) {
    return 0;
  }

  return num;
}

// ============================================================================
// MENU INTEGRATION (OPTIONAL)
// ============================================================================

/**
 * Creates or updates a custom menu for Attendance Missions operations.
 * Can be called from onOpen() or used standalone.
 *
 * Adds menu item: "Attendance Missions" -> "Recalculate Attendance Points"
 */
function onAttendanceMissionsMenu() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('Attendance Missions')
    .addItem('Recalculate Attendance Points', 'recalculateAttendanceMissionPoints')
    .addToUi();

  Logger.log('Attendance Missions menu created');
}

/**
 * Wrapper function to run recalculation with user confirmation.
 * Useful for menu integration with a confirmation step.
 */
function recalculateAttendanceMissionPointsWithConfirm() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert(
    'Recalculate Attendance Mission Points',
    'This will recalculate the "Attendance Missions Points" column for all players ' +
    'in the Attendance_Missions sheet.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    try {
      recalculateAttendanceMissionPoints();
      ui.alert(
        'Success',
        'Attendance Mission Points have been recalculated for all players.',
        ui.ButtonSet.OK
      );
    } catch (e) {
      ui.alert(
        'Error',
        'Failed to recalculate points:\n\n' + e.message,
        ui.ButtonSet.OK
      );
      Logger.log('Error in recalculateAttendanceMissionPointsWithConfirm: ' + e.message);
    }
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Validates the header structure of the Attendance_Missions sheet.
 * Returns an object with validation results and any issues found.
 *
 * @return {Object} {valid: boolean, issues: string[], headerMap: Object}
 */
function validateAttendanceMissionsStructure() {
  var result = {
    valid: true,
    issues: [],
    headerMap: null
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ATTENDANCE_MISSIONS_SHEET_NAME);

  if (!sheet) {
    result.valid = false;
    result.issues.push('Sheet "' + ATTENDANCE_MISSIONS_SHEET_NAME + '" not found');
    return result;
  }

  try {
    result.headerMap = getAttendanceHeaderMap_(sheet);
  } catch (e) {
    result.valid = false;
    result.issues.push(e.message);
    return result;
  }

  // Check if PreferredName column exists
  if (result.headerMap['PreferredName'] === undefined) {
    result.valid = false;
    result.issues.push('Missing "PreferredName" column');
  }

  // Check if Total Events Attended exists
  if (result.headerMap['Total Events Attended'] === undefined) {
    result.issues.push('Warning: "Total Events Attended" column not found');
  }

  Logger.log('Validation result: ' + (result.valid ? 'VALID' : 'INVALID'));
  if (result.issues.length > 0) {
    Logger.log('Issues: ' + result.issues.join('; '));
  }

  return result;
}

/**
 * Debug function to log the current header structure of the Attendance_Missions sheet.
 * Useful for troubleshooting column order issues.
 */
function debugLogAttendanceMissionsHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ATTENDANCE_MISSIONS_SHEET_NAME);

  if (!sheet) {
    Logger.log('Sheet not found: ' + ATTENDANCE_MISSIONS_SHEET_NAME);
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  Logger.log('=== Current Header Structure ===');
  for (var i = 0; i < headers.length; i++) {
    Logger.log('Column ' + String.fromCharCode(65 + i) + ' (' + (i + 1) + '): "' + headers[i] + '"');
  }
  Logger.log('================================');
}