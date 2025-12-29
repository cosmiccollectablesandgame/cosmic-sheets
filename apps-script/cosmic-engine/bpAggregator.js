/**
 * BP Aggregator - Bonus Points Sync & Refresh
 * @fileoverview Aggregates BP from source sheets into BP_Total with menu, full refresh, and onEdit sync
 *
 * Source Sheets:
 * - Attendance_Missions: PreferredName, "Attendance Missions Points" OR "Attendance Missions"
 * - Flag_Missions: PreferredName, Flag Points
 * - Dice_Points: PreferredName, Points
 * - Redeemed_BP: PreferredName, Total_Redeemed
 *
 * Target Sheet:
 * - BP_Total: preferred_name_id, Current_BP, Attendance Missions, Flag Missions, Dice Roll Points, Historical_BP, LastUpdated
 */

// ============================================================================
// CONSTANTS
// ============================================================================

const BP_AGGREGATOR_CONFIG = {
  // Sheet names
  SHEETS: {
    BP_TOTAL: 'BP_Total',
    ATTENDANCE: 'Attendance_Missions',
    FLAG: 'Flag_Missions',
    DICE: 'Dice_Points',
    REDEEMED: 'Redeemed_BP'
  },

  // BP_Total column headers (canonical)
  BP_TOTAL_HEADERS: [
    'preferred_name_id',
    'Current_BP',
    'Attendance Missions',
    'Flag Missions',
    'Dice Roll Points',
    'Historical_BP',
    'LastUpdated'
  ],

  // Column indices for BP_Total (0-indexed)
  COLS: {
    NAME: 0,
    CURRENT_BP: 1,
    ATTENDANCE: 2,
    FLAG: 3,
    DICE: 4,
    HISTORICAL: 5,
    UPDATED: 6
  },

  // BP constraints
  MIN_BP: 0,
  MAX_BP: 100
};

// ============================================================================
// MENU SETUP
// ============================================================================

/**
 * Adds the Cosmic Bonus Points menu to the spreadsheet UI.
 * Call this from onOpen() in Code.gs
 */
function addBonusPointsMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Cosmic Bonus Points')
    .addItem('Refresh BP Totals', 'refreshAllBonusPoints')
    .addSeparator()
    .addItem('Validate BP Sheets', 'validateBPSheets')
    .addToUi();
}

// ============================================================================
// MAIN REFRESH FUNCTION
// ============================================================================

/**
 * Refreshes all Bonus Points by aggregating from source sheets.
 * This is the core function called by the menu item.
 */
function refreshAllBonusPoints() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Get all required sheets
    const bpTotalSheet = getOrCreateBPTotalSheet_(ss);
    const attendanceSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.ATTENDANCE);
    const flagSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.FLAG);
    const diceSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.DICE);
    const redeemedSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.REDEEMED);

    // 2. Build source maps from each sheet
    const attendanceMap = buildAttendanceMap_(attendanceSheet);
    const flagMap = buildNameToValueMap_(flagSheet, 'PreferredName', 'Flag Points');
    const diceMap = buildNameToValueMap_(diceSheet, 'PreferredName', 'Points');
    const redeemedMap = buildNameToValueMap_(redeemedSheet, 'PreferredName', 'Total_Redeemed');

    // 3. Collect all unique player names from all sources
    const allPlayerNames = collectAllPlayerNames_(
      bpTotalSheet,
      attendanceMap,
      flagMap,
      diceMap,
      redeemedMap
    );

    // 4. Read current BP_Total data
    const bpTotalData = bpTotalSheet.getDataRange().getValues();
    const existingPlayerRows = buildExistingPlayerMap_(bpTotalData);

    // 5. Build output data array
    const outputData = buildOutputData_(
      allPlayerNames,
      existingPlayerRows,
      bpTotalData,
      attendanceMap,
      flagMap,
      diceMap,
      redeemedMap
    );

    // 6. Write all data back to BP_Total in one batch
    writeBPTotalData_(bpTotalSheet, outputData);

    // 7. Show success toast
    ss.toast(
      `Bonus Points refreshed for ${allPlayerNames.size} player(s).`,
      'BP Refresh Complete',
      5
    );

    // 8. Log the action
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('BP_AGGREGATE_REFRESH', {
        playerCount: allPlayerNames.size,
        timestamp: new Date().toISOString(),
        status: 'SUCCESS'
      });
    }

  } catch (error) {
    // Show error toast
    ss.toast(
      `Error: ${error.message}`,
      'BP Refresh Failed',
      10
    );

    // Log error
    console.error('refreshAllBonusPoints failed:', error);

    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('BP_AGGREGATE_REFRESH', {
        error: error.message,
        status: 'FAILED'
      });
    }

    throw error;
  }
}

// ============================================================================
// PER-PLAYER HELPER
// ============================================================================

/**
 * Refreshes Bonus Points for a single player.
 * Used by onEdit for responsive updates.
 *
 * @param {string} playerName - The player's canonical name (PreferredName / preferred_name_id)
 * @private
 */
function refreshBonusPointsForPlayer_(playerName) {
  if (!playerName || typeof playerName !== 'string') {
    return;
  }

  const normalizedName = normalizePlayerName_(playerName);
  if (!normalizedName) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Get sheets
    const bpTotalSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.BP_TOTAL);
    if (!bpTotalSheet) {
      console.warn('BP_Total sheet not found, skipping per-player refresh');
      return;
    }

    const attendanceSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.ATTENDANCE);
    const flagSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.FLAG);
    const diceSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.DICE);
    const redeemedSheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.REDEEMED);

    // Build source maps
    const attendanceMap = buildAttendanceMap_(attendanceSheet);
    const flagMap = buildNameToValueMap_(flagSheet, 'PreferredName', 'Flag Points');
    const diceMap = buildNameToValueMap_(diceSheet, 'PreferredName', 'Points');
    const redeemedMap = buildNameToValueMap_(redeemedSheet, 'PreferredName', 'Total_Redeemed');

    // Get values for this player
    const attendanceBP = getMapValue_(attendanceMap, normalizedName);
    const flagBP = getMapValue_(flagMap, normalizedName);
    const diceBP = getMapValue_(diceMap, normalizedName);
    const redeemedBP = getMapValue_(redeemedMap, normalizedName);

    // Calculate totals
    const rawTotal = attendanceBP + flagBP + diceBP - redeemedBP;
    const currentBP = Math.max(BP_AGGREGATOR_CONFIG.MIN_BP, Math.min(BP_AGGREGATOR_CONFIG.MAX_BP, rawTotal));
    const historicalBP = attendanceBP + flagBP + diceBP;
    const lastUpdated = new Date();

    // Find player row in BP_Total
    const bpTotalData = bpTotalSheet.getDataRange().getValues();
    let playerRowIndex = -1;

    for (let i = 1; i < bpTotalData.length; i++) {
      const rowName = normalizePlayerName_(String(bpTotalData[i][BP_AGGREGATOR_CONFIG.COLS.NAME] || ''));
      if (rowName === normalizedName) {
        playerRowIndex = i;
        break;
      }
    }

    if (playerRowIndex === -1) {
      // Player not found, append new row
      const newRow = [
        normalizedName,
        currentBP,
        attendanceBP,
        flagBP,
        diceBP,
        historicalBP,
        lastUpdated
      ];
      bpTotalSheet.appendRow(newRow);
    } else {
      // Update existing row (columns B-G, which is 2-7 in 1-indexed)
      const updateRange = bpTotalSheet.getRange(playerRowIndex + 1, 2, 1, 6);
      updateRange.setValues([[
        currentBP,
        attendanceBP,
        flagBP,
        diceBP,
        historicalBP,
        lastUpdated
      ]]);
    }

  } catch (error) {
    console.error(`refreshBonusPointsForPlayer_ failed for "${playerName}":`, error);
  }
}

// ============================================================================
// ONEDIT TRIGGER
// ============================================================================

/**
 * Handles edit events to keep BP_Total in sync.
 *
 * When a cell is edited in one of the source sheets (Attendance_Missions,
 * Flag_Missions, Dice_Points, Redeemed_BP), this function updates the
 * corresponding player's row in BP_Total.
 *
 * @param {Object} e - The onEdit event object
 */
function onEditBPAggregator(e) {
  // Guard against missing event object
  if (!e || !e.range) {
    return;
  }

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Check if this is a source sheet we care about
  const sourceSheets = [
    BP_AGGREGATOR_CONFIG.SHEETS.ATTENDANCE,
    BP_AGGREGATOR_CONFIG.SHEETS.FLAG,
    BP_AGGREGATOR_CONFIG.SHEETS.DICE,
    BP_AGGREGATOR_CONFIG.SHEETS.REDEEMED
  ];

  if (!sourceSheets.includes(sheetName)) {
    return;
  }

  // Get the edited row
  const editedRow = e.range.getRow();

  // Skip header row
  if (editedRow <= 1) {
    return;
  }

  // Get player name from column A of the edited row
  const playerName = sheet.getRange(editedRow, 1).getValue();

  if (!playerName) {
    return;
  }

  // Normalize and refresh
  const normalizedName = normalizePlayerName_(String(playerName));

  if (normalizedName) {
    refreshBonusPointsForPlayer_(normalizedName);
  }
}

// ============================================================================
// VALIDATION
// ============================================================================

/**
 * Validates that all required BP sheets exist and have correct schemas.
 * Shows a dialog with validation results.
 */
function validateBPSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const results = [];

  // Check BP_Total
  const bpTotal = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.BP_TOTAL);
  if (bpTotal) {
    const headers = bpTotal.getRange(1, 1, 1, 7).getValues()[0];
    const expectedHeaders = BP_AGGREGATOR_CONFIG.BP_TOTAL_HEADERS;
    const headersMatch = expectedHeaders.every((h, i) => headers[i] === h);
    results.push(`BP_Total: ${headersMatch ? 'OK' : 'Headers mismatch'}`);
  } else {
    results.push('BP_Total: Missing (will be created on refresh)');
  }

  // Check Attendance_Missions
  const attendance = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.ATTENDANCE);
  if (attendance) {
    const headers = attendance.getRange(1, 1, 1, attendance.getLastColumn()).getValues()[0];
    const hasName = headers.includes('PreferredName');
    const hasPoints = headers.includes('Attendance Missions Points') || headers.includes('Attendance Missions');
    results.push(`Attendance_Missions: ${hasName && hasPoints ? 'OK' : 'Missing required columns'}`);
  } else {
    results.push('Attendance_Missions: Missing');
  }

  // Check Flag_Missions
  const flag = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.FLAG);
  if (flag) {
    const headers = flag.getRange(1, 1, 1, flag.getLastColumn()).getValues()[0];
    const hasName = headers.includes('PreferredName');
    const hasPoints = headers.includes('Flag Points');
    results.push(`Flag_Missions: ${hasName && hasPoints ? 'OK' : 'Missing required columns'}`);
  } else {
    results.push('Flag_Missions: Missing');
  }

  // Check Dice_Points
  const dice = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.DICE);
  if (dice) {
    const headers = dice.getRange(1, 1, 1, dice.getLastColumn()).getValues()[0];
    const hasName = headers.includes('PreferredName');
    const hasPoints = headers.includes('Points');
    results.push(`Dice_Points: ${hasName && hasPoints ? 'OK' : 'Missing required columns'}`);
  } else {
    results.push('Dice_Points: Missing');
  }

  // Check Redeemed_BP
  const redeemed = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.REDEEMED);
  if (redeemed) {
    const headers = redeemed.getRange(1, 1, 1, redeemed.getLastColumn()).getValues()[0];
    const hasName = headers.includes('PreferredName');
    const hasTotal = headers.includes('Total_Redeemed');
    results.push(`Redeemed_BP: ${hasName && hasTotal ? 'OK' : 'Missing required columns'}`);
  } else {
    results.push('Redeemed_BP: Missing');
  }

  ui.alert(
    'BP Sheets Validation',
    results.join('\n'),
    ui.ButtonSet.OK
  );
}

// ============================================================================
// HELPER FUNCTIONS - Name Normalization
// ============================================================================

/**
 * Normalizes a player name for consistent matching.
 * Trims whitespace and treats as case-sensitive string.
 *
 * @param {string} name - Raw player name
 * @return {string} Normalized name, or empty string if invalid
 * @private
 */
function normalizePlayerName_(name) {
  if (name === null || name === undefined) {
    return '';
  }

  const trimmed = String(name).trim();

  // Return empty string for blank names
  if (trimmed === '') {
    return '';
  }

  return trimmed;
}

// ============================================================================
// HELPER FUNCTIONS - Map Building
// ============================================================================

/**
 * Builds a map from player name to value from a sheet.
 *
 * @param {Sheet} sheet - The sheet to read from (can be null)
 * @param {string} nameHeader - Header for the name column
 * @param {string} valueHeader - Header for the value column
 * @return {Map<string, number>} Map of normalized player names to values
 * @private
 */
function buildNameToValueMap_(sheet, nameHeader, valueHeader) {
  const map = new Map();

  if (!sheet) {
    return map;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return map;
  }

  const headers = data[0];
  const nameCol = headers.indexOf(nameHeader);
  const valueCol = headers.indexOf(valueHeader);

  if (nameCol === -1) {
    console.warn(`buildNameToValueMap_: "${nameHeader}" column not found in ${sheet.getName()}`);
    return map;
  }

  if (valueCol === -1) {
    console.warn(`buildNameToValueMap_: "${valueHeader}" column not found in ${sheet.getName()}`);
    return map;
  }

  // Process data rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const rawName = data[i][nameCol];
    const rawValue = data[i][valueCol];

    const normalizedName = normalizePlayerName_(rawName);

    if (normalizedName) {
      const numericValue = coerceToNumber_(rawValue);
      map.set(normalizedName, numericValue);
    }
  }

  return map;
}

/**
 * Builds the attendance map, detecting which column header is present.
 * Looks for "Attendance Missions Points" first, then "Attendance Missions".
 *
 * @param {Sheet} sheet - The Attendance_Missions sheet (can be null)
 * @return {Map<string, number>} Map of normalized player names to attendance BP
 * @private
 */
function buildAttendanceMap_(sheet) {
  const map = new Map();

  if (!sheet) {
    return map;
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return map;
  }

  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');

  if (nameCol === -1) {
    console.warn('buildAttendanceMap_: "PreferredName" column not found in Attendance_Missions');
    return map;
  }

  // Detect which attendance column header is present
  let valueCol = headers.indexOf('Attendance Missions Points');

  if (valueCol === -1) {
    valueCol = headers.indexOf('Attendance Missions');
  }

  if (valueCol === -1) {
    console.warn('buildAttendanceMap_: Neither "Attendance Missions Points" nor "Attendance Missions" column found');
    return map;
  }

  // Process data rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const rawName = data[i][nameCol];
    const rawValue = data[i][valueCol];

    const normalizedName = normalizePlayerName_(rawName);

    if (normalizedName) {
      const numericValue = coerceToNumber_(rawValue);
      map.set(normalizedName, numericValue);
    }
  }

  return map;
}

/**
 * Builds a map of existing player names to their row indices in BP_Total.
 *
 * @param {Array<Array>} bpTotalData - 2D array of BP_Total data
 * @return {Map<string, number>} Map of normalized names to row indices (0-indexed)
 * @private
 */
function buildExistingPlayerMap_(bpTotalData) {
  const map = new Map();

  // Skip header row
  for (let i = 1; i < bpTotalData.length; i++) {
    const rawName = bpTotalData[i][BP_AGGREGATOR_CONFIG.COLS.NAME];
    const normalizedName = normalizePlayerName_(rawName);

    if (normalizedName) {
      map.set(normalizedName, i);
    }
  }

  return map;
}

// ============================================================================
// HELPER FUNCTIONS - Data Collection
// ============================================================================

/**
 * Collects all unique player names from BP_Total and all source maps.
 *
 * @param {Sheet} bpTotalSheet - The BP_Total sheet
 * @param {Map<string, number>} attendanceMap - Attendance map
 * @param {Map<string, number>} flagMap - Flag missions map
 * @param {Map<string, number>} diceMap - Dice points map
 * @param {Map<string, number>} redeemedMap - Redeemed BP map
 * @return {Set<string>} Set of all unique normalized player names
 * @private
 */
function collectAllPlayerNames_(bpTotalSheet, attendanceMap, flagMap, diceMap, redeemedMap) {
  const allNames = new Set();

  // Add names from BP_Total
  if (bpTotalSheet && bpTotalSheet.getLastRow() > 1) {
    const data = bpTotalSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const normalizedName = normalizePlayerName_(data[i][BP_AGGREGATOR_CONFIG.COLS.NAME]);
      if (normalizedName) {
        allNames.add(normalizedName);
      }
    }
  }

  // Add names from source maps
  for (const name of attendanceMap.keys()) {
    allNames.add(name);
  }

  for (const name of flagMap.keys()) {
    allNames.add(name);
  }

  for (const name of diceMap.keys()) {
    allNames.add(name);
  }

  for (const name of redeemedMap.keys()) {
    allNames.add(name);
  }

  return allNames;
}

// ============================================================================
// HELPER FUNCTIONS - Output Building
// ============================================================================

/**
 * Builds the complete output data array for BP_Total.
 *
 * @param {Set<string>} allPlayerNames - Set of all player names
 * @param {Map<string, number>} existingPlayerRows - Map of names to existing row indices
 * @param {Array<Array>} currentData - Current BP_Total data
 * @param {Map<string, number>} attendanceMap - Attendance map
 * @param {Map<string, number>} flagMap - Flag missions map
 * @param {Map<string, number>} diceMap - Dice points map
 * @param {Map<string, number>} redeemedMap - Redeemed BP map
 * @return {Array<Array>} Complete output data including header
 * @private
 */
function buildOutputData_(
  allPlayerNames,
  existingPlayerRows,
  currentData,
  attendanceMap,
  flagMap,
  diceMap,
  redeemedMap
) {
  const output = [];
  const timestamp = new Date();

  // Add header row
  output.push(BP_AGGREGATOR_CONFIG.BP_TOTAL_HEADERS.slice());

  // Convert set to sorted array for consistent ordering
  const sortedNames = Array.from(allPlayerNames).sort();

  for (const playerName of sortedNames) {
    // Get values from each source (defaults to 0 if not found)
    const attendanceBP = getMapValue_(attendanceMap, playerName);
    const flagBP = getMapValue_(flagMap, playerName);
    const diceBP = getMapValue_(diceMap, playerName);
    const redeemedBP = getMapValue_(redeemedMap, playerName);

    // Calculate Current_BP with clamping
    const rawTotal = attendanceBP + flagBP + diceBP - redeemedBP;
    const currentBP = Math.max(
      BP_AGGREGATOR_CONFIG.MIN_BP,
      Math.min(BP_AGGREGATOR_CONFIG.MAX_BP, rawTotal)
    );

    // Calculate Historical_BP (lifetime earned, no redemptions subtracted)
    const historicalBP = attendanceBP + flagBP + diceBP;

    // Build row
    const row = [
      playerName,        // preferred_name_id (A)
      currentBP,         // Current_BP (B)
      attendanceBP,      // Attendance Missions (C)
      flagBP,            // Flag Missions (D)
      diceBP,            // Dice Roll Points (E)
      historicalBP,      // Historical_BP (F)
      timestamp          // LastUpdated (G)
    ];

    output.push(row);
  }

  return output;
}

/**
 * Gets a value from a map, returning 0 if not found.
 *
 * @param {Map<string, number>} map - The map to query
 * @param {string} key - The key to look up
 * @return {number} The value, or 0 if not found
 * @private
 */
function getMapValue_(map, key) {
  return map.has(key) ? map.get(key) : 0;
}

// ============================================================================
// HELPER FUNCTIONS - Data Writing
// ============================================================================

/**
 * Writes the complete output data to BP_Total sheet.
 *
 * @param {Sheet} sheet - The BP_Total sheet
 * @param {Array<Array>} data - Complete output data including header
 * @private
 */
function writeBPTotalData_(sheet, data) {
  if (data.length === 0) {
    return;
  }

  // Clear existing data (preserve sheet)
  sheet.clear();

  // Write all data in one batch
  const numRows = data.length;
  const numCols = data[0].length;

  sheet.getRange(1, 1, numRows, numCols).setValues(data);

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Auto-resize columns for readability
  for (let col = 1; col <= numCols; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Gets or creates the BP_Total sheet with proper schema.
 *
 * @param {Spreadsheet} ss - The active spreadsheet
 * @return {Sheet} The BP_Total sheet
 * @private
 */
function getOrCreateBPTotalSheet_(ss) {
  let sheet = ss.getSheetByName(BP_AGGREGATOR_CONFIG.SHEETS.BP_TOTAL);

  if (!sheet) {
    // Create new sheet
    sheet = ss.insertSheet(BP_AGGREGATOR_CONFIG.SHEETS.BP_TOTAL);

    // Add headers
    sheet.appendRow(BP_AGGREGATOR_CONFIG.BP_TOTAL_HEADERS);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, BP_AGGREGATOR_CONFIG.BP_TOTAL_HEADERS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Freeze header row
    sheet.setFrozenRows(1);

    console.log('Created BP_Total sheet with schema');
  }

  return sheet;
}

// ============================================================================
// HELPER FUNCTIONS - Type Coercion
// ============================================================================

/**
 * Coerces a value to a number, returning 0 for invalid values.
 * Uses the existing coerceNumber if available, otherwise implements locally.
 *
 * @param {*} value - Value to coerce
 * @return {number} Numeric value, or 0 if invalid
 * @private
 */
function coerceToNumber_(value) {
  // Use existing utility if available
  if (typeof coerceNumber === 'function') {
    return coerceNumber(value, 0);
  }

  // Local implementation
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

// ============================================================================
// INSTALLABLE TRIGGER SETUP
// ============================================================================

/**
 * Creates an installable onEdit trigger for the BP Aggregator.
 * Run this once to set up automatic syncing.
 *
 * Note: Simple onEdit triggers have limitations. For full functionality,
 * this creates an installable trigger that can access services.
 */
function setupBPAggregatorTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check for existing trigger
  const triggers = ScriptApp.getUserTriggers(ss);
  const existingTrigger = triggers.find(t =>
    t.getHandlerFunction() === 'onEditBPAggregator' &&
    t.getEventType() === ScriptApp.EventType.ON_EDIT
  );

  if (existingTrigger) {
    SpreadsheetApp.getUi().alert(
      'Trigger Exists',
      'The BP Aggregator onEdit trigger is already set up.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  // Create new trigger
  ScriptApp.newTrigger('onEditBPAggregator')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger Created',
    'The BP Aggregator onEdit trigger has been set up. BP_Total will now auto-update when source sheets change.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Removes the BP Aggregator onEdit trigger.
 */
function removeBPAggregatorTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getUserTriggers(ss);

  let removed = 0;

  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onEditBPAggregator') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  }

  SpreadsheetApp.getUi().alert(
    'Trigger Removed',
    removed > 0
      ? `Removed ${removed} BP Aggregator trigger(s).`
      : 'No BP Aggregator triggers found.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}