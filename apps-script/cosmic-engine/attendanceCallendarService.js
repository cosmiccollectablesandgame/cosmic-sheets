/**
 * Attendance Calendar Service - Visual Matrix for Cosmic Event Manager
 * Version 7.9.7
 *
 * @fileoverview Attendance Calendar visual matrix: players vs events with week banding
 * and month borders. Creates/maintains an Attendance_Calendar sheet showing player
 * attendance across all events in a chronological grid format.
 *
 * Features:
 * - Rows = players (one row per PreferredName)
 * - Columns = events (chronological order, left to right)
 * - Light blue fill for attendance
 * - 45-degree rotated event headers
 * - Alternating week banding (light gray)
 * - Bold vertical borders between months
 * - Total Events count per player
 */

// ============================================================================
// CONSTANTS
// ============================================================================

/** Light blue fill for attendance cells */
const ATTENDANCE_FILL_COLOR = '#cfe2f3';

/** Light gray for odd-week banding */
const WEEK_BAND_COLOR = '#f3f3f3';

/** Standard header background */
const HEADER_BG_COLOR = '#4285f4';

/** Standard header text color */
const HEADER_TEXT_COLOR = '#ffffff';

/** Attendance marker character */
const ATTENDANCE_MARKER = '✓';

// ============================================================================
// MAIN ENTRY POINT
// ============================================================================

/**
 * Builds or rebuilds the Attendance_Calendar sheet from scratch.
 * This is the primary entry point for the attendance calendar feature.
 *
 * Steps:
 * 1. Gather canonical players from PreferredNames (or fallback sources)
 * 2. Gather and sort event sheets by date
 * 3. Build header row (PreferredName, Total Events, event columns)
 * 4. Write all player names into Column A
 * 5. For each event column, mark attendance with blue fill
 * 6. Populate Total Events formulas in Column B
 * 7. Apply formatting (rotation, banding, borders, freeze panes)
 *
 * @return {Object} Summary {playerCount, eventCount, rebuilt}
 */
function buildAttendanceCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Get canonical players (sorted alphabetically)
  const players = getAllPlayers_();
  if (players.length === 0) {
    throw new Error('[CALENDAR_NO_PLAYERS] No players found in PreferredNames or canonical sheets.');
  }

  // Step 2: Get event sheets metadata sorted by date
  const eventsMeta = getEventSheets_();
  if (eventsMeta.length === 0) {
    throw new Error('[CALENDAR_NO_EVENTS] No event sheets found matching MM-DD-YYYY or MM-DDX-YYYY pattern.');
  }

  // Step 3: Get or create the Attendance_Calendar sheet
  let sheet = ss.getSheetByName('Attendance_Calendar');
  const isRebuild = !!sheet;

  if (sheet) {
    // Clear existing content and formatting
    sheet.clear();
    sheet.clearFormats();
    sheet.clearConditionalFormatRules();
  } else {
    sheet = ss.insertSheet('Attendance_Calendar');
  }

  // Step 4: Build the data matrix in memory for batch write
  const dataMatrix = buildDataMatrix_(players, eventsMeta, ss);

  // Step 5: Write data to sheet in one batch operation
  const numRows = dataMatrix.length;
  const numCols = dataMatrix[0].length;
  sheet.getRange(1, 1, numRows, numCols).setValues(dataMatrix);

  // Step 6: Apply Total Events formulas in column B
  applyTotalEventsFormulas_(sheet, players.length, eventsMeta.length);

  // Step 7: Apply all formatting
  applyHeaderFormatting_(sheet, eventsMeta.length);
  applyWeeklyBanding_(sheet, eventsMeta, players.length);
  applyMonthBorders_(sheet, eventsMeta, players.length);
  applyAttendanceFills_(sheet, dataMatrix, eventsMeta.length);

  // Step 8: Set column widths and freeze panes
  finalizeLayout_(sheet, eventsMeta.length);

  // Log the action
  logIntegrityAction('CALENDAR_BUILD', {
    details: `Built attendance calendar: ${players.length} players, ${eventsMeta.length} events`,
    status: 'SUCCESS'
  });

  return {
    playerCount: players.length,
    eventCount: eventsMeta.length,
    rebuilt: isRebuild
  };
}

// ============================================================================
// DATA GATHERING HELPERS
// ============================================================================

/**
 * Gets all event sheets with metadata, sorted chronologically.
 * Matches patterns: MM-DD-YYYY or MM-DDX-YYYY (X = A-Z suffix)
 *
 * @return {Array<Object>} Array of {sheetName, eventDate, weekNumber, month, year}
 * @private
 */
function getEventSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Pattern: MM-DD-YYYY or MM-DDX-YYYY (X = optional A-Z suffix)
  // Examples: 11-26-2025, 11-26C-2025
  const eventPattern = /^(\d{2})-(\d{2})([A-Z])?-(\d{4})$/;

  const events = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    const match = name.match(eventPattern);

    if (match) {
      const month = parseInt(match[1], 10);
      const day = parseInt(match[2], 10);
      const suffix = match[3] || ''; // Optional A-Z suffix
      const year = parseInt(match[4], 10);

      // Create date object for sorting and grouping
      const eventDate = new Date(year, month - 1, day);

      // Calculate ISO week number for banding
      const weekNumber = getISOWeekNumber_(eventDate);

      events.push({
        sheetName: name,
        eventDate: eventDate,
        month: month,
        year: year,
        day: day,
        suffix: suffix,
        weekNumber: weekNumber,
        weekYear: getISOWeekYear_(eventDate)
      });
    }
  });

  // Sort chronologically by date, then by suffix
  events.sort((a, b) => {
    const dateCompare = a.eventDate.getTime() - b.eventDate.getTime();
    if (dateCompare !== 0) return dateCompare;
    return (a.suffix || '').localeCompare(b.suffix || '');
  });

  return events;
}

/**
 * Gets all canonical players sorted alphabetically.
 * First checks for PreferredNames sheet, then falls back to Key_Tracker/BP_Total.
 *
 * @return {Array<string>} Sorted array of PreferredName strings
 * @private
 */
function getAllPlayers_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const players = new Set();

  // Primary source: PreferredNames sheet (if exists)
  const preferredNamesSheet = ss.getSheetByName('PreferredNames');
  if (preferredNamesSheet && preferredNamesSheet.getLastRow() > 1) {
    const data = preferredNamesSheet.getDataRange().getValues();
    const headers = data[0];

    // Find PreferredName column (could be column A or named header)
    let nameCol = 0; // Default to column A
    const headerIndex = headers.findIndex(h =>
      normalizeHeader(String(h)) === 'PreferredName' ||
      String(h).toLowerCase() === 'preferredname' ||
      String(h).toLowerCase() === 'preferred_name_id'
    );
    if (headerIndex !== -1) nameCol = headerIndex;

    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][nameCol]).trim();
      if (name && name !== '') {
        players.add(name);
      }
    }
  }

  // Fallback: Use existing getCanonicalNames() which reads from Key_Tracker and BP_Total
  if (players.size === 0) {
    const canonicalNames = getCanonicalNames();
    canonicalNames.forEach(name => players.add(name));
  }

  // Return sorted array
  return Array.from(players).sort((a, b) =>
    a.toLowerCase().localeCompare(b.toLowerCase())
  );
}

/**
 * Gets the roster (set of player names) from an event sheet.
 * Looks for PreferredName or preferred_name_id column (treated as synonyms).
 *
 * @param {Sheet} sheet - Event sheet to read
 * @return {Set<string>} Set of player names in the roster
 * @private
 */
function getEventRoster_(sheet) {
  const roster = new Set();

  if (!sheet || sheet.getLastRow() <= 1) {
    return roster;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());

  // Find PreferredName or preferred_name_id column
  const nameColIndex = headers.findIndex(h => {
    const normalized = normalizeHeader(h);
    const lower = h.toLowerCase();
    return normalized === 'PreferredName' ||
           lower === 'preferredname' ||
           lower === 'preferred_name_id';
  });

  // Default to column B (index 1) if not found by header
  const nameCol = nameColIndex !== -1 ? nameColIndex : 1;

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol]).trim();
    if (name && name !== '') {
      roster.add(name);
    }
  }

  return roster;
}

// ============================================================================
// DATA MATRIX BUILDER
// ============================================================================

/**
 * Builds the complete data matrix for the calendar sheet.
 *
 * @param {Array<string>} players - Sorted player names
 * @param {Array<Object>} eventsMeta - Event metadata array
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array<Array>} 2D array for batch writing
 * @private
 */
function buildDataMatrix_(players, eventsMeta, ss) {
  const numRows = players.length + 1; // +1 for header
  const numCols = 2 + eventsMeta.length; // PreferredName + Total Events + events

  // Initialize matrix
  const matrix = [];

  // Build header row
  const headerRow = ['PreferredName', 'Total Events'];
  eventsMeta.forEach(event => {
    headerRow.push(event.sheetName);
  });
  matrix.push(headerRow);

  // Pre-load all event rosters for efficiency
  const eventRosters = new Map();
  eventsMeta.forEach(event => {
    const sheet = ss.getSheetByName(event.sheetName);
    eventRosters.set(event.sheetName, getEventRoster_(sheet));
  });

  // Build player rows
  players.forEach(player => {
    const row = [player, '']; // Name and placeholder for formula

    eventsMeta.forEach(event => {
      const roster = eventRosters.get(event.sheetName);
      if (roster.has(player)) {
        row.push(ATTENDANCE_MARKER);
      } else {
        row.push('');
      }
    });

    matrix.push(row);
  });

  return matrix;
}

// ============================================================================
// FORMATTING FUNCTIONS
// ============================================================================

/**
 * Applies Total Events formulas in column B.
 * Uses COUNTIF to count non-blank cells across event columns.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {number} playerCount - Number of player rows
 * @param {number} eventCount - Number of event columns
 * @private
 */
function applyTotalEventsFormulas_(sheet, playerCount, eventCount) {
  if (playerCount === 0 || eventCount === 0) return;

  const firstEventCol = 3; // Column C
  const lastEventCol = 2 + eventCount;
  const lastColLetter = columnToLetter_(lastEventCol);

  const formulas = [];
  for (let i = 0; i < playerCount; i++) {
    const rowNum = i + 2; // Data starts at row 2
    formulas.push([`=COUNTIF(C${rowNum}:${lastColLetter}${rowNum},"<>")`]);
  }

  sheet.getRange(2, 2, playerCount, 1).setFormulas(formulas);
  sheet.getRange(2, 2, playerCount, 1)
    .setHorizontalAlignment('right')
    .setNumberFormat('0');
}

/**
 * Applies header row formatting.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {number} eventCount - Number of event columns
 * @private
 */
function applyHeaderFormatting_(sheet, eventCount) {
  const lastCol = 2 + eventCount;

  // Style entire header row
  const headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange
    .setFontWeight('bold')
    .setBackground(HEADER_BG_COLOR)
    .setFontColor(HEADER_TEXT_COLOR)
    .setVerticalAlignment('bottom')
    .setHorizontalAlignment('center');

  // Apply 45-degree rotation to event column headers (C1 onward)
  if (eventCount > 0) {
    const eventHeaderRange = sheet.getRange(1, 3, 1, eventCount);
    eventHeaderRange.setTextRotation(45);
  }

  // Make PreferredName and Total Events headers normal (no rotation)
  sheet.getRange(1, 1, 1, 2).setTextRotation(0);
}

/**
 * Applies alternating week banding to event columns.
 * Odd weeks get light gray background.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {Array<Object>} eventsMeta - Event metadata array
 * @param {number} playerCount - Number of player rows
 * @private
 */
function applyWeeklyBanding_(sheet, eventsMeta, playerCount) {
  if (eventsMeta.length === 0) return;

  const totalRows = playerCount + 1; // Include header

  // Group events by week
  const weekGroups = [];
  let currentWeekKey = null;
  let currentGroup = [];

  eventsMeta.forEach((event, idx) => {
    const weekKey = `${event.weekYear}-W${event.weekNumber}`;

    if (weekKey !== currentWeekKey) {
      if (currentGroup.length > 0) {
        weekGroups.push(currentGroup);
      }
      currentGroup = [idx];
      currentWeekKey = weekKey;
    } else {
      currentGroup.push(idx);
    }
  });

  if (currentGroup.length > 0) {
    weekGroups.push(currentGroup);
  }

  // Apply banding to odd week groups
  weekGroups.forEach((group, groupIdx) => {
    if (groupIdx % 2 === 1) { // Odd groups get banding
      group.forEach(eventIdx => {
        const colNum = 3 + eventIdx; // Events start at column C
        const range = sheet.getRange(1, colNum, totalRows, 1);
        range.setBackground(WEEK_BAND_COLOR);
      });
    }
  });
}

/**
 * Applies bold vertical borders between months.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {Array<Object>} eventsMeta - Event metadata array
 * @param {number} playerCount - Number of player rows
 * @private
 */
function applyMonthBorders_(sheet, eventsMeta, playerCount) {
  if (eventsMeta.length <= 1) return;

  const totalRows = playerCount + 1; // Include header

  // Find month transitions
  for (let i = 1; i < eventsMeta.length; i++) {
    const prevEvent = eventsMeta[i - 1];
    const currEvent = eventsMeta[i];

    // Check for month change (different month OR different year)
    const monthChanged = prevEvent.month !== currEvent.month ||
                         prevEvent.year !== currEvent.year;

    if (monthChanged) {
      const colNum = 3 + i; // Events start at column C
      const range = sheet.getRange(1, colNum, totalRows, 1);
      range.setBorder(
        null,   // top
        null,   // right
        null,   // bottom
        true,   // left - BOLD BORDER
        null,   // vertical
        null,   // horizontal
        '#000000', // color
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM // thick border
      );
    }
  }
}

/**
 * Applies light blue fill to attendance cells.
 * This overwrites any week banding for attended cells.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {Array<Array>} dataMatrix - The data matrix
 * @param {number} eventCount - Number of event columns
 * @private
 */
function applyAttendanceFills_(sheet, dataMatrix, eventCount) {
  if (eventCount === 0 || dataMatrix.length <= 1) return;

  // Collect ranges to fill (more efficient than cell-by-cell)
  const rangesToFill = [];

  for (let row = 1; row < dataMatrix.length; row++) { // Skip header
    for (let col = 2; col < dataMatrix[row].length; col++) { // Skip name and total
      if (dataMatrix[row][col] === ATTENDANCE_MARKER) {
        const sheetRow = row + 1; // Convert to 1-indexed
        const sheetCol = col + 1; // Convert to 1-indexed
        rangesToFill.push(`${columnToLetter_(sheetCol)}${sheetRow}`);
      }
    }
  }

  // Apply fills in batches using RangeList
  if (rangesToFill.length > 0) {
    // Process in chunks to avoid timeout (max 500 ranges at a time)
    const chunkSize = 500;
    for (let i = 0; i < rangesToFill.length; i += chunkSize) {
      const chunk = rangesToFill.slice(i, i + chunkSize);
      const rangeList = sheet.getRangeList(chunk);
      rangeList.setBackground(ATTENDANCE_FILL_COLOR);
    }
  }
}

/**
 * Finalizes layout with column widths and freeze panes.
 *
 * @param {Sheet} sheet - Calendar sheet
 * @param {number} eventCount - Number of event columns
 * @private
 */
function finalizeLayout_(sheet, eventCount) {
  // Freeze header row and name column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  // Set column widths
  sheet.setColumnWidth(1, 150); // PreferredName
  sheet.setColumnWidth(2, 80);  // Total Events

  // Set event columns to narrow width (rotated headers save space)
  if (eventCount > 0) {
    sheet.setColumnWidths(3, eventCount, 30);
  }

  // Set row height for header to accommodate rotated text
  sheet.setRowHeight(1, 80);
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Converts column number to letter(s) (1 = A, 26 = Z, 27 = AA, etc.)
 *
 * @param {number} col - Column number (1-indexed)
 * @return {string} Column letter(s)
 * @private
 */
function columnToLetter_(col) {
  let letter = '';
  let temp = col;

  while (temp > 0) {
    temp--;
    letter = String.fromCharCode(65 + (temp % 26)) + letter;
    temp = Math.floor(temp / 26);
  }

  return letter;
}

/**
 * Gets ISO week number for a date.
 *
 * @param {Date} date - Date object
 * @return {number} ISO week number (1-53)
 * @private
 */
function getISOWeekNumber_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Gets ISO week year for a date (may differ from calendar year for dates near year boundaries).
 *
 * @param {Date} date - Date object
 * @return {number} ISO week year
 * @private
 */
function getISOWeekYear_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  return d.getUTCFullYear();
}

// ============================================================================
// UI HANDLER
// ============================================================================

/**
 * Handler for menu action - shows confirmation and runs build.
 * Called from the Ops menu.
 */
function onRebuildAttendanceCalendar() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'Rebuild Attendance Calendar',
    'This will rebuild the Attendance_Calendar sheet from scratch.\n\n' +
    'All existing data and formatting will be replaced.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    return;
  }

  try {
    ui.alert(
      'Building...',
      'Please wait while the attendance calendar is being built.\n\n' +
      'This may take 30-60 seconds for large datasets.',
      ui.ButtonSet.OK
    );

    const summary = buildAttendanceCalendarSheet();

    ui.alert(
      'Attendance Calendar Built!',
      `Successfully built attendance calendar:\n\n` +
      `- Players: ${summary.playerCount}\n` +
      `- Events: ${summary.eventCount}\n` +
      `- Action: ${summary.rebuilt ? 'Rebuilt existing' : 'Created new'}\n\n` +
      `Navigate to the "Attendance_Calendar" tab to view.`,
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert(
      'Error Building Calendar',
      `Failed to build attendance calendar:\n\n${e.message}\n\n` +
      'Please check that you have event sheets and player data.',
      ui.ButtonSet.OK
    );
    console.error('Attendance calendar build failed:', e);
  }
}