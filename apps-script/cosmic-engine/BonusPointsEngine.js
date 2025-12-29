/**
 * Bonus Points Engine - BP_Total Synchronization & Redemption Ledger
 * @fileoverview Safe, header-driven BP_Total sync from mission sources + redemption log
 *
 * BP_Total Schema:
 *   PreferredName | BP_Current | Attendance Mission Points | Flag Mission Points |
 *   Dice Roll Points | LastUpdated | BP_Historical | BP_Redeemed
 *
 * Formula: BP_Current = BP_Historical - BP_Redeemed
 * Where:   BP_Historical = Attendance + Flag + Dice (sum of all earned)
 */

// ============================================================================
// CONSTANTS
// ============================================================================

/**
 * Canonical BP_Total headers in required order
 * @const {Array<string>}
 */
const BP_TOTAL_HEADERS = [
  'PreferredName',
  'BP_Current',
  'Attendance Mission Points',
  'Flag Mission Points',
  'Dice Roll Points',
  'LastUpdated',
  'BP_Historical',
  'BP_Redeemed'
];

/**
 * Canonical BP_Redeemed_Log headers
 * @const {Array<string>}
 */
const BP_REDEEMED_LOG_HEADERS = [
  'Timestamp',
  'PreferredName',
  'BP_Amount',
  'Reason',
  'Category',
  'Event_ID',
  'Staff',
  'RowId'
];

/**
 * Header synonyms for reading legacy data
 * @const {Object}
 */
const BP_HEADER_SYNONYMS = {
  // PreferredName synonyms
  'Preferred Name': 'PreferredName',
  'preferred_name_id': 'PreferredName',
  'Player': 'PreferredName',
  'Player Name': 'PreferredName',
  'Name': 'PreferredName',

  // BP_Current synonyms
  'Current_BP': 'BP_Current',
  'Current BP': 'BP_Current',
  'BP': 'BP_Current',
  'Bonus_Points': 'BP_Current',

  // BP_Historical synonyms
  'Historical_BP': 'BP_Historical',
  'Historical BP': 'BP_Historical',
  'Total_Earned': 'BP_Historical',

  // BP_Redeemed synonyms
  'Redeemed_BP': 'BP_Redeemed',
  'Redeemed BP': 'BP_Redeemed',
  'BP Redeemed': 'BP_Redeemed',
  'Total_Spent': 'BP_Redeemed',

  // Points column synonyms
  'Points': 'Points',
  'Attendance Points': 'Attendance Mission Points',
  'Flag Points': 'Flag Mission Points',
  'Dice Points': 'Dice Roll Points'
};

// ============================================================================
// SCHEMA ENFORCEMENT
// ============================================================================

/**
 * Ensures BP_Total sheet exists with correct 8-column schema.
 * Does NOT delete extra columns to the right (leaves anything beyond column 8 alone).
 * Uses synonym mapping to preserve existing data when renaming headers.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The BP_Total sheet
 * @private
 */
function ensureBPTotalSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  // Create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet('BP_Total');
    sheet.getRange(1, 1, 1, BP_TOTAL_HEADERS.length).setValues([BP_TOTAL_HEADERS]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_TOTAL_HEADERS.length);
    return sheet;
  }

  // Handle empty sheet
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BP_TOTAL_HEADERS.length).setValues([BP_TOTAL_HEADERS]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_TOTAL_HEADERS.length);
    return sheet;
  }

  // Read existing headers
  const lastCol = Math.max(sheet.getLastColumn(), BP_TOTAL_HEADERS.length);
  const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Build map of existing header -> column index (1-based)
  const existingMap = {};
  existingHeaders.forEach((h, idx) => {
    if (h) {
      const canonical = BP_HEADER_SYNONYMS[h] || h;
      existingMap[canonical] = idx + 1;
    }
  });

  // Ensure each required header exists in correct position
  const newHeaders = [...existingHeaders];

  for (let i = 0; i < BP_TOTAL_HEADERS.length; i++) {
    const required = BP_TOTAL_HEADERS[i];
    const targetCol = i; // 0-indexed in array

    if (newHeaders[targetCol] !== required) {
      // Check if this header exists elsewhere (possibly as synonym)
      const existingCol = existingMap[required];

      if (existingCol && existingCol !== targetCol + 1) {
        // Header exists in wrong position - we need to be careful here
        // For safety, just update the header name at target position
        newHeaders[targetCol] = required;
      } else {
        // Header doesn't exist or is already in position
        newHeaders[targetCol] = required;
      }
    }
  }

  // Write corrected headers (only first 8 columns, leave rest alone)
  sheet.getRange(1, 1, 1, BP_TOTAL_HEADERS.length).setValues([BP_TOTAL_HEADERS]);
  sheet.setFrozenRows(1);
  formatHeaderRow_(sheet, BP_TOTAL_HEADERS.length);

  return sheet;
}

/**
 * Ensures BP_Redeemed_Log sheet exists with correct schema.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The BP_Redeemed_Log sheet
 * @private
 */
function ensureBPRedeemedLogSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Redeemed_Log');

  if (!sheet) {
    sheet = ss.insertSheet('BP_Redeemed_Log');
    sheet.getRange(1, 1, 1, BP_REDEEMED_LOG_HEADERS.length).setValues([BP_REDEEMED_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_REDEEMED_LOG_HEADERS.length, '#d32f2f'); // Red theme for spend log
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BP_REDEEMED_LOG_HEADERS.length).setValues([BP_REDEEMED_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_REDEEMED_LOG_HEADERS.length, '#d32f2f');
  }

  return sheet;
}

/**
 * Formats header row with styling
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet to format
 * @param {number} numCols - Number of header columns
 * @param {string} color - Background color (default: blue)
 * @private
 */
function formatHeaderRow_(sheet, numCols, color = '#4285f4') {
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setBackground(color)
    .setFontColor('#ffffff');
}

// ============================================================================
// POINTS MAP HELPERS
// ============================================================================

/**
 * Reads points from a source sheet and returns a map of PreferredName -> Points.
 * Handles header synonyms for flexibility.
 *
 * @param {string} sheetName - Name of the sheet to read
 * @param {string} keyHeaderName - Header name for the player column (or synonyms)
 * @param {string} pointsHeaderName - Header name for the points column
 * @return {Object} Map of { 'PlayerName': pointsValue, ... }
 * @private
 */
function getPointsMapFromSheet_(sheetName, keyHeaderName, pointsHeaderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return {}; // Sheet missing or empty (only header)
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find key column (player name) with synonym support
  let keyCol = -1;
  const keyAliases = ['PreferredName', 'Preferred Name', 'preferred_name_id', 'Player', 'Player Name', 'Name'];

  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim();
    if (h === keyHeaderName || keyAliases.includes(h)) {
      keyCol = i;
      break;
    }
  }

  // Find points column with synonym support
  let pointsCol = -1;
  const pointsAliases = [pointsHeaderName, 'Points', pointsHeaderName.replace(/ /g, '_')];

  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]).trim();
    if (pointsAliases.includes(h)) {
      pointsCol = i;
      break;
    }
  }

  if (keyCol === -1 || pointsCol === -1) {
    return {}; // Required columns not found
  }

  // Build map
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][keyCol] || '').trim();
    const points = coerceNumber(data[i][pointsCol], 0);

    if (name) {
      // Sum if player appears multiple times
      map[name] = (map[name] || 0) + points;
    }
  }

  return map;
}

/**
 * Gets combined dice points from both "Dice Roll Points" and "Dice_Points" sheets.
 * Returns summed values if both exist.
 *
 * @return {Object} Map of { 'PlayerName': totalDicePoints, ... }
 * @private
 */
function getCombinedDicePointsMap_() {
  const map1 = getPointsMapFromSheet_('Dice Roll Points', 'PreferredName', 'Points');
  const map2 = getPointsMapFromSheet_('Dice_Points', 'PreferredName', 'Points');

  // Combine maps
  const combined = { ...map1 };

  for (const name of Object.keys(map2)) {
    combined[name] = (combined[name] || 0) + map2[name];
  }

  return combined;
}

/**
 * Reads BP_Redeemed_Log and returns sum of BP_Amount per player.
 *
 * @return {Object} Map of { 'PlayerName': totalRedeemed, ... }
 * @private
 */
function getRedeemedMapFromLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Redeemed_Log');

  if (!sheet || sheet.getLastRow() <= 1) {
    return {}; // No redemptions yet
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find columns
  const nameCol = headers.indexOf('PreferredName');
  const amountCol = headers.indexOf('BP_Amount');

  if (nameCol === -1 || amountCol === -1) {
    return {};
  }

  const map = {};

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    const amount = coerceNumber(data[i][amountCol], 0);

    if (name && amount > 0) {
      map[name] = (map[name] || 0) + amount;
    }
  }

  return map;
}

/**
 * Gets all PreferredNames from the canonical PreferredNames sheet.
 *
 * @return {Array<string>} List of player names
 * @private
 */
function getAllPreferredNames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');

  if (!sheet) {
    throw new Error('[SHEET_MISSING] PreferredNames sheet not found. Create it first.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return []; // Only header or empty
  }

  const data = sheet.getRange('A2:A' + lastRow).getValues();
  const names = [];

  for (let i = 0; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) {
      names.push(name);
    }
  }

  return names;
}

// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Rebuilds BP_Total from mission sources + redemption log.
 *
 * Sources:
 *  - Attendance_Missions!PreferredName, Attendance Mission Points
 *  - Flag_Missions!PreferredName, Flag Mission Points
 *  - Dice Roll Points!PreferredName, Points (and legacy Dice_Points)
 *  - BP_Redeemed_Log!PreferredName, BP_Amount
 *
 * For each PreferredName in PreferredNames:
 *  - Ensures they have a row in BP_Total
 *  - Updates:
 *      Attendance Mission Points
 *      Flag Mission Points
 *      Dice Roll Points
 *      BP_Historical  = sum of the above (for now)
 *      BP_Redeemed    = sum of BP_Amount in BP_Redeemed_Log
 *      BP_Current     = BP_Historical - BP_Redeemed
 *      LastUpdated    = now
 *
 * @return {number} Number of players updated
 */
function updateBPTotalFromSources() {
  // Ensure schemas
  const bpSheet = ensureBPTotalSchema_();
  ensureBPRedeemedLogSchema_(); // Ensure log exists (no-op if already there)

  // Load all canonical player names
  let allNames;
  try {
    allNames = getAllPreferredNames_();
  } catch (e) {
    // If PreferredNames doesn't exist, fall back to Key_Tracker
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const keySheet = ss.getSheetByName('Key_Tracker');
    if (!keySheet) {
      throw new Error('[SHEET_MISSING] Neither PreferredNames nor Key_Tracker found.');
    }
    const lastRow = keySheet.getLastRow();
    if (lastRow <= 1) {
      return 0;
    }
    const data = keySheet.getRange('A2:A' + lastRow).getValues();
    allNames = data.map(row => String(row[0] || '').trim()).filter(n => n);
  }

  if (allNames.length === 0) {
    return 0;
  }

  // Build points maps from sources
  const attendanceMap = getPointsMapFromSheet_('Attendance_Missions', 'PreferredName', 'Attendance Mission Points');
  const flagMap = getPointsMapFromSheet_('Flag_Missions', 'PreferredName', 'Flag Mission Points');
  const diceMap = getCombinedDicePointsMap_();
  const redeemedMap = getRedeemedMapFromLog_();

  // Read existing BP_Total data
  const existingData = bpSheet.getDataRange().getValues();
  const headers = existingData[0];

  // Build column index map
  const colIdx = {};
  for (let i = 0; i < headers.length; i++) {
    colIdx[headers[i]] = i;
  }

  // Verify required columns exist
  const requiredCols = ['PreferredName', 'BP_Current', 'Attendance Mission Points',
                        'Flag Mission Points', 'Dice Roll Points', 'LastUpdated',
                        'BP_Historical', 'BP_Redeemed'];
  for (const col of requiredCols) {
    if (colIdx[col] === undefined) {
      throw new Error(`[SCHEMA_INVALID] BP_Total missing column: ${col}`);
    }
  }

  // Build map of existing players -> row index (1-based, for sheet operations)
  const existingPlayers = {};
  for (let i = 1; i < existingData.length; i++) {
    const name = String(existingData[i][colIdx['PreferredName']] || '').trim();
    if (name) {
      existingPlayers[name] = i + 1; // 1-based row number
    }
  }

  // Prepare batch updates
  const now = new Date();
  const updates = []; // Array of {row, data} for existing rows
  const newRows = []; // Array of complete row arrays for new players

  for (const name of allNames) {
    // Calculate values
    const att = attendanceMap[name] || 0;
    const flag = flagMap[name] || 0;
    const dice = diceMap[name] || 0;
    const hist = att + flag + dice;
    const red = redeemedMap[name] || 0;
    const curr = Math.max(0, hist - red); // Clamp to 0 minimum

    if (existingPlayers[name]) {
      // Update existing row
      updates.push({
        row: existingPlayers[name],
        name: name,
        att: att,
        flag: flag,
        dice: dice,
        hist: hist,
        red: red,
        curr: curr,
        updated: now
      });
    } else {
      // New player - prepare row
      const newRow = new Array(headers.length).fill('');
      newRow[colIdx['PreferredName']] = name;
      newRow[colIdx['BP_Current']] = curr;
      newRow[colIdx['Attendance Mission Points']] = att;
      newRow[colIdx['Flag Mission Points']] = flag;
      newRow[colIdx['Dice Roll Points']] = dice;
      newRow[colIdx['LastUpdated']] = now;
      newRow[colIdx['BP_Historical']] = hist;
      newRow[colIdx['BP_Redeemed']] = red;
      newRows.push(newRow);
    }
  }

  // Batch update existing rows
  for (const upd of updates) {
    const rowData = [];
    // Build array in column order for columns 1-8
    for (let c = 0; c < 8; c++) {
      switch (c) {
        case colIdx['PreferredName']: rowData.push(upd.name); break;
        case colIdx['BP_Current']: rowData.push(upd.curr); break;
        case colIdx['Attendance Mission Points']: rowData.push(upd.att); break;
        case colIdx['Flag Mission Points']: rowData.push(upd.flag); break;
        case colIdx['Dice Roll Points']: rowData.push(upd.dice); break;
        case colIdx['LastUpdated']: rowData.push(upd.updated); break;
        case colIdx['BP_Historical']: rowData.push(upd.hist); break;
        case colIdx['BP_Redeemed']: rowData.push(upd.red); break;
        default: rowData.push(''); break;
      }
    }
    bpSheet.getRange(upd.row, 1, 1, 8).setValues([rowData]);
  }

  // Append new rows
  if (newRows.length > 0) {
    const startRow = bpSheet.getLastRow() + 1;
    bpSheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Log the sync
  logIntegrityAction('BP_SYNC', {
    details: `Updated ${updates.length} existing, added ${newRows.length} new players`,
    status: 'SUCCESS'
  });

  return allNames.length;
}

// ============================================================================
// REDEMPTION LOGGING
// ============================================================================

/**
 * Logs a BP redemption into BP_Redeemed_Log.
 * Called by the Redeem BP UI to record point spending.
 *
 * @param {Object} payload
 *   - preferredName: string (required) - Player name
 *   - bpAmount: number (required, positive) - Points being redeemed
 *   - reason: string (optional) - e.g., "Prize Pack", "Lockbox Promo"
 *   - category: string (optional) - e.g., "Prize", "Adjustment"
 *   - eventId: string (optional) - Event sheet name or ID
 * @return {Object} { success: boolean, newBalance?: number, error?: string }
 */
function logBPRedeemTransaction(payload) {
  try {
    // Validate required fields
    if (!payload || !payload.preferredName) {
      return { success: false, error: 'Missing preferredName' };
    }

    const preferredName = String(payload.preferredName).trim();
    const bpAmount = coerceNumber(payload.bpAmount, 0);

    if (!preferredName) {
      return { success: false, error: 'preferredName cannot be empty' };
    }

    if (bpAmount <= 0) {
      return { success: false, error: 'bpAmount must be positive' };
    }

    // Check current balance before redeeming
    const currentBalance = getPlayerBP(preferredName);
    if (currentBalance < bpAmount) {
      return {
        success: false,
        error: `Insufficient BP. Player has ${currentBalance} BP, tried to redeem ${bpAmount}`
      };
    }

    // Ensure BP_Redeemed_Log exists
    const logSheet = ensureBPRedeemedLogSchema_();

    // Prepare row data
    const rowData = [
      new Date(),                                           // Timestamp
      preferredName,                                        // PreferredName
      bpAmount,                                             // BP_Amount
      payload.reason || '',                                 // Reason
      payload.category || '',                               // Category
      payload.eventId || '',                                // Event_ID
      Session.getActiveUser().getEmail() || 'Unknown',      // Staff
      Utilities.getUuid()                                   // RowId
    ];

    // Append to log
    logSheet.appendRow(rowData);

    // Sync BP_Total to reflect new redemption
    updateBPTotalFromSources();

    // Get new balance after sync
    const newBalance = getPlayerBP(preferredName);

    // Log to integrity
    logIntegrityAction('BP_REDEEM_LOG', {
      preferredName: preferredName,
      details: `Redeemed ${bpAmount} BP. Balance: ${currentBalance} → ${newBalance}. Reason: ${payload.reason || 'N/A'}`,
      status: 'SUCCESS'
    });

    return {
      success: true,
      newBalance: newBalance,
      previousBalance: currentBalance,
      redeemed: bpAmount
    };

  } catch (e) {
    console.error('logBPRedeemTransaction error:', e);
    return { success: false, error: e.message };
  }
}

/**
 * Gets a player's current BP balance from BP_Total.
 * Wrapper that uses the existing getPlayerBP from bpService or provides fallback.
 *
 * @param {string} preferredName - Player name
 * @return {number} Current BP balance (0 if not found)
 */
function getBPBalance(preferredName) {
  // Use existing function if available, otherwise implement inline
  if (typeof getPlayerBP === 'function') {
    return getPlayerBP(preferredName);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('BP_Current');

  if (nameCol === -1 || bpCol === -1) {
    return 0;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === String(preferredName).trim()) {
      return coerceNumber(data[i][bpCol], 0);
    }
  }

  return 0;
}

/**
 * Gets detailed BP breakdown for a player.
 *
 * @param {string} preferredName - Player name
 * @return {Object} { current, attendance, flag, dice, historical, redeemed }
 */
function getPlayerBPBreakdown(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  const defaults = {
    current: 0,
    attendance: 0,
    flag: 0,
    dice: 0,
    historical: 0,
    redeemed: 0,
    lastUpdated: null
  };

  if (!sheet || sheet.getLastRow() <= 1) {
    return defaults;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Build column index map
  const colIdx = {};
  for (let i = 0; i < headers.length; i++) {
    colIdx[headers[i]] = i;
  }

  // Find player row
  const nameCol = colIdx['PreferredName'];
  if (nameCol === undefined) {
    return defaults;
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === String(preferredName).trim()) {
      return {
        current: coerceNumber(data[i][colIdx['BP_Current']], 0),
        attendance: coerceNumber(data[i][colIdx['Attendance Mission Points']], 0),
        flag: coerceNumber(data[i][colIdx['Flag Mission Points']], 0),
        dice: coerceNumber(data[i][colIdx['Dice Roll Points']], 0),
        historical: coerceNumber(data[i][colIdx['BP_Historical']], 0),
        redeemed: coerceNumber(data[i][colIdx['BP_Redeemed']], 0),
        lastUpdated: data[i][colIdx['LastUpdated']] || null
      };
    }
  }

  return defaults;
}

/**
 * Gets redemption history for a player from BP_Redeemed_Log.
 *
 * @param {string} preferredName - Player name
 * @return {Array<Object>} Array of redemption records
 */
function getPlayerRedemptionHistory(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Redeemed_Log');

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Build column index map
  const colIdx = {};
  for (let i = 0; i < headers.length; i++) {
    colIdx[headers[i]] = i;
  }

  const nameCol = colIdx['PreferredName'];
  if (nameCol === undefined) {
    return [];
  }

  const history = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim() === String(preferredName).trim()) {
      history.push({
        timestamp: data[i][colIdx['Timestamp']],
        amount: coerceNumber(data[i][colIdx['BP_Amount']], 0),
        reason: data[i][colIdx['Reason']] || '',
        category: data[i][colIdx['Category']] || '',
        eventId: data[i][colIdx['Event_ID']] || '',
        staff: data[i][colIdx['Staff']] || '',
        rowId: data[i][colIdx['RowId']] || ''
      });
    }
  }

  // Sort by timestamp descending (most recent first)
  history.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  return history;
}

// ============================================================================
// MANUAL ADJUSTMENT
// ============================================================================

/**
 * Manually adjusts BP_Historical for a player (for corrections).
 * This bypasses the normal source aggregation and directly sets values.
 * Use with caution - prefer using the mission sheets as sources.
 *
 * @param {string} preferredName - Player name
 * @param {number} adjustment - Amount to add (can be negative for corrections)
 * @param {string} reason - Reason for adjustment
 * @return {Object} { success, newHistorical, newCurrent }
 */
function adjustBPHistorical(preferredName, adjustment, reason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('BP_Total');

    if (!sheet) {
      return { success: false, error: 'BP_Total not found' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Find columns
    const colIdx = {};
    for (let i = 0; i < headers.length; i++) {
      colIdx[headers[i]] = i;
    }

    // Find player row
    const nameCol = colIdx['PreferredName'];
    let playerRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nameCol]).trim() === String(preferredName).trim()) {
        playerRow = i + 1; // 1-based
        break;
      }
    }

    if (playerRow === -1) {
      return { success: false, error: 'Player not found in BP_Total' };
    }

    // Get current values
    const currentHist = coerceNumber(data[playerRow - 1][colIdx['BP_Historical']], 0);
    const currentRed = coerceNumber(data[playerRow - 1][colIdx['BP_Redeemed']], 0);

    // Calculate new values
    const newHist = currentHist + adjustment;
    const newCurr = Math.max(0, newHist - currentRed);

    // Update sheet
    sheet.getRange(playerRow, colIdx['BP_Historical'] + 1).setValue(newHist);
    sheet.getRange(playerRow, colIdx['BP_Current'] + 1).setValue(newCurr);
    sheet.getRange(playerRow, colIdx['LastUpdated'] + 1).setValue(new Date());

    // Log adjustment
    logIntegrityAction('BP_ADJUST', {
      preferredName: preferredName,
      details: `Manual adjustment: ${adjustment} BP. Historical: ${currentHist} → ${newHist}. Current: ${newCurr}. Reason: ${reason}`,
      status: 'SUCCESS'
    });

    return {
      success: true,
      previousHistorical: currentHist,
      newHistorical: newHist,
      newCurrent: newCurr
    };

  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================================
// UTILITY / DEBUG
// ============================================================================

/**
 * Forces a full BP_Total rebuild from all sources.
 * Useful for manual recalculation or debugging.
 *
 * @return {Object} { success, playersUpdated }
 */
function forceBPTotalRebuild() {
  try {
    const count = updateBPTotalFromSources();
    return { success: true, playersUpdated: count };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Gets summary statistics for BP_Total.
 *
 * @return {Object} Summary stats
 */
function getBPTotalStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet || sheet.getLastRow() <= 1) {
    return { playerCount: 0, totalCurrent: 0, totalHistorical: 0, totalRedeemed: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colIdx = {};
  for (let i = 0; i < headers.length; i++) {
    colIdx[headers[i]] = i;
  }

  let totalCurrent = 0;
  let totalHistorical = 0;
  let totalRedeemed = 0;
  let playerCount = 0;

  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][colIdx['PreferredName']] || '').trim();
    if (name) {
      playerCount++;
      totalCurrent += coerceNumber(data[i][colIdx['BP_Current']], 0);
      totalHistorical += coerceNumber(data[i][colIdx['BP_Historical']], 0);
      totalRedeemed += coerceNumber(data[i][colIdx['BP_Redeemed']], 0);
    }
  }

  return {
    playerCount: playerCount,
    totalCurrent: totalCurrent,
    totalHistorical: totalHistorical,
    totalRedeemed: totalRedeemed,
    avgCurrent: playerCount > 0 ? Math.round(totalCurrent / playerCount * 10) / 10 : 0
  };
}