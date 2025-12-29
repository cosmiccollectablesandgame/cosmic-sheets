/**
 * BP Total Pipeline Service v7.9.7
 * @fileoverview CANONICAL pipeline for synchronizing BP_Total from source sheets:
 *   - Attendance_Missions (Attendance Mission Points)
 *   - Flag_Missions (Flag Mission Points)
 *   - Dice Roll Points (Dice Roll Points)
 *
 * AUTHORITATIVE BP_Total SCHEMA:
 *   PreferredName | Current_BP | Attendance Mission Points | Flag Mission Points | Dice Roll Points | LastUpdated | BP_Historical
 *
 * This is the ONLY file that should write to the mission columns in BP_Total.
 */

// ============================================================================
// AUTHORITATIVE SCHEMA CONSTANTS
// ============================================================================

/**
 * Required headers for BP_Total (authoritative schema)
 * @const {Array<string>}
 */
const BP_TOTAL_REQUIRED_HEADERS = [
  'PreferredName',
  'Current_BP',
  'Attendance Mission Points',
  'Flag Mission Points',
  'Dice Roll Points',
  'LastUpdated',
  'BP_Historical'
];

// ============================================================================
// MAIN PIPELINE: updateBPTotalFromSources
// ============================================================================

/**
 * Synchronizes BP_Total from the three mission source sheets.
 * This is the CANONICAL entry point for BP sync operations.
 *
 * Pipeline:
 *   1. Reads mission totals from Attendance_Missions, Flag_Missions, Dice Roll Points
 *   2. Writes them into BP_Total mission columns
 *   3. Computes totalFromSources = att + flag + dice
 *   4. Applies BP_Global_Cap from Prize_Throttle to clamp Current_BP
 *   5. Updates BP_Historical as monotonic lifetime max
 *   6. Updates LastUpdated timestamp
 *
 * @return {number} Count of players updated
 */
function updateBPTotalFromSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure BP_Total has correct schema
  ensureBPTotalSchemaEnhanced_();

  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) {
    console.error('BP_Total sheet not found after schema ensure');
    return 0;
  }

  // Get throttle cap
  const globalCap = getBPGlobalCap_();

  // Read source data
  const attendanceMap = getAttendanceMissionPoints_();
  const flagMap = getFlagMissionPoints_();
  const diceMap = getDiceRollPoints_();

  // Build set of all player names across all sources
  const allNames = new Set([
    ...attendanceMap.keys(),
    ...flagMap.keys(),
    ...diceMap.keys()
  ]);

  if (allNames.size === 0) {
    console.log('No players found in source sheets');
    return 0;
  }

  // Read BP_Total data
  const bpData = bpSheet.getDataRange().getValues();
  const headers = bpData[0];

  // Build column map
  const colMap = mapColumns_(headers, BP_TOTAL_REQUIRED_HEADERS);

  // Validate we have all required columns
  const missingCols = BP_TOTAL_REQUIRED_HEADERS.filter(h => colMap[h] === undefined);
  if (missingCols.length > 0) {
    console.error('BP_Total missing columns:', missingCols.join(', '));
    return 0;
  }

  // Build map of existing BP_Total rows by PreferredName
  const existingRows = new Map();
  for (let i = 1; i < bpData.length; i++) {
    const name = String(bpData[i][colMap['PreferredName']] || '').trim();
    if (name) {
      existingRows.set(name, i); // Row index in data array (0-based)
    }
  }

  const now = new Date();
  let updatedCount = 0;
  const newRows = [];

  // Process all players
  for (const playerName of allNames) {
    const attPoints = attendanceMap.get(playerName) || 0;
    const flagPts = flagMap.get(playerName) || 0;
    const dicePts = diceMap.get(playerName) || 0;

    const totalFromSources = attPoints + flagPts + dicePts;
    const newCurrentBP = Math.min(totalFromSources, globalCap);

    if (existingRows.has(playerName)) {
      // Update existing row
      const rowIdx = existingRows.get(playerName);
      const rowNum = rowIdx + 1; // 1-based for sheet

      // Get current values
      const oldAttendance = coerceNumber(bpData[rowIdx][colMap['Attendance Mission Points']], 0);
      const oldFlag = coerceNumber(bpData[rowIdx][colMap['Flag Mission Points']], 0);
      const oldDice = coerceNumber(bpData[rowIdx][colMap['Dice Roll Points']], 0);
      const currentHistorical = coerceNumber(bpData[rowIdx][colMap['BP_Historical']], 0);

      // Check if anything changed
      const hasChanges = (
        oldAttendance !== attPoints ||
        oldFlag !== flagPts ||
        oldDice !== dicePts
      );

      if (hasChanges) {
        // Update BP_Historical as monotonic max
        const newHistorical = Math.max(currentHistorical, totalFromSources);

        // Write updates
        bpSheet.getRange(rowNum, colMap['Attendance Mission Points'] + 1).setValue(attPoints);
        bpSheet.getRange(rowNum, colMap['Flag Mission Points'] + 1).setValue(flagPts);
        bpSheet.getRange(rowNum, colMap['Dice Roll Points'] + 1).setValue(dicePts);
        bpSheet.getRange(rowNum, colMap['Current_BP'] + 1).setValue(newCurrentBP);
        bpSheet.getRange(rowNum, colMap['BP_Historical'] + 1).setValue(newHistorical);
        bpSheet.getRange(rowNum, colMap['LastUpdated'] + 1).setValue(now);

        updatedCount++;
      }
    } else {
      // New player - prepare row for batch append
      const newRow = new Array(headers.length).fill('');
      newRow[colMap['PreferredName']] = playerName;
      newRow[colMap['Current_BP']] = newCurrentBP;
      newRow[colMap['Attendance Mission Points']] = attPoints;
      newRow[colMap['Flag Mission Points']] = flagPts;
      newRow[colMap['Dice Roll Points']] = dicePts;
      newRow[colMap['LastUpdated']] = now;
      newRow[colMap['BP_Historical']] = totalFromSources;

      newRows.push(newRow);
      updatedCount++;
    }
  }

  // Batch append new rows
  if (newRows.length > 0) {
    const startRow = bpSheet.getLastRow() + 1;
    bpSheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Log the sync
  if (typeof logIntegrityAction === 'function') {
    logIntegrityAction('BP_TOTAL_SYNC', {
      details: `Synced ${updatedCount} player(s) from sources. Cap: ${globalCap}`,
      status: 'SUCCESS'
    });
  }

  return updatedCount;
}

// ============================================================================
// SOURCE SHEET READERS
// ============================================================================

/**
 * Reads Attendance_Missions and returns Map of PreferredName -> Attendance Mission Points
 * @return {Map<string, number>}
 * @private
 */
function getAttendanceMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find name column
  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) {
    console.warn('Attendance_Missions: PreferredName column not found');
    return new Map();
  }

  // Find points column - try multiple names for compatibility
  const pointsColNames = [
    'Attendance Mission Points',
    'Attendance Points',
    'Points',
    'Total Points'
  ];

  let pointsCol = -1;
  for (const colName of pointsColNames) {
    pointsCol = headers.indexOf(colName);
    if (pointsCol !== -1) break;
  }

  if (pointsCol === -1) {
    console.warn('Attendance_Missions: Points column not found');
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

/**
 * Reads Flag_Missions and returns Map of PreferredName -> Flag Mission Points
 * @return {Map<string, number>}
 * @private
 */
function getFlagMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find name column
  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) {
    console.warn('Flag_Missions: PreferredName column not found');
    return new Map();
  }

  // Find points column - try multiple names for compatibility
  const pointsColNames = [
    'Flag Mission Points',
    'Flag Points',
    'Points',
    'Total Points'
  ];

  let pointsCol = -1;
  for (const colName of pointsColNames) {
    pointsCol = headers.indexOf(colName);
    if (pointsCol !== -1) break;
  }

  if (pointsCol === -1) {
    console.warn('Flag_Missions: Points column not found');
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

/**
 * Reads Dice Roll Points (or legacy Dice_Points) and returns Map of PreferredName -> Dice Roll Points
 * @return {Map<string, number>}
 * @private
 */
function getDiceRollPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Try primary sheet name first, then legacy fallback
  let sheet = ss.getSheetByName('Dice Roll Points');
  if (!sheet) {
    sheet = ss.getSheetByName('Dice_Points');
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return new Map();
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find name column
  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) {
    console.warn('Dice Roll Points: PreferredName column not found');
    return new Map();
  }

  // Find points column - try multiple names for compatibility
  const pointsColNames = [
    'Dice Roll Points',
    'Dice Points',
    'Points',
    'Total Points'
  ];

  let pointsCol = -1;
  for (const colName of pointsColNames) {
    pointsCol = headers.indexOf(colName);
    if (pointsCol !== -1) break;
  }

  if (pointsCol === -1) {
    console.warn('Dice Roll Points: Points column not found');
    return new Map();
  }

  const result = new Map();
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      const points = coerceNumber(data[i][pointsCol], 0);
      result.set(name, points);
    }
  }

  return result;
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures BP_Total has the authoritative schema with all required headers.
 * Creates the sheet if missing, or adds missing columns to existing sheet.
 * Does NOT remove or rename existing columns.
 * @private
 */
function ensureBPTotalSchemaEnhanced_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    // Create new sheet with full schema
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(BP_TOTAL_REQUIRED_HEADERS);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_TOTAL_REQUIRED_HEADERS.length);
    return;
  }

  if (sheet.getLastRow() === 0) {
    // Empty sheet - add headers
    sheet.appendRow(BP_TOTAL_REQUIRED_HEADERS);
    sheet.setFrozenRows(1);
    formatHeaderRow_(sheet, BP_TOTAL_REQUIRED_HEADERS.length);
    return;
  }

  // Check existing headers and add missing ones
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = BP_TOTAL_REQUIRED_HEADERS.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });

    console.log('Added missing BP_Total columns:', missing.join(', '));
  }
}

/**
 * Maps column names to their 0-based indices
 * @param {Array<string>} headers - Header row
 * @param {Array<string>} requiredHeaders - Headers to map
 * @return {Object} Map of header name -> column index
 * @private
 */
function mapColumns_(headers, requiredHeaders) {
  const colMap = {};
  requiredHeaders.forEach(header => {
    const idx = headers.indexOf(header);
    if (idx !== -1) {
      colMap[header] = idx;
    }
  });
  return colMap;
}

/**
 * Formats the header row with styling
 * @param {Sheet} sheet - The sheet
 * @param {number} numCols - Number of columns
 * @private
 */
function formatHeaderRow_(sheet, numCols) {
  sheet.getRange(1, 1, 1, numCols)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
}

/**
 * Gets BP_Global_Cap from Prize_Throttle (default: 100)
 * @return {number}
 * @private
 */
function getBPGlobalCap_() {
  try {
    if (typeof getThrottleParam === 'function') {
      const cap = getThrottleParam('BP_Global_Cap', 100);
      return coerceNumber(cap, 100);
    }

    if (typeof getThrottleKV === 'function') {
      const throttle = getThrottleKV();
      return coerceNumber(throttle.BP_Global_Cap, 100);
    }
  } catch (e) {
    console.warn('Could not get BP_Global_Cap from throttle:', e.message);
  }

  return 100; // Default cap
}

// ============================================================================
// LEGACY FUNCTION STUBS (DEPRECATED)
// These are kept for backward compatibility but redirect to the canonical pipeline
// ============================================================================

/**
 * @deprecated Use updateBPTotalFromSources() instead
 * Legacy function - redirects to canonical pipeline
 */
function syncBPTotals(silent) {
  console.warn('syncBPTotals is deprecated. Use updateBPTotalFromSources instead.');
  const count = updateBPTotalFromSources();
  return {
    synced: count,
    errors: []
  };
}

/**
 * @deprecated Use ensureBPTotalSchemaEnhanced_() instead
 * Legacy consolidated schema function - now a no-op
 */
function ensureBPTotalConsolidatedSchema() {
  console.warn('ensureBPTotalConsolidatedSchema is deprecated. Using canonical schema.');
  ensureBPTotalSchemaEnhanced_();
}

/**
 * @deprecated No longer needed
 * Legacy migration function - now a no-op
 */
function migrateBPTotalSchema_() {
  console.warn('migrateBPTotalSchema_ is deprecated. Schema is now managed by ensureBPTotalSchemaEnhanced_.');
  ensureBPTotalSchemaEnhanced_();
}

// ============================================================================
// VALIDATION HELPERS
// ============================================================================

/**
 * Validates mission points integrity across source sheets and BP_Total
 * @return {Object} {pass: boolean, issues: Array}
 */
function validateMissionPointsIntegrity() {
  const issues = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check source sheets exist
  const sourceSheets = [
    { name: 'Attendance_Missions', pointsCol: 'Attendance Mission Points' },
    { name: 'Flag_Missions', pointsCol: 'Flag Mission Points' },
    { name: 'Dice Roll Points', pointsCol: 'Dice Roll Points' }
  ];

  sourceSheets.forEach(spec => {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      issues.push({
        sheet: spec.name,
        row: 0,
        issue: 'Sheet not found'
      });
      return;
    }

    if (sheet.getLastRow() <= 1) {
      // Empty sheet is OK
      return;
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    if (!headers.includes('PreferredName')) {
      issues.push({
        sheet: spec.name,
        row: 1,
        issue: 'Missing PreferredName column'
      });
    }

    // Check for points column (allowing synonyms)
    const hasPointsCol = headers.some(h =>
      h === spec.pointsCol ||
      h.toLowerCase().includes('points')
    );

    if (!hasPointsCol) {
      issues.push({
        sheet: spec.name,
        row: 1,
        issue: `Missing ${spec.pointsCol} column`
      });
    }
  });

  // Check BP_Total schema
  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) {
    issues.push({
      sheet: 'BP_Total',
      row: 0,
      issue: 'Sheet not found'
    });
  } else if (bpSheet.getLastRow() > 0) {
    const headers = bpSheet.getRange(1, 1, 1, bpSheet.getLastColumn()).getValues()[0];

    BP_TOTAL_REQUIRED_HEADERS.forEach(required => {
      if (!headers.includes(required)) {
        issues.push({
          sheet: 'BP_Total',
          row: 1,
          issue: `Missing column: ${required}`
        });
      }
    });
  }

  return {
    pass: issues.length === 0,
    issues: issues
  };
}