/**
 * BP Sync Service - Synchronize BP_Total from External Sources
 * @fileoverview Keeps BP_Total in sync with:
 *  - Attendance_Missions
 *  - Flag_Missions
 *  - Dice Roll Points
 *
 * Tab Headers (as provided):
 *
 * Attendance_Missions:
 *   PreferredName | Attendance Mission Points | First Contact | ... | Black Hole Survivor
 *
 * BP_Total (extended):
 *   PreferredName | Current_BP | Attendance Mission Points | Flag Mission Points
 *   | Dice Roll Points | Historical_BP | LastUpdated
 *
 * Dice Roll Points:
 *   PreferredName | Points | LastUpdated
 *
 * Flag_Missions:
 *   PreferredName | Flag Mission Points | Cosmic_Selfie | ... | Quantum_Collector | LastUpdated
 */

// ============================================================================
// SYNC CONFIG
// ============================================================================

const BP_SYNC_CONFIG = {
  BP_TOTAL_SHEET_NAME: 'BP_Total',
  ATTENDANCE_MISSIONS_SHEET_NAME: 'Attendance_Missions',
  FLAG_MISSIONS_SHEET_NAME: 'Flag_Missions',
  DICE_POINTS_SHEET_NAME: 'Dice Roll Points',
  BP_CURRENT_CAP: 100 // cap Current_BP at 0-100
};

// ============================================================================
// PUBLIC ENTRYPOINTS
// ============================================================================

/**
 * Full refresh: recomputes BP_Total from all three source sheets.
 * Attach this to a menu item or a "Refresh BP" button.
 */
function refreshBP_Total() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bpSheet = ss.getSheetByName(BP_SYNC_CONFIG.BP_TOTAL_SHEET_NAME);
  if (!bpSheet) {
    throw new Error('BP_Total sheet not found.');
  }

  const bpRange = bpSheet.getDataRange();
  const bpValues = bpRange.getValues();
  if (bpValues.length < 2) {
    // header only or empty
    return;
  }

  const bpHeader = bpValues[0];
  const bpHeaderMap = mapSyncHeaders_(bpHeader);

  const idxBP_PreferredName = bpHeaderMap['PreferredName'];
  const idxBP_CurrentBP = bpHeaderMap['Current_BP'];
  const idxBP_Attendance = bpHeaderMap['Attendance Mission Points'];
  const idxBP_Flag = bpHeaderMap['Flag Mission Points'];
  const idxBP_Dice = bpHeaderMap['Dice Roll Points'];
  const idxBP_Historical = bpHeaderMap['Historical_BP'];
  const idxBP_LastUpdated = bpHeaderMap['LastUpdated'];

  // Build lookup maps from each source sheet: { PreferredName -> points }
  const attendanceMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.ATTENDANCE_MISSIONS_SHEET_NAME,
    'PreferredName',
    'Attendance Mission Points'
  );

  const flagMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.FLAG_MISSIONS_SHEET_NAME,
    'PreferredName',
    'Flag Mission Points'
  );

  const diceMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.DICE_POINTS_SHEET_NAME,
    'PreferredName',
    'Points' // source column name
  );

  const now = new Date();
  let anyChanged = false;
  let playersUpdated = 0;

  // Work on a copy of the data (skip header row at index 0)
  for (let r = 1; r < bpValues.length; r++) {
    const row = bpValues[r];
    const preferredName = String(row[idxBP_PreferredName] || '').trim();
    if (!preferredName) continue;

    const attPoints = attendanceMap[preferredName] || 0;
    const flagPoints = flagMap[preferredName] || 0;
    const dicePoints = diceMap[preferredName] || 0;

    const newEarnedTotal = attPoints + flagPoints + dicePoints;

    const oldAttendance = toSyncNumber_(row[idxBP_Attendance]);
    const oldFlag = toSyncNumber_(row[idxBP_Flag]);
    const oldDice = toSyncNumber_(row[idxBP_Dice]);
    const oldHistorical = toSyncNumber_(row[idxBP_Historical]);
    const oldCurrent = toSyncNumber_(row[idxBP_CurrentBP]);

    // Update the breakdown columns directly from the source maps
    row[idxBP_Attendance] = attPoints;
    row[idxBP_Flag] = flagPoints;
    row[idxBP_Dice] = dicePoints;

    // Historical_BP logic:
    // - If source totals increased, bump Historical_BP by the delta.
    // - If they decreased (correction), reset Historical_BP to the new total,
    //   but do NOT auto-reduce Current_BP.
    let newHistorical = oldHistorical;
    let deltaEarned = 0;

    if (newEarnedTotal > oldHistorical) {
      newHistorical = newEarnedTotal;
      deltaEarned = newEarnedTotal - oldHistorical;
    } else if (newEarnedTotal < oldHistorical) {
      newHistorical = newEarnedTotal;
      deltaEarned = 0;
    }

    // Current_BP: add only the positive delta, clamp to [0, BP_CURRENT_CAP]
    let newCurrent = oldCurrent + deltaEarned;
    if (newCurrent < 0) newCurrent = 0;
    if (newCurrent > BP_SYNC_CONFIG.BP_CURRENT_CAP) newCurrent = BP_SYNC_CONFIG.BP_CURRENT_CAP;

    row[idxBP_Historical] = newHistorical;
    row[idxBP_CurrentBP] = newCurrent;

    // If anything in this row changed, stamp LastUpdated
    const rowChanged =
      oldAttendance !== attPoints ||
      oldFlag !== flagPoints ||
      oldDice !== dicePoints ||
      oldHistorical !== newHistorical ||
      oldCurrent !== newCurrent;

    if (rowChanged) {
      row[idxBP_LastUpdated] = now;
      anyChanged = true;
      playersUpdated++;
    }
  }

  if (anyChanged) {
    // Write back everything except the header row
    const bodyRange = bpSheet.getRange(
      2, // row 2 (1-based)
      1, // col 1
      bpValues.length - 1,
      bpValues[0].length
    );
    bodyRange.setValues(bpValues.slice(1));

    logIntegrityAction('BP_SYNC', {
      details: `Full BP sync completed. Updated ${playersUpdated} player(s).`,
      status: 'SUCCESS'
    });
  }
}

/**
 * Lightweight single-player refresh, useful for onEdit hooks.
 * @param {string} preferredName - Player's preferred name
 */
function refreshBP_ForPlayer_(preferredName) {
  if (!preferredName) return;
  preferredName = String(preferredName).trim();
  if (!preferredName) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpSheet = ss.getSheetByName(BP_SYNC_CONFIG.BP_TOTAL_SHEET_NAME);
  if (!bpSheet) {
    throw new Error('BP_Total sheet not found.');
  }

  const bpRange = bpSheet.getDataRange();
  const bpValues = bpRange.getValues();
  if (bpValues.length < 2) return;

  const bpHeader = bpValues[0];
  const bpHeaderMap = mapSyncHeaders_(bpHeader);

  const idxBP_PreferredName = bpHeaderMap['PreferredName'];
  const idxBP_CurrentBP = bpHeaderMap['Current_BP'];
  const idxBP_Attendance = bpHeaderMap['Attendance Mission Points'];
  const idxBP_Flag = bpHeaderMap['Flag Mission Points'];
  const idxBP_Dice = bpHeaderMap['Dice Roll Points'];
  const idxBP_Historical = bpHeaderMap['Historical_BP'];
  const idxBP_LastUpdated = bpHeaderMap['LastUpdated'];

  // Build source maps once
  const attendanceMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.ATTENDANCE_MISSIONS_SHEET_NAME,
    'PreferredName',
    'Attendance Mission Points'
  );
  const flagMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.FLAG_MISSIONS_SHEET_NAME,
    'PreferredName',
    'Flag Mission Points'
  );
  const diceMap = buildPointsMap_(
    ss,
    BP_SYNC_CONFIG.DICE_POINTS_SHEET_NAME,
    'PreferredName',
    'Points'
  );

  const attPoints = attendanceMap[preferredName] || 0;
  const flagPoints = flagMap[preferredName] || 0;
  const dicePoints = diceMap[preferredName] || 0;
  const newEarnedTotal = attPoints + flagPoints + dicePoints;

  const now = new Date();

  let targetRowIndex = -1;
  for (let r = 1; r < bpValues.length; r++) {
    const name = String(bpValues[r][idxBP_PreferredName] || '').trim();
    if (name === preferredName) {
      targetRowIndex = r;
      break;
    }
  }
  if (targetRowIndex === -1) {
    // Player not found in BP_Total; you can optionally auto-add them here.
    return;
  }

  const row = bpValues[targetRowIndex];

  const oldAttendance = toSyncNumber_(row[idxBP_Attendance]);
  const oldFlag = toSyncNumber_(row[idxBP_Flag]);
  const oldDice = toSyncNumber_(row[idxBP_Dice]);
  const oldHistorical = toSyncNumber_(row[idxBP_Historical]);
  const oldCurrent = toSyncNumber_(row[idxBP_CurrentBP]);

  row[idxBP_Attendance] = attPoints;
  row[idxBP_Flag] = flagPoints;
  row[idxBP_Dice] = dicePoints;

  let newHistorical = oldHistorical;
  let deltaEarned = 0;

  if (newEarnedTotal > oldHistorical) {
    newHistorical = newEarnedTotal;
    deltaEarned = newEarnedTotal - oldHistorical;
  } else if (newEarnedTotal < oldHistorical) {
    newHistorical = newEarnedTotal;
    deltaEarned = 0;
  }

  let newCurrent = oldCurrent + deltaEarned;
  if (newCurrent < 0) newCurrent = 0;
  if (newCurrent > BP_SYNC_CONFIG.BP_CURRENT_CAP) newCurrent = BP_SYNC_CONFIG.BP_CURRENT_CAP;

  row[idxBP_Historical] = newHistorical;
  row[idxBP_CurrentBP] = newCurrent;

  const rowChanged =
    oldAttendance !== attPoints ||
    oldFlag !== flagPoints ||
    oldDice !== oldDice ||
    oldHistorical !== newHistorical ||
    oldCurrent !== newCurrent;

  if (rowChanged) {
    row[idxBP_LastUpdated] = now;
    // Write just this one row back
    bpSheet
      .getRange(targetRowIndex + 1, 1, 1, bpValues[0].length) // +1 because sheet is 1-based
      .setValues([row]);

    logIntegrityAction('BP_SYNC_PLAYER', {
      preferredName,
      details: `Player BP synced. Current: ${oldCurrent} → ${newCurrent}, Historical: ${oldHistorical} → ${newHistorical}`,
      status: 'SUCCESS'
    });
  }
}

// ============================================================================
// ON EDIT TRIGGER
// ============================================================================

/**
 * Simple trigger - keeps BP_Total in sync when source sheets change.
 *
 * Watches:
 *  - Attendance_Missions
 *  - Flag_Missions
 *  - Dice Roll Points
 *
 * On any edit to those sheets (not header row), it:
 *  - Reads the PreferredName in that row
 *  - Calls refreshBP_ForPlayer_(PreferredName)
 *
 * @param {Object} e - Edit event object
 */
function onEditBPSync_(e) {
  try {
    if (!e) return;

    const range = e.range;
    const sheet = range.getSheet();
    const sheetName = sheet.getName();

    const watchedSheets = [
      BP_SYNC_CONFIG.ATTENDANCE_MISSIONS_SHEET_NAME,
      BP_SYNC_CONFIG.FLAG_MISSIONS_SHEET_NAME,
      BP_SYNC_CONFIG.DICE_POINTS_SHEET_NAME
    ];

    // Ignore sheets we don't care about
    if (watchedSheets.indexOf(sheetName) === -1) {
      return;
    }

    const row = range.getRow();
    if (row === 1) {
      // Header row - ignore
      return;
    }

    // Read header row to find PreferredName column
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = mapSyncHeaders_(headerRow);
    const idxPreferred = headerMap['PreferredName'];

    if (idxPreferred == null) {
      // Sheet doesn't have a PreferredName header - bail
      return;
    }

    // Sheet is 1-based, idxPreferred is 0-based
    const preferredName = String(
      sheet.getRange(row, idxPreferred + 1).getValue() || ''
    ).trim();

    if (!preferredName) {
      // No player on this row - nothing to sync
      return;
    }

    // Update this player's BP_Total row
    refreshBP_ForPlayer_(preferredName);

  } catch (err) {
    // Log error but don't break the edit
    console.error('onEditBPSync_ error:', err);
  }
}

// ============================================================================
// SYNC HELPERS
// ============================================================================

/**
 * Maps header names to their column index (0-based).
 * @param {string[]} headerRow
 * @return {Object<string, number>}
 * @private
 */
function mapSyncHeaders_(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    const name = String(headerRow[i] || '').trim();
    if (name) {
      map[name] = i;
    }
  }
  return map;
}

/**
 * Builds a lookup map from a sheet: { keyValue -> pointsValue }.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} sheetName
 * @param {string} keyHeader - e.g., 'PreferredName'
 * @param {string} pointsHeader - e.g., 'Attendance Mission Points' or 'Points'
 * @return {Object<string, number>}
 * @private
 */
function buildPointsMap_(ss, sheetName, keyHeader, pointsHeader) {
  const map = {};
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return map;

  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return map;

  const header = values[0];
  const headerMap = mapSyncHeaders_(header);

  const idxKey = headerMap[keyHeader];
  const idxPoints = headerMap[pointsHeader];

  if (idxKey == null || idxPoints == null) {
    // silently return empty map if headers missing
    return map;
  }

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const key = String(row[idxKey] || '').trim();
    if (!key) continue;

    const rawPoints = row[idxPoints];
    const points = toSyncNumber_(rawPoints);
    if (isNaN(points) || points === 0) continue;

    map[key] = points;
  }

  return map;
}

/**
 * Safe number conversion.
 * @param {*} value
 * @return {number}
 * @private
 */
function toSyncNumber_(value) {
  if (value === '' || value == null) return 0;
  const n = Number(value);
  return isNaN(n) ? 0 : n;
}

// ============================================================================
// SCHEMA MIGRATION
// ============================================================================

/**
 * Ensures BP_Total has the extended schema for sync operations.
 * Adds missing columns: Current_BP, Attendance Mission Points,
 * Flag Mission Points, Dice Roll Points, Historical_BP
 */
function ensureBPTotalSyncSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BP_SYNC_CONFIG.BP_TOTAL_SHEET_NAME);

  const requiredHeaders = [
    'PreferredName',
    'Current_BP',
    'Attendance Mission Points',
    'Flag Mission Points',
    'Dice Roll Points',
    'Historical_BP',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(BP_SYNC_CONFIG.BP_TOTAL_SHEET_NAME);
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');

    logIntegrityAction('BP_SYNC_SCHEMA', {
      details: 'Created BP_Total with extended sync schema',
      status: 'SUCCESS'
    });
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });

    logIntegrityAction('BP_SYNC_SCHEMA', {
      details: `Added missing columns to BP_Total: ${missing.join(', ')}`,
      status: 'SUCCESS'
    });
  }
}