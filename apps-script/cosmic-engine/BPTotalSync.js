/**
 * BP Total Sync Service
 * @fileoverview Synchronizes BP_Total from source sheets:
 *   - Attendance_Missions (Attendance Mission Points)
 *   - Flag_Missions (Flag Mission Points)
 *   - Dice_Points / Dice Roll Points
 *
 * ASSUMPTION: BP_Total has columns: PreferredName, Current_BP, Attendance Missions,
 *             Flag Missions, Dice Roll Points, Historical_BP, LastUpdated
 * ASSUMPTION: Source sheets have PreferredName as the key column
 */

// ============================================================================
// MAIN SYNC FUNCTION
// ============================================================================

/**
 * Synchronizes BP_Total from all source sheets
 * Called by onEdit when source sheets are modified, and from menu
 */
function updateBPTotalFromSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ensure BP_Total schema exists
  ensureBPTotalSchemaEnhanced_();

  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) {
    console.error('BP_Total sheet not found');
    return;
  }

  // Get source data
  const attendancePoints = getAttendanceMissionPoints_();
  const flagPoints = getFlagMissionPoints_();
  const dicePoints = getDiceRollPoints_();

  // Get BP_Total data
  const bpData = bpSheet.getDataRange().getValues();
  if (bpData.length <= 1) {
    // Only header, nothing to sync
    return;
  }

  const headers = bpData[0];
  const colMap = mapColumns_(headers, [
    'PreferredName',
    'Current_BP',
    'Attendance Missions',
    'Flag Missions',
    'Dice Roll Points',
    'Historical_BP',
    'LastUpdated'
  ]);

  // Get throttle settings
  const bpGlobalCap = getThrottleNumber('BP_Global_Cap', 100);

  // Track changes
  let updatedCount = 0;
  const now = new Date();

  // Process each player row
  for (let i = 1; i < bpData.length; i++) {
    const playerName = bpData[i][colMap.PreferredName];
    if (!playerName) continue;

    const attPoints = attendancePoints[playerName] || 0;
    const flagPts = flagPoints[playerName] || 0;
    const dicePts = dicePoints[playerName] || 0;

    // Calculate new current BP
    const totalFromSources = attPoints + flagPts + dicePts;

    // Get historical (lifetime earned) - only increases
    const currentHistorical = coerceNumber(bpData[i][colMap.Historical_BP], 0);
    const newHistorical = Math.max(currentHistorical, totalFromSources);

    // Current BP is clamped to global cap
    const newCurrentBP = Math.min(totalFromSources, bpGlobalCap);

    // Check if anything changed
    const oldCurrentBP = coerceNumber(bpData[i][colMap.Current_BP], 0);
    const oldAttendance = coerceNumber(bpData[i][colMap['Attendance Missions']], 0);
    const oldFlag = coerceNumber(bpData[i][colMap['Flag Missions']], 0);
    const oldDice = coerceNumber(bpData[i][colMap['Dice Roll Points']], 0);

    if (newCurrentBP !== oldCurrentBP ||
        attPoints !== oldAttendance ||
        flagPts !== oldFlag ||
        dicePts !== oldDice) {

      // Update the row
      const rowNum = i + 1;

      if (colMap['Attendance Missions'] !== undefined) {
        bpSheet.getRange(rowNum, colMap['Attendance Missions'] + 1).setValue(attPoints);
      }
      if (colMap['Flag Missions'] !== undefined) {
        bpSheet.getRange(rowNum, colMap['Flag Missions'] + 1).setValue(flagPts);
      }
      if (colMap['Dice Roll Points'] !== undefined) {
        bpSheet.getRange(rowNum, colMap['Dice Roll Points'] + 1).setValue(dicePts);
      }
      if (colMap.Current_BP !== undefined) {
        bpSheet.getRange(rowNum, colMap.Current_BP + 1).setValue(newCurrentBP);
      }
      if (colMap.Historical_BP !== undefined) {
        bpSheet.getRange(rowNum, colMap.Historical_BP + 1).setValue(newHistorical);
      }
      if (colMap.LastUpdated !== undefined) {
        bpSheet.getRange(rowNum, colMap.LastUpdated + 1).setValue(now);
      }

      updatedCount++;
    }
  }

  // Log if changes were made
  if (updatedCount > 0) {
    logIntegrityAction('BP_SYNC', {
      details: `Synced ${updatedCount} player(s) from source sheets`,
      status: 'SUCCESS'
    });
  }

  return updatedCount;
}

// ============================================================================
// SOURCE DATA EXTRACTORS
// ============================================================================

/**
 * Gets attendance mission points by player
 * @return {Object} Map of PreferredName -> points
 * @private
 */
function getAttendanceMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Find PreferredName column
  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) return {};

  // Find points column - could be named various things
  const pointsColNames = ['Attendance Mission Points', 'Attendance Points', 'Points', 'Total Points','Total Events Attended'];
  let pointsCol = -1;
  for (const name of pointsColNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      pointsCol = idx;
      break;
    }
  }

  // If no explicit points column, sum all numeric mission columns
  const result = {};

  for (let i = 1; i < data.length; i++) {
    const playerName = data[i][nameCol];
    if (!playerName) continue;

    if (pointsCol !== -1) {
      result[playerName] = coerceNumber(data[i][pointsCol], 0);
    } else {
      // Sum numeric columns (excluding PreferredName, LastUpdated, etc.)
      let sum = 0;
      for (let j = 0; j < headers.length; j++) {
        const header = String(headers[j]).toLowerCase();
        if (header === 'preferredname' || header === 'lastupdated') continue;
        const val = data[i][j];
        if (typeof val === 'number') {
          sum += val;
        } else if (val === true) {
          sum += 1; // Count checkboxes as 1 point each
        }
      }
      result[playerName] = sum;
    }
  }

  return result;
}

/**
 * Gets flag mission points by player
 * @return {Object} Map of PreferredName -> points
 * @private
 */
function getFlagMissionPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) return {};

  // Look for explicit points column
  const pointsColNames = ['Flag Mission Points', 'Flag Points', 'Points', 'Total Points'];
  let pointsCol = -1;
  for (const name of pointsColNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      pointsCol = idx;
      break;
    }
  }

  const result = {};

  for (let i = 1; i < data.length; i++) {
    const playerName = data[i][nameCol];
    if (!playerName) continue;

    if (pointsCol !== -1) {
      result[playerName] = coerceNumber(data[i][pointsCol], 0);
    } else {
      // Sum completed flag missions
      // Flag missions are typically checkboxes (true/false) or numeric values
      let sum = 0;
      for (let j = 0; j < headers.length; j++) {
        const header = String(headers[j]).toLowerCase();
        if (header === 'preferredname' || header === 'lastupdated') continue;
        const val = data[i][j];
        if (typeof val === 'number') {
          sum += val;
        } else if (val === true) {
          // Award based on mission type - default 1 point per mission
          sum += getFlagMissionValue_(headers[j]);
        }
      }
      result[playerName] = sum;
    }
  }

  return result;
}

/**
 * Gets value for a flag mission by name
 * @param {string} missionName - The mission column name
 * @return {number} Point value
 * @private
 */
function getFlagMissionValue_(missionName) {
  // Standard flag mission point values
  const values = {
    'Cosmic_Selfie': 1,
    'Review_Writer': 2,
    'Social_Media_Star': 2,
    'App_Explorer': 1,
    'Cosmic_Merchant': 3,
    'Precon_Pioneer': 2,
    'Gravitational_Pull': 5,
    'Rogue_Planet': 3,
    'Quantum_Collector': 5
  };

  return values[missionName] || 1;
}

/**
 * Gets dice roll points by player
 * @return {Object} Map of PreferredName -> points
 * @private
 */
function getDiceRollPoints_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Try both possible sheet names
  let sheet = ss.getSheetByName('Dice_Points');
  if (!sheet) {
    sheet = ss.getSheetByName('Dice Roll Points');
  }

  if (!sheet || sheet.getLastRow() <= 1) {
    return {};
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) return {};

  // Look for points column
  const pointsColNames = ['Dice Roll Points', 'Points', 'DicePoints', 'Total'];
  let pointsCol = -1;
  for (const name of pointsColNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      pointsCol = idx;
      break;
    }
  }

  const result = {};

  for (let i = 1; i < data.length; i++) {
    const playerName = data[i][nameCol];
    if (!playerName) continue;

    if (pointsCol !== -1) {
      result[playerName] = coerceNumber(data[i][pointsCol], 0);
    } else {
      // Default to 0 if no points column found
      result[playerName] = 0;
    }
  }

  return result;
}

// ============================================================================
// SCHEMA MANAGEMENT
// ============================================================================

/**
 * Ensures BP_Total has enhanced schema with all source columns
 * @private
 */
function ensureBPTotalSchemaEnhanced_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  const requiredHeaders = [
    'PreferredName',
    'Current_BP',
    'Attendance Missions',
    'Flag Missions',
    'Dice Roll Points',
    'Historical_BP',
    'LastUpdated'
  ];

  if (!sheet) {
    // Create new sheet with proper schema
    sheet = ss.insertSheet('BP_Total');
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, requiredHeaders.length);
    return;
  }

  // Check existing headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.autoResizeColumns(1, requiredHeaders.length);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Add missing columns
  requiredHeaders.forEach(header => {
    if (!existingHeaders.includes(header)) {
      const newCol = existingHeaders.length + 1;
      sheet.getRange(1, newCol).setValue(header);
      existingHeaders.push(header);
    }
  });
}

// ============================================================================
// HELPERS
// ============================================================================

/**
 * Maps column names to indices
 * @param {Array} headers - Header row
 * @param {Array} colNames - Column names to find
 * @return {Object} Map of column name -> index
 * @private
 */
function mapColumns_(headers, colNames) {
  const result = {};
  colNames.forEach(name => {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      result[name] = idx;
    }
  });
  return result;
}

/**
 * Provisions a new player row in BP_Total
 * @param {string} preferredName - The player name
 * @return {boolean} True if row was added
 */
function provisionBPTotalRow(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    ensureBPTotalSchemaEnhanced_();
    return provisionBPTotalRow(preferredName);
  }

  // Check if player already exists
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === preferredName) {
      return false; // Already exists
    }
  }

  // Add new row with zeros
  const headers = data[0];
  const newRow = headers.map(h => {
    if (h === 'PreferredName') return preferredName;
    if (h === 'LastUpdated') return new Date();
    return 0;
  });

  sheet.appendRow(newRow);
  return true;
}

/**
 * Gets a player's current BP from BP_Total
 * @param {string} preferredName - The player name
 * @return {number} Current BP
 */
function getPlayerBP(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');

  if (nameCol === -1 || bpCol === -1) return 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      return coerceNumber(data[i][bpCol], 0);
    }
  }

  return 0;
}

/**
 * Deducts BP from a player (for redemptions)
 * @param {string} preferredName - The player name
 * @param {number} amount - Amount to deduct
 * @return {Object} {success, remaining}
 */
function deductBP(preferredName, amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    return { success: false, remaining: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');

  if (nameCol === -1 || bpCol === -1) {
    return { success: false, remaining: 0 };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const currentBP = coerceNumber(data[i][bpCol], 0);

      if (currentBP < amount) {
        return { success: false, remaining: currentBP };
      }

      const newBP = currentBP - amount;
      sheet.getRange(i + 1, bpCol + 1).setValue(newBP);

      // Update LastUpdated
      const lastUpdatedCol = headers.indexOf('LastUpdated');
      if (lastUpdatedCol !== -1) {
        sheet.getRange(i + 1, lastUpdatedCol + 1).setValue(new Date());
      }

      return { success: true, remaining: newBP };
    }
  }

  return { success: false, remaining: 0 };
}