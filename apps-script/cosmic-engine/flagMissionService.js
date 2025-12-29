/**
 * Flag Missions Service
 * @fileoverview Manages Flag_Missions sheet:
 *   - Calculates flag mission points based on checkboxes
 *   - Syncs individual rows when edited
 *   - Syncs all rows on demand
 *
 * ASSUMPTION: Flag_Missions has PreferredName and various flag mission columns
 * ASSUMPTION: Flag missions use checkboxes (TRUE/FALSE) to mark completion
 */

// ============================================================================
// FLAG MISSION DEFINITIONS
// ============================================================================

/**
 * Flag mission point values
 * Each mission awards a specific number of points when completed
 */
const FLAG_MISSION_VALUES = {
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

/**
 * Gets the list of known flag mission column names
 * @return {Array<string>} Mission column names
 */
function getFlagMissionColumns() {
  return Object.keys(FLAG_MISSION_VALUES);
}

// ============================================================================
// SYNC FUNCTIONS
// ============================================================================

/**
 * Syncs flag mission points for a single row
 * Called by onEdit when Flag_Missions sheet is edited
 * @param {number} rowIndex - The 1-based row number that was edited
 */
function syncFlagMissionsRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || rowIndex <= 1) {
    return; // No sheet or header row
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find required columns
  const nameCol = headers.indexOf('PreferredName');
  const pointsColNames = ['Flag Mission Points', 'Flag Points'];
  let pointsCol = -1;
  for (const name of pointsColNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) {
      pointsCol = idx;
      break;
    }
  }

  if (nameCol === -1) {
    console.error('Flag_Missions missing PreferredName column');
    return;
  }

  // Get row data
  const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const playerName = rowData[nameCol];

  if (!playerName) return;

  // Calculate total points from completed missions
  let totalPoints = 0;
  for (let j = 0; j < headers.length; j++) {
    const header = headers[j];

    // Skip non-mission columns
    if (header === 'PreferredName' ||
        header === 'LastUpdated' ||
        header === 'Flag Mission Points' ||
        header === 'Flag Points') {
      continue;
    }

    const value = rowData[j];

    if (value === true) {
      // Checkbox is checked - award mission points
      totalPoints += FLAG_MISSION_VALUES[header] || 1;
    } else if (typeof value === 'number' && value > 0) {
      // Numeric value - use as-is
      totalPoints += value;
    }
  }

  // Update points column if it exists
  if (pointsCol !== -1) {
    sheet.getRange(rowIndex, pointsCol + 1).setValue(totalPoints);
  }

  // Update LastUpdated column if it exists
  const lastUpdatedCol = headers.indexOf('LastUpdated');
  if (lastUpdatedCol !== -1) {
    sheet.getRange(rowIndex, lastUpdatedCol + 1).setValue(new Date());
  }
}

/**
 * Syncs all flag missions for all players
 * Called from Scan Attendance / Missions menu
 */
function syncAllFlagMissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  let syncedCount = 0;
  const lastRow = sheet.getLastRow();

  for (let i = 2; i <= lastRow; i++) {
    syncFlagMissionsRow(i);
    syncedCount++;
  }

  logIntegrityAction('FLAG_MISSION_SYNC', {
    details: `Synced ${syncedCount} player flag missions`,
    status: 'SUCCESS'
  });

  return syncedCount;
}

/**
 * Recalculates flag missions for all players
 * Alias for syncAllFlagMissions
 */
function recalculateFlagMissions() {
  return syncAllFlagMissions();
}

// ============================================================================
// SCHEMA MANAGEMENT
// ============================================================================

/**
 * Ensures Flag_Missions sheet has proper schema
 */
function ensureFlagMissionsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Flag_Missions');

  const requiredHeaders = [
    'PreferredName',
    'Flag Mission Points',
    ...getFlagMissionColumns(),
    'LastUpdated'
  ];

  if (!sheet) {
    // Create new sheet
    sheet = ss.insertSheet('Flag_Missions');
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);

    // Format header
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#9c27b0')
      .setFontColor('#ffffff');

    // Set checkbox validation for mission columns
    const missionCols = getFlagMissionColumns();
    missionCols.forEach((name, idx) => {
      const colNum = requiredHeaders.indexOf(name) + 1;
      if (colNum > 0) {
        // Can't set validation on entire column without data, but ready for when data is added
      }
    });

    sheet.autoResizeColumns(1, requiredHeaders.length);

    logIntegrityAction('FLAG_SCHEMA_CREATE', {
      details: 'Created Flag_Missions sheet with schema',
      status: 'SUCCESS'
    });

    return sheet;
  }

  // Check existing headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
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

  return sheet;
}

// ============================================================================
// PLAYER MANAGEMENT
// ============================================================================

/**
 * Gets flag mission status for a player
 * @param {string} preferredName - The player name
 * @return {Object} Mission status {missions: {name: completed}, totalPoints}
 */
function getPlayerFlagMissions(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) {
    return { missions: {}, totalPoints: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');

  if (nameCol === -1) {
    return { missions: {}, totalPoints: 0 };
  }

  // Find player row
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const missions = {};
      let totalPoints = 0;

      for (let j = 0; j < headers.length; j++) {
        const header = headers[j];

        if (header === 'PreferredName' ||
            header === 'LastUpdated' ||
            header === 'Flag Mission Points' ||
            header === 'Flag Points') {
          continue;
        }

        const value = data[i][j];
        const completed = value === true || (typeof value === 'number' && value > 0);
        missions[header] = completed;

        if (completed) {
          totalPoints += FLAG_MISSION_VALUES[header] || 1;
        }
      }

      return { missions, totalPoints };
    }
  }

  return { missions: {}, totalPoints: 0 };
}

/**
 * Provisions a new player row in Flag_Missions
 * @param {string} preferredName - The player name
 * @return {boolean} True if row was added
 */
function provisionFlagMissionsRow(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet) {
    ensureFlagMissionsSchema();
    sheet = ss.getSheetByName('Flag_Missions');
  }

  // Check if player already exists
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');

  if (nameCol === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      return false; // Already exists
    }
  }

  // Create new row with defaults
  const newRow = headers.map(h => {
    if (h === 'PreferredName') return preferredName;
    if (h === 'LastUpdated') return new Date();
    if (h === 'Flag Mission Points' || h === 'Flag Points') return 0;
    return false; // Default checkbox state
  });

  sheet.appendRow(newRow);

  return true;
}

/**
 * Awards a flag mission to a player
 * @param {string} preferredName - The player name
 * @param {string} missionName - The mission column name
 * @return {Object} {success, message}
 */
function awardFlagMission(preferredName, missionName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet) {
    ensureFlagMissionsSchema();
    sheet = ss.getSheetByName('Flag_Missions');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const missionCol = headers.indexOf(missionName);

  if (nameCol === -1) {
    return { success: false, message: 'PreferredName column not found' };
  }

  if (missionCol === -1) {
    return { success: false, message: `Mission "${missionName}" column not found` };
  }

  // Find player row
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      // Check if already completed
      if (data[i][missionCol] === true) {
        return { success: false, message: 'Mission already completed' };
      }

      // Award mission
      sheet.getRange(i + 1, missionCol + 1).setValue(true);

      // Sync the row
      syncFlagMissionsRow(i + 1);

      const points = FLAG_MISSION_VALUES[missionName] || 1;

      logIntegrityAction('FLAG_MISSION_AWARD', {
        preferredName: preferredName,
        details: `Awarded ${missionName} (+${points} points)`,
        status: 'SUCCESS'
      });

      return { success: true, message: `Awarded ${points} points for ${missionName}` };
    }
  }

  // Player not found - provision and retry
  provisionFlagMissionsRow(preferredName);
  return awardFlagMission(preferredName, missionName);
}