/**
 * Dice Service - Dice Point Management
 * @fileoverview Manages dice roll points, syncing to Prize-Wall-Points and BP_Total
 */

// ============================================================================
// DICE POINT AWARDING
// ============================================================================

/**
 * Awards dice point to a player (called from HTML UI)
 * @param {string} playerName - The canonical preferred_name_id
 * @param {number} pointsToAdd - Usually 1
 * @returns {Object} Result with success status
 */
function awardDicePoint(playerName, pointsToAdd = 1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Update Dice_Points
  const diceSheet = ss.getSheetByName('Dice_Points');
  if (!diceSheet) {
    return { success: false, message: 'Dice_Points sheet not found' };
  }

  const diceData = diceSheet.getDataRange().getValues();
  const diceHeaders = diceData[0];
  const nameCol = diceHeaders.indexOf('preferred_name_id');
  const pointsCol = diceHeaders.indexOf('Points') !== -1
    ? diceHeaders.indexOf('Points')
    : diceHeaders.indexOf('Dice Roll Points');

  if (nameCol === -1) {
    return { success: false, message: 'preferred_name_id column not found in Dice_Points' };
  }

  if (pointsCol === -1) {
    return { success: false, message: 'Points column not found in Dice_Points' };
  }

  let playerRow = -1;
  for (let i = 1; i < diceData.length; i++) {
    if (diceData[i][nameCol] === playerName) {
      playerRow = i + 1; // 1-indexed for Sheets
      break;
    }
  }

  if (playerRow === -1) {
    return { success: false, message: 'Player not found: ' + playerName };
  }

  // Increment points
  const currentPoints = diceData[playerRow - 1][pointsCol] || 0;
  const newPoints = currentPoints + pointsToAdd;
  diceSheet.getRange(playerRow, pointsCol + 1).setValue(newPoints);

  // 2. Sync to Players_Prize-Wall-Points
  syncDiceToWallPoints(playerName, newPoints);

  // 3. Sync to BP_Total.Dice Roll Points
  syncDiceToBPTotal(playerName, newPoints);

  // 4. Log the action
  logEvent('DICE_POINT', 'AWARD', playerName, {
    pointsAdded: pointsToAdd,
    newTotal: newPoints,
    source: 'HTML_UI'
  });

  return {
    success: true,
    message: `Awarded ${pointsToAdd} point(s) to ${playerName}. New total: ${newPoints}`
  };
}

/**
 * Sync dice points to Players_Prize-Wall-Points
 * @param {string} playerName - Player's preferred name
 * @param {number} points - New total points
 */
function syncDiceToWallPoints(playerName, points) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wallSheet = ss.getSheetByName('Players_Prize-Wall-Points');
  if (!wallSheet) return;

  const data = wallSheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const availCol = headers.indexOf('Dice_Points_Available');
  const updateCol = headers.indexOf('LastUpdated');

  if (nameCol === -1 || availCol === -1) return;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === playerName) {
      wallSheet.getRange(i + 1, availCol + 1).setValue(points);
      if (updateCol !== -1) {
        wallSheet.getRange(i + 1, updateCol + 1).setValue(new Date());
      }
      return;
    }
  }
}

/**
 * Sync dice points to BP_Total.Dice Roll Points
 * @param {string} playerName - Player's preferred name
 * @param {number} points - New total points
 */
function syncDiceToBPTotal(playerName, points) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpSheet = ss.getSheetByName('BP_Total');
  if (!bpSheet) return;

  const data = bpSheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const diceCol = headers.indexOf('Dice Roll Points');
  const updateCol = headers.indexOf('LastUpdated');

  if (nameCol === -1 || diceCol === -1) return; // Column doesn't exist

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === playerName) {
      bpSheet.getRange(i + 1, diceCol + 1).setValue(points);
      if (updateCol !== -1) {
        bpSheet.getRange(i + 1, updateCol + 1).setValue(new Date());
      }
      return;
    }
  }
}

// ============================================================================
// PLAYER LIST FOR UI
// ============================================================================

/**
 * Get list of players with current dice points for HTML UI
 * @returns {Array<Object>} Array of {name, dicePoints}
 */
function getPlayerList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefSheet = ss.getSheetByName('PreferredNames');
  const diceSheet = ss.getSheetByName('Dice_Points');

  if (!prefSheet) {
    return [];
  }

  // Get all preferred names
  const prefData = prefSheet.getDataRange().getValues();
  const prefNames = prefData.slice(1).map(row => row[0]).filter(n => n);

  // Get dice points lookup
  const diceMap = {};
  if (diceSheet) {
    const diceData = diceSheet.getDataRange().getValues();
    const diceHeaders = diceData[0];
    const nameCol = diceHeaders.indexOf('preferred_name_id');
    const pointsCol = diceHeaders.indexOf('Points') !== -1
      ? diceHeaders.indexOf('Points')
      : diceHeaders.indexOf('Dice Roll Points');

    if (nameCol !== -1 && pointsCol !== -1) {
      for (let i = 1; i < diceData.length; i++) {
        diceMap[diceData[i][nameCol]] = diceData[i][pointsCol] || 0;
      }
    }
  }

  // Return combined list
  return prefNames.map(name => ({
    name: name,
    dicePoints: diceMap[name] || 0
  })).sort((a, b) => a.name.localeCompare(b.name));
}

// ============================================================================
// LOGGING WRAPPER
// ============================================================================

/**
 * Logs a dice-related event to Integrity_Log
 * @param {string} category - Event category (e.g., 'DICE_POINT')
 * @param {string} action - Action type (e.g., 'AWARD')
 * @param {string} playerName - Player's preferred name
 * @param {Object} details - Additional details
 */
function logEvent(category, action, playerName, details = {}) {
  logIntegrityAction(`${category}_${action}`, {
    preferredName: playerName,
    details: JSON.stringify(details),
    status: 'SUCCESS'
  });
}

// ============================================================================
// DICE SHEET SETUP
// ============================================================================

/**
 * Ensures Dice_Points sheet exists with proper schema
 */
function ensureDicePointsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dice_Points');

  if (!sheet) {
    sheet = ss.insertSheet('Dice_Points');
    sheet.appendRow([
      'preferred_name_id',
      'Points',
      'LastUpdated'
    ]);
    sheet.setFrozenRows(1);

    logIntegrityAction('SHEET_CREATE', {
      details: 'Created Dice_Points sheet',
      status: 'SUCCESS'
    });
  }

  return sheet;
}