/**
 * ════════════════════════════════════════════════════════════════════════
 * COSMIC MISSION SYSTEM - SUFFIX SERVICE
 *
 * Mission evaluation and awarding based on suffix-driven attendance
 * Updates per-player mission counters in MissionLog sheet
 *
 * Compatible with: Engine v7.9.6+
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════
// MISSION EVALUATION
// ══════════════════════════════════════════════════════════════════════

/**
 * Evaluate mission triggers for a given player/event based on suffix.
 * Called during MissionLog sync / after Prize Engine commit.
 *
 * @param {string} playerId - canonical PreferredNames ID
 * @param {string} eventId - sheet/tab name (e.g. "11-23-B-2025")
 * @param {number|null} rank - final rank (1 = first, etc.) or null
 */
function evaluateSuffixMissions_(playerId, eventId, rank) {
  if (!playerId || !eventId) {
    Logger.log('evaluateSuffixMissions_: missing playerId or eventId');
    return;
  }

  var suffix = getSuffixFromEventId_(eventId);
  var meta = getSuffixMeta_(suffix);

  // If no suffix or invalid suffix, still log but don't award missions
  if (!meta) {
    Logger.log('No suffix meta for event ' + eventId + ', skipping mission evaluation');
    return;
  }

  // Commander bracketed missions
  if (suffix === 'B') {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_CMD_CASUAL, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  if (suffix === 'C') {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_CMD_TRANSITION, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  if (suffix === 'T') {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_CMD_CEDH, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  // Limited formats: Draft, Proxy/Cube Draft, Prerelease, Sealed
  if (meta.requiresKitPrompt) {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_LIMITED_EVENT, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  // Academy / Outreach
  if (suffix === 'A') {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_ACADEMY, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  if (suffix === 'E') {
    awardMissionProgress_(playerId, MISSION_IDS.ATTEND_OUTREACH, {
      eventId: eventId,
      rank: rank,
      suffix: suffix
    });
  }

  // Future: non-suffix missions (streaks, Nth event, etc.) can be added here.
}

/**
 * Batch process attendance-based missions over a date range.
 * This is the batch version that uses the attendance scan.
 *
 * @param {Date} startDate - Start date
 * @param {Date} endDate - End date
 * @return {Object} Processing stats
 */
function runAttendanceMissionsForRange_(startDate, endDate) {
  var records = scanAttendanceForRange_(startDate, endDate);

  var processedCount = 0;
  var errorCount = 0;

  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    try {
      evaluateSuffixMissions_(rec.playerId, rec.eventId, rec.rank);
      processedCount++;
    } catch (e) {
      Logger.log('Error evaluating missions for ' + rec.playerId + ' at ' + rec.eventId + ': ' + e);
      errorCount++;
    }
  }

  return {
    totalRecords: records.length,
    processed: processedCount,
    errors: errorCount
  };
}

// ══════════════════════════════════════════════════════════════════════
// MISSION AWARDING
// ══════════════════════════════════════════════════════════════════════

/**
 * Award mission progress to a player.
 *
 * Responsibilities:
 *  - Ensure the player has a MissionLog row.
 *  - Increment appropriate mission counters.
 *  - Enforce repeat rules (e.g., once per event, once per day).
 *
 * This function:
 *  - Uses `missionId` to locate the correct column in MissionLog
 *    (missionId == column header string).
 *  - Logs or safely no-ops if the column is missing.
 *  - Updates Total_Points based on mission progression.
 *
 * @param {string} playerId - Player ID
 * @param {string} missionId - Mission ID (column header)
 * @param {Object} context - Mission context {eventId, rank, suffix}
 */
function awardMissionProgress_(playerId, missionId, context) {
  if (!playerId || !missionId) {
    Logger.log('awardMissionProgress_: missing playerId or missionId');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);

    if (!missionSheet) {
      Logger.log('MissionLog sheet not found');
      return;
    }

    // Get headers
    var lastCol = missionSheet.getLastColumn();
    var headers = missionSheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Find column indices
    var playerIdCol = headers.indexOf('preferred_name_id');
    var totalPointsCol = headers.indexOf('Total_Points');
    var missionCol = headers.indexOf(missionId);

    if (playerIdCol === -1) {
      Logger.log('preferred_name_id column not found in MissionLog');
      return;
    }

    if (missionCol === -1) {
      Logger.log('Mission column ' + missionId + ' not found in MissionLog, skipping');
      return;
    }

    // Find player row
    var lastRow = missionSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('MissionLog has no data rows');
      return;
    }

    var playerIdData = missionSheet.getRange(2, playerIdCol + 1, lastRow - 1, 1).getValues();
    var playerRow = -1;

    for (var i = 0; i < playerIdData.length; i++) {
      if (String(playerIdData[i][0]).trim() === playerId) {
        playerRow = i + 2; // +2 because we start from row 2
        break;
      }
    }

    if (playerRow === -1) {
      Logger.log('Player ' + playerId + ' not found in MissionLog, cannot award mission');
      return;
    }

    // Get current mission progress
    var currentValue = missionSheet.getRange(playerRow, missionCol + 1).getValue();
    var currentProgress = coerceNumber(currentValue, 0);

    // Increment mission progress (simple increment for now)
    // TODO: Later can be customized to use per-mission Reward_Qty from MissionLog_1/2
    var newProgress = currentProgress + 1;
    missionSheet.getRange(playerRow, missionCol + 1).setValue(newProgress);

    // Update Total_Points if column exists
    if (totalPointsCol !== -1) {
      var currentPoints = coerceNumber(
        missionSheet.getRange(playerRow, totalPointsCol + 1).getValue(),
        0
      );
      // Award 1 point per mission completion (can be customized)
      // TODO: Later can pull actual point value from mission definitions
      var newPoints = currentPoints + 1;
      missionSheet.getRange(playerRow, totalPointsCol + 1).setValue(newPoints);
    }

    // Log the award
    Logger.log(
      'Awarded ' + missionId + ' to ' + playerId + ': ' + currentProgress + ' -> ' + newProgress +
      ' (event: ' + (context.eventId || 'unknown') + ')'
    );

  } catch (e) {
    Logger.log('Failed to award mission ' + missionId + ' to ' + playerId + ': ' + e);
  }
}

// ══════════════════════════════════════════════════════════════════════
// MISSIONLOG ROW MANAGEMENT
// ══════════════════════════════════════════════════════════════════════

/**
 * Ensure a player has a MissionLog row
 * Creates row if missing
 *
 * @param {string} playerId - Player ID
 * @return {boolean} True if row exists or was created
 */
function ensureMissionLogRow_(playerId) {
  if (!playerId) return false;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);

    if (!missionSheet) {
      Logger.log('MissionLog sheet not found');
      return false;
    }

    // Check if player already has a row
    var lastCol = missionSheet.getLastColumn();
    var headers = missionSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var playerIdCol = headers.indexOf('preferred_name_id');

    if (playerIdCol === -1) {
      Logger.log('preferred_name_id column not found');
      return false;
    }

    var lastRow = missionSheet.getLastRow();
    if (lastRow > 1) {
      var playerIds = missionSheet.getRange(2, playerIdCol + 1, lastRow - 1, 1).getValues();

      for (var i = 0; i < playerIds.length; i++) {
        if (String(playerIds[i][0]).trim() === playerId) {
          return true; // Row already exists
        }
      }
    }

    // Create new row for player
    var newRow = [];
    for (var j = 0; j < lastCol; j++) {
      newRow.push(0);
    }
    newRow[playerIdCol] = playerId;

    // Initialize Total_Points to 0
    var totalPointsCol = headers.indexOf('Total_Points');
    if (totalPointsCol !== -1) {
      newRow[totalPointsCol] = 0;
    }

    missionSheet.appendRow(newRow);
    Logger.log('Created MissionLog row for ' + playerId);
    return true;

  } catch (e) {
    Logger.log('Failed to ensure MissionLog row for ' + playerId + ': ' + e);
    return false;
  }
}

/**
 * Sync all PreferredNames players to MissionLog
 * Ensures every player has a MissionLog row
 *
 * @return {Object} Sync stats {existing, created, errors}
 */
function syncAllPlayersToMissionLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prefSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.PREFERRED_NAMES);
  var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);

  if (!prefSheet || !missionSheet) {
    Logger.log('PreferredNames or MissionLog sheet missing');
    return { existing: 0, created: 0, errors: 1 };
  }

  // Get all player IDs from PreferredNames (column A)
  var prefLastRow = prefSheet.getLastRow();
  if (prefLastRow <= 1) {
    return { existing: 0, created: 0, errors: 0 };
  }

  var prefIds = prefSheet.getRange(2, 1, prefLastRow - 1, 1).getValues();

  var existingCount = 0;
  var createdCount = 0;
  var errorCount = 0;

  // Get existing MissionLog IDs for batch checking
  var missionLastRow = missionSheet.getLastRow();
  var missionLastCol = missionSheet.getLastColumn();
  var headers = missionSheet.getRange(1, 1, 1, missionLastCol).getValues()[0];
  var playerIdCol = headers.indexOf('preferred_name_id');

  if (playerIdCol === -1) {
    Logger.log('preferred_name_id column not found');
    return { existing: 0, created: 0, errors: 1 };
  }

  var existingIds = new Set();
  if (missionLastRow > 1) {
    var existingData = missionSheet.getRange(2, playerIdCol + 1, missionLastRow - 1, 1).getValues();
    for (var i = 0; i < existingData.length; i++) {
      existingIds.add(String(existingData[i][0]).trim());
    }
  }

  // Create rows for missing players
  var rowsToAdd = [];
  for (var i = 0; i < prefIds.length; i++) {
    var playerId = String(prefIds[i][0]).trim();
    if (!playerId) continue;

    try {
      if (existingIds.has(playerId)) {
        existingCount++;
      } else {
        // Build new row
        var newRow = [];
        for (var j = 0; j < missionLastCol; j++) {
          newRow.push(0);
        }
        newRow[playerIdCol] = playerId;

        var totalPointsCol = headers.indexOf('Total_Points');
        if (totalPointsCol !== -1) {
          newRow[totalPointsCol] = 0;
        }

        rowsToAdd.push(newRow);
        createdCount++;
      }
    } catch (e) {
      Logger.log('Failed to sync ' + playerId + ': ' + e);
      errorCount++;
    }
  }

  // Batch append all new rows
  if (rowsToAdd.length > 0) {
    var startRow = missionSheet.getLastRow() + 1;
    missionSheet.getRange(startRow, 1, rowsToAdd.length, missionLastCol).setValues(rowsToAdd);
    Logger.log('Batch created ' + rowsToAdd.length + ' MissionLog rows');
  }

  return {
    existing: existingCount,
    created: createdCount,
    errors: errorCount
  };
}