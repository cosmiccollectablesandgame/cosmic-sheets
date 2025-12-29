/**
 * Attendance Service - Attendance Scanning and Mission Computation
 * @fileoverview Scans event sheets, builds Attendance_Calendar, computes missions, and syncs BP
 */

// ============================================================================
// ATTENDANCE_CALENDAR OPERATIONS
// ============================================================================

/**
 * Gets or creates Attendance_Calendar sheet
 * @return {Sheet} Attendance_Calendar sheet
 */
function getOrCreateAttendanceCalendar_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Calendar');

  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Calendar');
    sheet.appendRow([
      'Event_Date',
      'Event_ID',
      'Suffix',
      'preferred_name_id',
      'PreferredName',
      'Game',
      'Format',
      'Attended'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
  }

  return sheet;
}

/**
 * Gets or creates Attendance_Missions sheet
 * @return {Sheet} Attendance_Missions sheet
 */
function getOrCreateAttendanceMissions_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow([
      'PreferredName',
      'Points_From_Dice_Rolls',
      'Bonus_Points_From_Top4',
      'C_Total',
      'Badges',
      'LastUpdated'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
  }

  return sheet;
}

// ============================================================================
// ATTENDANCE SCAN
// ============================================================================

/**
 * Runs full attendance scan across all event sheets
 * Updates Attendance_Calendar with all attendance records
 * @return {Object} Scan result {eventsScanned, playersTracked, duration}
 */
function runAttendanceScan_() {
  const startTime = new Date();

  // Get or create calendar sheet
  const calendar = getOrCreateAttendanceCalendar_();

  // Clear existing data (keep headers)
  if (calendar.getLastRow() > 1) {
    calendar.deleteRows(2, calendar.getLastRow() - 1);
  }

  // Get all event sheets
  const eventSheets = getAllEventSheets_();

  let totalAttendanceRecords = 0;
  const playerSet = new Set();

  // Scan each event sheet
  eventSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const parsedEvent = parseEventId_(sheetName);

    if (!parsedEvent) return; // Skip if parsing failed

    const { date, id, suffix } = parsedEvent;

    // Read event sheet data
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // Skip empty sheets

    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');

    if (nameCol === -1) return; // Skip if no PreferredName column

    // Process each player row
    for (let i = 1; i < data.length; i++) {
      const preferredName = data[i][nameCol];
      if (!preferredName) continue;

      // Add attendance record
      const eventDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      calendar.appendRow([
        eventDate,
        id,
        suffix || '',
        preferredName, // preferred_name_id
        preferredName, // PreferredName
        suffix ? getGameFromSuffix_(suffix) : 'Commander',
        suffix ? getFormatFromSuffix_(suffix) : 'Commander',
        true // Attended
      ]);

      totalAttendanceRecords++;
      playerSet.add(preferredName);
    }
  });

  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2) + 's';

  // Log the scan
  logIntegrityAction('ATTENDANCE_SCAN', {
    details: `Scanned ${eventSheets.length} events, ${totalAttendanceRecords} attendance records, ${playerSet.size} unique players`,
    status: 'SUCCESS'
  });

  return {
    success: true,
    eventsScanned: eventSheets.length,
    playersTracked: playerSet.size,
    attendanceRecords: totalAttendanceRecords,
    duration
  };
}

/**
 * Gets game type from suffix
 * @param {string} suffix - Event suffix (single letter)
 * @return {string} Game type
 * @private
 */
function getGameFromSuffix_(suffix) {
  // Default mapping - can be customized
  const gameMap = {
    'A': 'Commander',
    'B': 'Commander',
    'C': 'Commander',
    'D': 'Modern',
    'E': 'Standard',
    'F': 'Pioneer'
  };
  return gameMap[suffix] || 'Commander';
}

/**
 * Gets format from suffix
 * @param {string} suffix - Event suffix (single letter)
 * @return {string} Format
 * @private
 */
function getFormatFromSuffix_(suffix) {
  // Default mapping - can be customized
  const formatMap = {
    'A': 'Commander',
    'B': 'Commander',
    'C': 'Commander',
    'D': 'Modern',
    'E': 'Standard',
    'F': 'Pioneer'
  };
  return formatMap[suffix] || 'Commander';
}

// ============================================================================
// MISSION COMPUTATION
// ============================================================================

/**
 * Runs mission computation based on Attendance_Calendar data
 * Updates Attendance_Missions sheet with computed mission progress
 * @return {Object} Computation result
 */
function runMissionComputation_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get Attendance_Calendar
  const calendarSheet = ss.getSheetByName('Attendance_Calendar');
  if (!calendarSheet) {
    throwError('Attendance_Calendar not found', 'CALENDAR_MISSING', 'Run attendance scan first');
  }

  // Get or create Attendance_Missions
  const missionsSheet = getOrCreateAttendanceMissions_();

  // Read calendar data
  const calendarData = calendarSheet.getDataRange().getValues();
  if (calendarData.length <= 1) {
    // No attendance data
    return {
      success: true,
      playersProcessed: 0,
      missionsComputed: 0
    };
  }

  const headers = calendarData[0];
  const nameCol = headers.indexOf('PreferredName');
  const dateCol = headers.indexOf('Event_Date');

  if (nameCol === -1 || dateCol === -1) {
    throwError('Invalid Attendance_Calendar schema', 'SCHEMA_INVALID');
  }

  // Group attendance by player
  const attendanceByPlayer = {};

  for (let i = 1; i < calendarData.length; i++) {
    const name = calendarData[i][nameCol];
    const date = calendarData[i][dateCol];

    if (!name) continue;

    if (!attendanceByPlayer[name]) {
      attendanceByPlayer[name] = [];
    }

    attendanceByPlayer[name].push({
      date: new Date(date),
      eventId: calendarData[i][1] // Event_ID column
    });
  }

  // Read MissionLog_1 and MissionLog_2 if they exist (for mission definitions)
  // For now, we'll compute simple attendance-based missions
  const missionResults = {};

  Object.keys(attendanceByPlayer).forEach(playerName => {
    const attendance = attendanceByPlayer[playerName];

    // Compute simple missions:
    // - Total attendance count
    // - Attendance in last 30 days
    // - Attendance in last 90 days

    const now = new Date();
    const thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const ninetyDaysAgo = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

    const totalAttendance = attendance.length;
    const last30Days = attendance.filter(a => a.date >= thirtyDaysAgo).length;
    const last90Days = attendance.filter(a => a.date >= ninetyDaysAgo).length;

    missionResults[playerName] = {
      totalAttendance,
      last30Days,
      last90Days,
      pointsFromDice: 0, // Placeholder - will be populated from another source
      bonusPointsFromTop4: 0 // Placeholder - will be populated from event results
    };
  });

  // Clear existing mission data (keep headers)
  if (missionsSheet.getLastRow() > 1) {
    missionsSheet.deleteRows(2, missionsSheet.getLastRow() - 1);
  }

  // Write mission results
  const missionRows = Object.keys(missionResults).map(playerName => {
    const result = missionResults[playerName];
    return [
      playerName,
      result.pointsFromDice,
      result.bonusPointsFromTop4,
      result.totalAttendance,
      '', // Badges placeholder
      dateISO()
    ];
  });

  if (missionRows.length > 0) {
    missionsSheet.getRange(2, 1, missionRows.length, 6).setValues(missionRows);
  }

  logIntegrityAction('MISSION_COMPUTE', {
    details: `Computed missions for ${missionRows.length} players`,
    status: 'SUCCESS'
  });

  return {
    success: true,
    playersProcessed: missionRows.length,
    missionsComputed: missionRows.length
  };
}

// ============================================================================
// BONUS POINTS AGGREGATION
// ============================================================================

/**
 * Runs bonus points aggregation from Attendance_Missions to BP_Total
 * Syncs mission-based BP into the player BP totals
 * @return {Object} Aggregation result
 */
function runBonusPointsAggregation_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get Attendance_Missions
  const missionsSheet = ss.getSheetByName('Attendance_Missions');
  if (!missionsSheet) {
    return {
      success: false,
      message: 'Attendance_Missions not found'
    };
  }

  // Get BP_Total
  ensureBPTotalSchema();
  const bpSheet = ss.getSheetByName('BP_Total');

  const missionsData = missionsSheet.getDataRange().getValues();
  if (missionsData.length <= 1) {
    return {
      success: true,
      playersUpdated: 0
    };
  }

  const headers = missionsData[0];
  const nameCol = headers.indexOf('PreferredName');
  const diceCol = headers.indexOf('Points_From_Dice_Rolls');
  const top4Col = headers.indexOf('Bonus_Points_From_Top4');
  const attendanceCol = headers.indexOf('C_Total');

  if (nameCol === -1) {
    throwError('Invalid Attendance_Missions schema', 'SCHEMA_INVALID');
  }

  let updatedCount = 0;

  // For each player in missions, update their BP
  for (let i = 1; i < missionsData.length; i++) {
    const playerName = missionsData[i][nameCol];
    if (!playerName) continue;

    const dicePoints = coerceNumber(missionsData[i][diceCol], 0);
    const top4Points = coerceNumber(missionsData[i][top4Col], 0);
    const attendanceCount = coerceNumber(missionsData[i][attendanceCol], 0);

    // Compute BP from missions:
    // - 1 BP per 2 attendance (simple rule)
    // - Add dice roll points
    // - Add top 4 bonus points
    const attendanceBP = Math.floor(attendanceCount / 2);
    const totalMissionBP = attendanceBP + dicePoints + top4Points;

    // Get current BP
    const currentBP = getPlayerBP(playerName);

    // For now, we'll just ensure the player has at least the mission BP
    // (Not subtracting, just ensuring minimum)
    if (totalMissionBP > 0) {
      const newBP = Math.max(currentBP, totalMissionBP);
      setPlayerBP(playerName, newBP);
      updatedCount++;
    }
  }

  logIntegrityAction('BP_AGGREGATE', {
    details: `Aggregated BP for ${updatedCount} players from missions`,
    status: 'SUCCESS'
  });

  return {
    success: true,
    playersUpdated: updatedCount
  };
}

// ============================================================================
// UNIFIED UPDATE FUNCTION
// ============================================================================

/**
 * Runs complete attendance → missions → BP pipeline
 * This is the main entry point for updating all attendance-related data
 * @return {Object} Complete result
 */
function onUpdateMissionsAndPoints() {
  try {
    const results = {
      scan: null,
      missions: null,
      bp: null,
      overallSuccess: false,
      errors: []
    };

    // Step 1: Attendance scan
    try {
      results.scan = runAttendanceScan_();
    } catch (e) {
      results.errors.push('Attendance Scan: ' + e.message);
      Logger.log('Attendance scan failed: ' + e.message);
    }

    // Step 2: Mission computation
    try {
      results.missions = runMissionComputation_();
    } catch (e) {
      results.errors.push('Mission Computation: ' + e.message);
      Logger.log('Mission computation failed: ' + e.message);
    }

    // Step 3: BP aggregation
    try {
      results.bp = runBonusPointsAggregation_();
    } catch (e) {
      results.errors.push('BP Aggregation: ' + e.message);
      Logger.log('BP aggregation failed: ' + e.message);
    }

    results.overallSuccess = results.errors.length === 0;

    // Show result to user
    const ui = SpreadsheetApp.getUi();
    if (results.overallSuccess) {
      ui.alert(
        '✅ Update Complete',
        `Attendance & Missions Updated!\n\n` +
        `Events Scanned: ${results.scan?.eventsScanned || 0}\n` +
        `Players Tracked: ${results.scan?.playersTracked || 0}\n` +
        `Missions Computed: ${results.missions?.missionsComputed || 0}\n` +
        `BP Updated: ${results.bp?.playersUpdated || 0}\n` +
        `Duration: ${results.scan?.duration || 'N/A'}`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '⚠️ Update Completed with Errors',
        'Some steps failed:\n\n' + results.errors.join('\n\n') +
        '\n\nCheck Execution Log for details.',
        ui.ButtonSet.OK
      );
    }

    return results;

  } catch (error) {
    Logger.log('onUpdateMissionsAndPoints error: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to update missions/points:\n\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    return {
      overallSuccess: false,
      errors: [error.message]
    };
  }
}