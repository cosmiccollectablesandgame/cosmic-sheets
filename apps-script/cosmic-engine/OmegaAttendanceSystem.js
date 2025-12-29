/**
 * ════════════════════════════════════════════════════════════════════════
 * COSMIC ATTENDANCE TRACKING SYSTEM - OMEGA
 *
 * Purpose: Scan all event sheets, build Attendance_Calendar, compute mission progress
 * Source of Truth: MissionLog_1 and MissionLog_2 (mission definitions)
 * Computed Outputs: Attendance_Calendar, Attendance_Missions
 *
 * Version: 1.1.0 - Performance Optimized with Robust Error Handling
 * Compatible with: Engine v7.9.6+
 * Manual: v3.9.6
 *
 * Improvements:
 * - Batch operations for 100x faster sheet writes
 * - Fixed streak calculation for actual consecutive weeks
 * - Comprehensive validation and error handling
 * - Support for all placement tiers (not just Top 4)
 * - Memory-optimized data structures
 * - Integration with Mission Suffix Service and Gate I
 * ════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════
// MAIN ATTENDANCE SCANNING FUNCTION
// ══════════════════════════════════════════════════════════════════════

/**
 * Omega Attendance Scan - Complete mission tracking system
 *
 * Workflow:
 * 1. Validate required sheets exist
 * 2. Load mission definitions (source of truth)
 * 3. Scan all event sheets (MM-DD-YYYY format)
 * 4. Extract attendance and placement data
 * 5. Build Attendance_Calendar (player × event matrix)
 * 6. Compute mission progress from MissionLog_1 and MissionLog_2
 * 7. Update Attendance_Missions with computed values
 * 8. Log summary to Integrity_Log
 *
 * Menu Trigger: 🎯 Systems > Scan + Update Missions
 */
function runOmegaAttendanceScan() {
  var startTime = new Date();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    ui.alert('Starting Omega Attendance Scan...',
             'This will scan all event sheets and update mission progress.',
             ui.ButtonSet.OK);

    // Step 1: Validate environment
    validateRequiredSheets(ss);

    // Step 2: Load mission definitions (source of truth)
    var missionDefs = loadMissionDefinitions(ss);
    if (Object.keys(missionDefs).length === 0) {
      throw new Error('No mission definitions found in MissionLog_1 or MissionLog_2');
    }

    // Step 3: Scan all event sheets
    var eventData = scanAllEventSheets(ss);
    if (eventData.events.length === 0) {
      ui.alert('No Events Found',
               'No valid event sheets were found. Event sheets must follow the format MM-DD-YYYY or MM-DD-[SUFFIX]-YYYY',
               ui.ButtonSet.OK);
      return;
    }

    // Step 4: Build attendance calendar (with batch writes)
    buildAttendanceCalendar(ss, eventData);

    // Step 5: Compute mission progress
    var missionProgress = computeMissionProgress(eventData, missionDefs);

    // Step 6: Update Attendance_Missions (with batch writes)
    updateAttendanceMissions(ss, missionProgress, missionDefs);

    // Step 7: Log to Integrity_Log
    var duration = (new Date() - startTime) / 1000;
    logAttendanceScan(ss, {
      eventsScanned: eventData.events.length,
      playersTracked: eventData.players.size,
      missionsComputed: Object.keys(missionDefs).length,
      duration: duration,
      dateRange: {
        earliest: eventData.events[0] ? eventData.events[0].date.toISOString().split('T')[0] : null,
        latest: eventData.events[eventData.events.length - 1] ? eventData.events[eventData.events.length - 1].date.toISOString().split('T')[0] : null
      }
    });

    // Success message
    ui.alert('✅ Attendance Scan Complete!',
             'Events Scanned: ' + eventData.events.length + '\n' +
             'Players Tracked: ' + eventData.players.size + '\n' +
             'Missions Computed: ' + Object.keys(missionDefs).length + '\n' +
             'Duration: ' + duration.toFixed(2) + 's\n\n' +
             'Date Range: ' + (eventData.events[0] ? eventData.events[0].date.toLocaleDateString() : 'N/A') +
             ' to ' + (eventData.events[eventData.events.length - 1] ? eventData.events[eventData.events.length - 1].date.toLocaleDateString() : 'N/A'),
             ui.ButtonSet.OK);

  } catch (error) {
    Logger.log('ERROR in runOmegaAttendanceScan: ' + error.toString());
    Logger.log('Stack trace: ' + (error.stack || 'N/A'));
    ui.alert('❌ Error in Attendance Scan',
             error.toString() + '\n\nCheck execution logs for details.',
             ui.ButtonSet.OK);

    // Log error to Integrity_Log
    try {
      logAttendanceError(ss, error);
    } catch (logError) {
      Logger.log('Failed to log error: ' + logError.toString());
    }
  }
}

/**
 * Wrapper function for menu integration
 */
function scanAndUpdateMissions() {
  runOmegaAttendanceScan();
}

// ══════════════════════════════════════════════════════════════════════
// VALIDATION
// ══════════════════════════════════════════════════════════════════════

/**
 * Validate that required sheets exist and create them if missing
 */
function validateRequiredSheets(ss) {
  var requiredSheets = [
    ATTENDANCE_CONFIG.SHEETS.MISSION_LOG_1,
    ATTENDANCE_CONFIG.SHEETS.MISSION_LOG_2
  ];

  var missingSheets = [];

  for (var i = 0; i < requiredSheets.length; i++) {
    if (!ss.getSheetByName(requiredSheets[i])) {
      missingSheets.push(requiredSheets[i]);
    }
  }

  if (missingSheets.length > 0) {
    throw new Error('Missing required sheets: ' + missingSheets.join(', ') + '\n\nPlease create these sheets with mission definitions before running the scan.');
  }

  // Create output sheets if they don't exist
  var outputSheets = [
    ATTENDANCE_CONFIG.SHEETS.CALENDAR,
    ATTENDANCE_CONFIG.SHEETS.MISSIONS,
    ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG
  ];

  for (var i = 0; i < outputSheets.length; i++) {
    if (!ss.getSheetByName(outputSheets[i])) {
      Logger.log('Creating missing output sheet: ' + outputSheets[i]);
      ss.insertSheet(outputSheets[i]);
    }
  }
}

// ══════════════════════════════════════════════════════════════════════
// MISSION DEFINITION LOADING
// ══════════════════════════════════════════════════════════════════════

/**
 * Load mission definitions from MissionLog_1 and MissionLog_2
 * These sheets are the SOURCE OF TRUTH for mission rules
 *
 * Expected Format:
 * Column A: Mission_ID (e.g., "ATTEND_10", "TOP4_MTG", etc.)
 * Column B: Mission_Name
 * Column C: Mission_Type (attendance, placement, format, streak, etc.)
 * Column D: Criteria (JSON or text describing requirements)
 * Column E: Points_Value
 * Column F: Cap (unlimited = 0, otherwise max count)
 * Column G: (Optional) Active (TRUE/FALSE, defaults to TRUE)
 */
function loadMissionDefinitions(ss) {
  var missions = {};

  // Load from both MissionLog sheets
  var logSheets = [
    ATTENDANCE_CONFIG.SHEETS.MISSION_LOG_1,
    ATTENDANCE_CONFIG.SHEETS.MISSION_LOG_2
  ];

  for (var s = 0; s < logSheets.length; s++) {
    var sheetName = logSheets[s];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    try {
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) continue; // Only headers or empty

      for (var i = 1; i < data.length; i++) {
        var row = data[i];

        // Skip empty rows
        if (!row[0]) continue;

        var missionId = String(row[0]).trim();

        // Skip if marked inactive
        if (row[6] === false || String(row[6]).toLowerCase() === 'false') {
          Logger.log('Skipping inactive mission: ' + missionId);
          continue;
        }

        missions[missionId] = {
          id: missionId,
          name: row[1] ? String(row[1]).trim() : missionId,
          type: row[2] ? String(row[2]).trim() : 'attendance',
          criteria: row[3] ? String(row[3]).trim() : '',
          pointsValue: parseFloat(row[4]) || 1,
          cap: parseInt(row[5], 10) || 0, // 0 = unlimited
          source: sheetName
        };
      }
    } catch (error) {
      Logger.log('Error loading missions from ' + sheetName + ': ' + error.toString());
    }
  }

  Logger.log('Loaded ' + Object.keys(missions).length + ' mission definitions');
  return missions;
}

// ══════════════════════════════════════════════════════════════════════
// ATTENDANCE CALENDAR BUILDER
// ══════════════════════════════════════════════════════════════════════

/**
 * Build Attendance_Calendar sheet using BATCH OPERATIONS for performance
 *
 * Format:
 * Column A: PreferredName
 * Column B: Total_Events_Attended
 * Column C+: Event dates (chronological) with ✓ for attendance
 */
function buildAttendanceCalendar(ss, eventData) {
  var calendarSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.CALENDAR);

  if (!calendarSheet) {
    calendarSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.CALENDAR);
  }

  // Clear existing data
  calendarSheet.clear();

  // Build header row
  var headers = ['PreferredName', 'Total_Events_Attended'];
  for (var i = 0; i < eventData.events.length; i++) {
    headers.push(eventData.events[i].sheetName);
  }

  // Build all rows in memory (MUCH faster than appendRow)
  var rows = [headers];
  var playerArray = [];
  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    playerArray.push(result.value);
    result = playerIterator.next();
  }

  playerArray.sort();

  for (var p = 0; p < playerArray.length; p++) {
    var player = playerArray[p];
    var row = [player];
    var totalAttended = 0;

    // Check attendance for each event
    for (var e = 0; e < eventData.events.length; e++) {
      var event = eventData.events[e];
      if (event.players.indexOf(player) !== -1) {
        row.push('✓');
        totalAttended++;
      } else {
        row.push('');
      }
    }

    // Insert total at column B
    row.splice(1, 0, totalAttended);
    rows.push(row);
  }

  // Write all data at once (100x faster than appendRow loop)
  if (rows.length > 0) {
    calendarSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  }

  // Format header
  calendarSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#e0e0e0')
    .setHorizontalAlignment('center');

  // Format total column
  if (playerArray.length > 0) {
    calendarSheet.getRange(2, 2, playerArray.length, 1)
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
  }

  // Freeze header and name columns
  calendarSheet.setFrozenRows(1);
  calendarSheet.setFrozenColumns(2);

  Logger.log('Built Attendance_Calendar with ' + playerArray.length + ' players and ' + eventData.events.length + ' events');
}

// ══════════════════════════════════════════════════════════════════════
// MISSION PROGRESS COMPUTATION
// ══════════════════════════════════════════════════════════════════════

/**
 * Compute mission progress for all players based on event data and mission definitions
 *
 * Returns:
 * {
 *   'PlayerName1': {
 *     'ATTEND_10': 8,  // Attended 8 events (cap at 10)
 *     'TOP4_MTG': 3,   // Placed Top 4 three times (unlimited cap)
 *     ...
 *   },
 *   ...
 * }
 */
function computeMissionProgress(eventData, missionDefs) {
  var progress = {};

  // Initialize progress for all players
  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    progress[player] = {};
    result = playerIterator.next();
  }

  // Compute each mission type
  var missionIds = Object.keys(missionDefs);
  for (var m = 0; m < missionIds.length; m++) {
    var missionId = missionIds[m];
    var mission = missionDefs[missionId];

    try {
      var missionType = mission.type.toLowerCase();

      if (missionType === 'attendance') {
        computeAttendanceMission(mission, eventData, progress);
      } else if (missionType === 'placement') {
        computePlacementMission(mission, eventData, progress);
      } else if (missionType === 'format') {
        computeFormatMission(mission, eventData, progress);
      } else if (missionType === 'streak') {
        computeStreakMission(mission, eventData, progress);
      } else if (missionType === 'win' || missionType === '1st') {
        computeWinMission(mission, eventData, progress);
      } else {
        Logger.log('Unknown mission type: ' + mission.type + ' for ' + missionId);
      }
    } catch (error) {
      Logger.log('Error computing mission ' + missionId + ': ' + error.toString());
      // Initialize to 0 for all players if computation fails
      var playerIterator2 = eventData.players.values();
      var result2 = playerIterator2.next();

      while (!result2.done) {
        var player2 = result2.value;
        progress[player2][missionId] = 0;
        result2 = playerIterator2.next();
      }
    }
  }

  return progress;
}

/**
 * Compute attendance-based missions
 * Example: "Attend 10 events" → count total events attended
 */
function computeAttendanceMission(mission, eventData, progress) {
  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    var count = 0;

    // Count total attendance
    for (var i = 0; i < eventData.events.length; i++) {
      var event = eventData.events[i];
      if (event.players.indexOf(player) !== -1) {
        count++;
      }
    }

    // Apply cap if set
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }

    progress[player][mission.id] = count;
    result = playerIterator.next();
  }
}

/**
 * Compute placement-based missions
 * Example: "Finish Top 4" → count Top 4 placements
 * Criteria can be: "4" for Top 4, "8" for Top 8, "2-3" for 2nd or 3rd, etc.
 */
function computePlacementMission(mission, eventData, progress) {
  // Parse placement criteria
  var minPlace = 1;
  var maxPlace = 4;

  if (mission.criteria) {
    var criteria = String(mission.criteria).trim();

    if (criteria.indexOf('-') !== -1) {
      // Range: "2-3" or "1-4"
      var parts = criteria.split('-');
      minPlace = parseInt(parts[0], 10) || 1;
      maxPlace = parseInt(parts[1], 10) || 4;
    } else {
      // Single number: "4" means 1-4, "8" means 1-8
      maxPlace = parseInt(criteria, 10) || 4;
      minPlace = 1;
    }
  }

  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    var count = 0;

    for (var i = 0; i < eventData.events.length; i++) {
      var event = eventData.events[i];
      var standing = event.placements[player];
      if (standing && standing >= minPlace && standing <= maxPlace) {
        count++;
      }
    }

    // Apply cap if set (unlimited missions accumulate)
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }

    progress[player][mission.id] = count;
    result = playerIterator.next();
  }
}

/**
 * Compute win-based missions (1st place only)
 */
function computeWinMission(mission, eventData, progress) {
  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    var count = 0;

    for (var i = 0; i < eventData.events.length; i++) {
      var event = eventData.events[i];
      if (event.placements[player] === 1) {
        count++;
      }
    }

    // Apply cap if set
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }

    progress[player][mission.id] = count;
    result = playerIterator.next();
  }
}

/**
 * Compute format-specific missions
 * Example: "Attend 5 Commander events"
 * Criteria: format name or suffix (e.g., "Commander", "C", "Draft", "D")
 */
function computeFormatMission(mission, eventData, progress) {
  // Parse criteria for format requirement (e.g., "Commander", "Draft", "C")
  var targetCriteria = String(mission.criteria).toLowerCase().trim();

  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    var count = 0;

    for (var i = 0; i < eventData.events.length; i++) {
      var event = eventData.events[i];
      var matchesFormat = event.format.toLowerCase().indexOf(targetCriteria) !== -1 ||
                         event.suffix.toLowerCase() === targetCriteria;

      if (matchesFormat && event.players.indexOf(player) !== -1) {
        count++;
      }
    }

    // Apply cap if set
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }

    progress[player][mission.id] = count;
    result = playerIterator.next();
  }
}

/**
 * Compute streak-based missions
 * Example: "Attend 3 consecutive weeks"
 *
 * FIXED: Now correctly checks consecutive weeks, not consecutive events
 * Criteria: number of weeks required (e.g., "3" for 3 consecutive weeks)
 */
function computeStreakMission(mission, eventData, progress) {
  var requiredWeeks = parseInt(mission.criteria, 10) || 3;

  var playerIterator = eventData.players.values();
  var result = playerIterator.next();

  while (!result.done) {
    var player = result.value;
    var history = eventData.playerEventHistory.get(player);
    if (!history || history.length === 0) {
      progress[player][mission.id] = 0;
      result = playerIterator.next();
      continue;
    }

    var maxStreak = 0;
    var currentStreak = 1;
    var lastWeekNumber = getWeekNumber(history[0].event.date);
    var lastYear = history[0].event.date.getFullYear();

    for (var i = 1; i < history.length; i++) {
      var eventDate = history[i].event.date;
      var currentWeekNumber = getWeekNumber(eventDate);
      var currentYear = eventDate.getFullYear();

      // Check if this is the next consecutive week
      var isConsecutive = false;

      if (currentYear === lastYear) {
        isConsecutive = (currentWeekNumber === lastWeekNumber + 1);
      } else if (currentYear === lastYear + 1) {
        // Handle year boundary: last week of year → first week of next year
        var weeksInLastYear = getWeeksInYear(lastYear);
        isConsecutive = (lastWeekNumber === weeksInLastYear && currentWeekNumber === 1);
      }

      if (isConsecutive) {
        currentStreak++;
        maxStreak = Math.max(maxStreak, currentStreak);
      } else if (currentWeekNumber !== lastWeekNumber || currentYear !== lastYear) {
        // Not consecutive, but not the same week either
        currentStreak = 1;
      }
      // If same week, don't increment or reset streak

      lastWeekNumber = currentWeekNumber;
      lastYear = currentYear;
    }

    // Mission typically checks if streak >= required weeks
    // Store the max streak achieved
    var value = maxStreak >= requiredWeeks ? maxStreak : 0;

    // Apply cap if set
    if (mission.cap > 0) {
      value = Math.min(value, mission.cap);
    }

    progress[player][mission.id] = value;
    result = playerIterator.next();
  }
}

// ══════════════════════════════════════════════════════════════════════
// ATTENDANCE_MISSIONS UPDATE
// ══════════════════════════════════════════════════════════════════════

/**
 * Update Attendance_Missions sheet with computed progress
 * Uses BATCH OPERATIONS for performance
 *
 * Format:
 * Row 1: Headers [PreferredName, Mission_1, Mission_2, ...]
 * Row 2+: [PlayerName, progress_1, progress_2, ...]
 */
function updateAttendanceMissions(ss, missionProgress, missionDefs) {
  var missionsSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSIONS);

  if (!missionsSheet) {
    missionsSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.MISSIONS);
  }

  // Clear existing data
  missionsSheet.clear();

  // Get mission IDs (sorted) and build headers
  var playerList = Object.keys(missionProgress);
  if (playerList.length === 0) {
    Logger.log('No players to update in Attendance_Missions');
    return;
  }

  var missionIds = Object.keys(missionProgress[playerList[0]] || {});
  missionIds.sort();

  // Build header row with mission names
  var headers = ['PreferredName'];
  for (var i = 0; i < missionIds.length; i++) {
    var missionId = missionIds[i];
    var missionName = missionDefs[missionId] ? missionDefs[missionId].name : missionId;
    headers.push(missionName);
  }

  // Build all rows in memory
  var rows = [headers];
  playerList.sort();

  for (var p = 0; p < playerList.length; p++) {
    var player = playerList[p];
    var row = [player];

    for (var m = 0; m < missionIds.length; m++) {
      var missionId = missionIds[m];
      row.push(missionProgress[player][missionId] || 0);
    }

    rows.push(row);
  }

  // Write all data at once (BATCH OPERATION)
  if (rows.length > 0) {
    missionsSheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  }

  // Format header
  missionsSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Format data columns as numbers
  if (playerList.length > 0 && missionIds.length > 0) {
    missionsSheet.getRange(2, 2, playerList.length, missionIds.length)
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
  }

  // Freeze header and name column
  missionsSheet.setFrozenRows(1);
  missionsSheet.setFrozenColumns(1);

  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    missionsSheet.autoResizeColumn(i);
  }

  Logger.log('Updated Attendance_Missions for ' + playerList.length + ' players with ' + missionIds.length + ' missions');
}