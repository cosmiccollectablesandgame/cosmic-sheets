/**
 * ════════════════════════════════════════════════════════════════════════
 * COSMIC MISSION SYSTEM - GATE SERVICE
 *
 * Gate I - MissionLog & Attendance Sync Health Check
 * Part of Ship-Gates system (A-I)
 *
 * Verifies:
 * - MissionLog sheet exists and has required columns
 * - PreferredNames sheet exists
 * - Every PreferredNames player has a MissionLog row
 * - Attendance scan for the last 7 days can run without throwing
 * - No recent attendance players are missing MissionLog rows
 *
 * Compatible with: Engine v7.9.6+
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════
// GATE I: MISSIONLOG & ATTENDANCE SYNC
// ══════════════════════════════════════════════════════════════════════

/**
 * Gate I - MissionLog & Attendance Sync
 *
 * Verifies:
 *  - MissionLog sheet exists and has required columns
 *  - PreferredNames sheet exists
 *  - Every PreferredNames player has a MissionLog row
 *  - Attendance scan for the last 7 days can run without throwing
 *  - No recent attendance players are missing MissionLog rows
 *
 * @return {Object} { gate, name, pass, details, autoFixApplied }
 */
function checkGateI_MissionLog_() {
  var result = {
    gate: 'I',
    name: 'MissionLog & Attendance',
    pass: true,
    details: [],
    autoFixApplied: false
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);
    var prefSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.PREFERRED_NAMES);

    // -----------------------------------------------------------------------
    // 1) Check sheet existence
    // -----------------------------------------------------------------------
    if (!missionSheet) {
      result.pass = false;
      result.details.push('MissionLog sheet is missing');
      return result;
    }

    if (!prefSheet) {
      result.pass = false;
      result.details.push('PreferredNames sheet is missing');
      return result;
    }

    // -----------------------------------------------------------------------
    // 2) Schema sanity check (headers)
    // -----------------------------------------------------------------------
    var lastCol = missionSheet.getLastColumn();
    if (lastCol === 0) {
      result.pass = false;
      result.details.push('MissionLog has no columns');
      return result;
    }

    var headerRow = missionSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var headerSet = new Set();
    for (var i = 0; i < headerRow.length; i++) {
      headerSet.add(String(headerRow[i]).trim());
    }

    // Required primary columns
    var requiredPrimaryCols = [
      'preferred_name_id',
      'Total_Points'
    ];

    for (var i = 0; i < requiredPrimaryCols.length; i++) {
      if (!headerSet.has(requiredPrimaryCols[i])) {
        result.pass = false;
        result.details.push('MissionLog missing required column: ' + requiredPrimaryCols[i]);
      }
    }

    // Required mission columns
    var requiredMissionCols = getAllMissionIds_();

    for (var i = 0; i < requiredMissionCols.length; i++) {
      if (!headerSet.has(requiredMissionCols[i])) {
        result.pass = false;
        result.details.push('MissionLog missing mission column: ' + requiredMissionCols[i]);
      }
    }

    // Short-circuit if headers are badly broken
    if (!result.pass) {
      return result;
    }

    // -----------------------------------------------------------------------
    // 3) Row alignment: PreferredNames vs MissionLog
    // -----------------------------------------------------------------------
    var prefLastRow = prefSheet.getLastRow();
    var missionLastRow = missionSheet.getLastRow();

    if (prefLastRow > 1) {
      // Get PreferredNames IDs (column A, row 2+)
      var prefValues = prefSheet.getRange(2, 1, prefLastRow - 1, 1).getValues();
      var prefIds = new Set();

      for (var i = 0; i < prefValues.length; i++) {
        var id = String(prefValues[i][0]).trim();
        if (id) prefIds.add(id);
      }

      // Get MissionLog IDs (column A, row 2+)
      var missionValues = missionLastRow > 1
        ? missionSheet.getRange(2, 1, missionLastRow - 1, 1).getValues()
        : [];

      var missionIds = new Set();
      for (var i = 0; i < missionValues.length; i++) {
        var id = String(missionValues[i][0]).trim();
        if (id) missionIds.add(id);
      }

      // Players in PreferredNames but not in MissionLog
      var missingInMission = [];
      var prefIterator = prefIds.values();
      var prefResult = prefIterator.next();

      while (!prefResult.done) {
        var id = prefResult.value;
        if (!missionIds.has(id)) {
          missingInMission.push(id);
        }
        prefResult = prefIterator.next();
      }

      if (missingInMission.length > 0) {
        result.pass = false;
        var displayCount = Math.min(10, missingInMission.length);
        var sampleIds = missingInMission.slice(0, displayCount).join(', ');
        var moreText = missingInMission.length > displayCount
          ? ' (and ' + (missingInMission.length - displayCount) + ' more)'
          : '';

        result.details.push(
          missingInMission.length + ' player(s) missing MissionLog rows: ' + sampleIds + moreText
        );
      }
    }

    // -----------------------------------------------------------------------
    // 4) Smoke test: attendance scan for last 7 days
    // -----------------------------------------------------------------------
    var today = new Date();
    var startDate = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);

    var records;
    try {
      records = scanAttendanceForRange_(startDate, today);
    } catch (e) {
      result.pass = false;
      result.details.push('Attendance scan failed: ' + e.message);
      return result;
    }

    // -----------------------------------------------------------------------
    // 5) Consistency: everyone seen in attendance has a MissionLog row
    // -----------------------------------------------------------------------
    if (records && records.length > 0) {
      var missionLastRow2 = missionSheet.getLastRow();
      var missionValues2 = missionLastRow2 > 1
        ? missionSheet.getRange(2, 1, missionLastRow2 - 1, 1).getValues()
        : [];

      var missionIds2 = new Set();
      for (var i = 0; i < missionValues2.length; i++) {
        var id = String(missionValues2[i][0]).trim();
        if (id) missionIds2.add(id);
      }

      var attendanceOnlyIds = new Set();
      for (var i = 0; i < records.length; i++) {
        if (records[i].playerId) {
          attendanceOnlyIds.add(records[i].playerId);
        }
      }

      var attendanceMissingMission = [];
      var attendanceIterator = attendanceOnlyIds.values();
      var attendanceResult = attendanceIterator.next();

      while (!attendanceResult.done) {
        var id = attendanceResult.value;
        if (!missionIds2.has(id)) {
          attendanceMissingMission.push(id);
        }
        attendanceResult = attendanceIterator.next();
      }

      if (attendanceMissingMission.length > 0) {
        result.pass = false;
        var displayCount2 = Math.min(10, attendanceMissingMission.length);
        var sampleIds2 = attendanceMissingMission.slice(0, displayCount2).join(', ');
        var moreText2 = attendanceMissingMission.length > displayCount2
          ? ' (and ' + (attendanceMissingMission.length - displayCount2) + ' more)'
          : '';

        result.details.push(
          attendanceMissingMission.length + ' recent attendee(s) without MissionLog rows: ' + sampleIds2 + moreText2
        );
      }
    }

    // -----------------------------------------------------------------------
    // 6) Final status
    // -----------------------------------------------------------------------
    if (result.pass) {
      result.details.push(
        'MissionLog schema, player coverage, and attendance scan all healthy'
      );
    }

  } catch (err) {
    result.pass = false;
    result.details.push('Gate I error: ' + err.message);
    Logger.log('Gate I threw an error: ' + err);
  }

  return result;
}

// ══════════════════════════════════════════════════════════════════════
// GATE I: AUTO-FIX
// ══════════════════════════════════════════════════════════════════════

/**
 * Fixes Gate I issues by:
 *  - Creating MissionLog sheet if missing
 *  - Adding required columns
 *  - Syncing PreferredNames players to MissionLog
 *
 * @return {string} Fix message
 */
function fixGateI_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // 1. Ensure MissionLog sheet exists
    var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);
    if (!missionSheet) {
      missionSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);
    }

    // 2. Ensure schema is correct
    ensureMissionLogSchema_(missionSheet);

    // 3. Sync all PreferredNames players to MissionLog
    var syncStats = syncAllPlayersToMissionLog_();

    return 'MissionLog repaired: ' + syncStats.created + ' rows created, ' + syncStats.existing + ' existing';

  } catch (e) {
    Logger.log('Failed to fix Gate I: ' + e);
    return 'Fix failed: ' + e.message;
  }
}

/**
 * Ensure MissionLog has correct schema
 * @param {Sheet} sheet - MissionLog sheet
 * @private
 */
function ensureMissionLogSchema_(sheet) {
  if (!sheet) return;

  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  // If sheet is empty, create headers
  if (lastRow === 0) {
    var headers = ['preferred_name_id', 'Total_Points'].concat(getAllMissionIds_());
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return;
  }

  // Check if headers exist and add missing ones
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerSet = new Set();
  for (var i = 0; i < headerRow.length; i++) {
    headerSet.add(String(headerRow[i]).trim());
  }

  var requiredHeaders = ['preferred_name_id', 'Total_Points'].concat(getAllMissionIds_());
  var missingHeaders = [];

  for (var i = 0; i < requiredHeaders.length; i++) {
    if (!headerSet.has(requiredHeaders[i])) {
      missingHeaders.push(requiredHeaders[i]);
    }
  }

  if (missingHeaders.length > 0) {
    // Add missing headers to the end
    var newColStart = lastCol + 1;
    for (var i = 0; i < missingHeaders.length; i++) {
      sheet.getRange(1, newColStart + i).setValue(missingHeaders[i]);
    }
  }

  // Ensure frozen rows
  if (sheet.getFrozenRows() === 0) {
    sheet.setFrozenRows(1);
  }
}

// ══════════════════════════════════════════════════════════════════════
// MISSIONLOG MANAGEMENT
// ══════════════════════════════════════════════════════════════════════

/**
 * Create MissionLog sheet from scratch with proper schema
 * @return {Sheet} Created sheet
 */
function createMissionLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing MissionLog if present
  var existing = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);
  if (existing) {
    ss.deleteSheet(existing);
  }

  // Create new sheet
  var sheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);

  // Add headers
  var headers = [
    'preferred_name_id',
    'Total_Points',
    'ATTEND_CMD_CASUAL',
    'ATTEND_CMD_TRANSITION',
    'ATTEND_CMD_CEDH',
    'ATTEND_LIMITED_EVENT',
    'ATTEND_ACADEMY',
    'ATTEND_OUTREACH'
  ];

  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  // Format header row
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4A86E8');
  headerRange.setFontColor('#FFFFFF');

  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  Logger.log('MissionLog sheet created with canonical schema');

  return sheet;
}

/**
 * Initialize MissionLog with all PreferredNames players
 * @return {number} Number of players added
 */
function initializeMissionLog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prefSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.PREFERRED_NAMES);

  if (!prefSheet) {
    throw new Error('PreferredNames sheet not found');
  }

  // Ensure MissionLog exists
  var missionSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.MISSION_LOG);
  if (!missionSheet) {
    missionSheet = createMissionLogSheet_();
  }

  // Sync all players
  var syncStats = syncAllPlayersToMissionLog_();

  logIntegrityAction('INIT_MISSIONLOG', {
    details: 'Initialized MissionLog: ' + syncStats.created + ' created, ' + syncStats.existing + ' existing',
    status: 'SUCCESS'
  });

  return syncStats.created;
}