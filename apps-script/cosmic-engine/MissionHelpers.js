/**
 * ════════════════════════════════════════════════════════════════════════
 * COSMIC MISSION & ATTENDANCE SYSTEM - HELPER UTILITIES
 *
 * Shared utility functions used across all mission and attendance systems
 * Compatible with: Engine v7.9.6+
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════
// NUMBER UTILITIES
// ══════════════════════════════════════════════════════════════════════

/**
 * Coerce value to number with fallback
 * @param {*} value - Value to coerce
 * @param {number} fallback - Fallback value
 * @return {number} Coerced number
 * @private
 */
function coerceNumber(value, fallback) {
  if (value === null || value === undefined || value === '') {
    return fallback;
  }

  var num = Number(value);
  return isNaN(num) ? fallback : num;
}

// ══════════════════════════════════════════════════════════════════════
// PLAYER NAME RESOLUTION
// ══════════════════════════════════════════════════════════════════════

/**
 * Load PreferredNames for canonical name resolution
 * @param {Spreadsheet} ss - Spreadsheet object
 * @return {Set} Set of canonical player names
 */
function loadPreferredNames(ss) {
  var sheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.PREFERRED_NAMES);
  if (!sheet) {
    Logger.log('Warning: PreferredNames sheet not found. Name resolution may be inconsistent.');
    return new Set();
  }

  try {
    var data = sheet.getDataRange().getValues();
    var names = new Set();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        names.add(String(data[i][0]).trim());
      }
    }

    Logger.log('Loaded ' + names.size + ' preferred names');
    return names;
  } catch (error) {
    Logger.log('Error loading PreferredNames: ' + error.toString());
    return new Set();
  }
}

/**
 * Resolve canonical name (fuzzy matching if needed)
 * @param {string} name - Player name to resolve
 * @param {Set} preferredNames - Set of canonical names
 * @return {string|null} Canonical name or null
 */
function resolveCanonicalName(name, preferredNames) {
  if (!name) return null;

  // Direct match
  if (preferredNames.has(name)) return name;

  // Case-insensitive match
  var lowerName = name.toLowerCase();
  var iterator = preferredNames.values();
  var result = iterator.next();

  while (!result.done) {
    var canonical = result.value;
    if (canonical.toLowerCase() === lowerName) {
      return canonical;
    }
    result = iterator.next();
  }

  // If PreferredNames is empty, accept all names
  if (preferredNames.size === 0) {
    return name;
  }

  // If not found, return as-is (will be flagged in UndiscoveredNames later)
  Logger.log('Warning: Player name "' + name + '" not found in PreferredNames');
  return name;
}

// ══════════════════════════════════════════════════════════════════════
// DATE & WEEK UTILITIES
// ══════════════════════════════════════════════════════════════════════

/**
 * Get ISO week number for a date
 * @param {Date} date - Date object
 * @return {number} ISO week number
 */
function getWeekNumber(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Get number of ISO weeks in a year
 * @param {number} year - Year
 * @return {number} Number of weeks
 */
function getWeeksInYear(year) {
  var d = new Date(Date.UTC(year, 11, 31));
  var weekNum = getWeekNumber(d);
  return weekNum === 1 ? 52 : weekNum;
}

// ══════════════════════════════════════════════════════════════════════
// LOGGING UTILITIES
// ══════════════════════════════════════════════════════════════════════

/**
 * Log an integrity action to Integrity_Log
 * @param {string} eventType - Event type (e.g., 'INIT_MISSIONLOG')
 * @param {Object} data - Event data {details, status, etc.}
 */
function logIntegrityAction(eventType, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);

    if (!logSheet) {
      logSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);
      logSheet.appendRow([
        'timestamp', 'user', 'event', 'target', 'details',
        'engine_version', 'seed', 'checksum', 'df_tags', 'rl_band'
      ]);
    }

    var timestamp = new Date().toISOString();
    var user = Session.getActiveUser().getEmail();
    var engineVersion = typeof ENGINE_VERSION !== 'undefined' ? ENGINE_VERSION : '7.9.6';

    logSheet.appendRow([
      timestamp,
      user,
      eventType,
      data.target || 'System',
      data.details || '',
      engineVersion,
      data.seed || '',
      data.checksum || '',
      data.df_tags || '',
      data.rl_band || ''
    ]);

    Logger.log('Logged integrity action: ' + eventType);
  } catch (error) {
    Logger.log('Error logging integrity action: ' + error.toString());
  }
}

/**
 * Log attendance scan summary to Integrity_Log
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Object} summary - Scan summary data
 */
function logAttendanceScan(ss, summary) {
  try {
    var logSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);

    if (!logSheet) {
      logSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);
      logSheet.appendRow([
        'timestamp', 'user', 'event', 'target', 'details',
        'engine_version', 'seed', 'checksum', 'df_tags', 'rl_band'
      ]);
    }

    var timestamp = new Date().toISOString();
    var user = Session.getActiveUser().getEmail();
    var engineVersion = typeof ENGINE_VERSION !== 'undefined' ? ENGINE_VERSION : '7.9.6';

    logSheet.appendRow([
      timestamp,
      user,
      'ATTENDANCE_SCAN',
      'System',
      JSON.stringify(summary),
      engineVersion,
      '',
      '',
      'ATTENDANCE',
      ''
    ]);

    Logger.log('Logged attendance scan: ' + JSON.stringify(summary));
  } catch (error) {
    Logger.log('Error logging attendance scan: ' + error.toString());
  }
}

/**
 * Log error to Integrity_Log
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {Error} error - Error object
 */
function logAttendanceError(ss, error) {
  try {
    var logSheet = ss.getSheetByName(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);

    if (!logSheet) {
      logSheet = ss.insertSheet(ATTENDANCE_CONFIG.SHEETS.INTEGRITY_LOG);
      logSheet.appendRow([
        'timestamp', 'user', 'event', 'target', 'details',
        'engine_version', 'seed', 'checksum', 'df_tags', 'rl_band'
      ]);
    }

    var timestamp = new Date().toISOString();
    var user = Session.getActiveUser().getEmail();
    var engineVersion = typeof ENGINE_VERSION !== 'undefined' ? ENGINE_VERSION : '7.9.6';

    logSheet.appendRow([
      timestamp,
      user,
      'ATTENDANCE_SCAN_ERROR',
      'System',
      error.toString() + '\n\nStack: ' + (error.stack || 'N/A'),
      engineVersion,
      '',
      '',
      'ERROR',
      ''
    ]);
  } catch (logError) {
    Logger.log('Failed to log error: ' + logError.toString());
  }
}