/**
 * Tab Organizer Service for Cosmic Event Manager v7.9.6
 *
 * Organizes sheet tabs into logical groups:
 * 1. Non-event tabs sorted alphabetically (left, A–Z)
 * 2. Event tabs sorted by date (right, chronological)
 *
 * FEATURES:
 * - Trims whitespace from tab names before pattern matching
 * - Case-insensitive letter suffix matching
 * - Comprehensive logging to identify stragglers
 * - Integrity logging for audit trail
 *
 * @version 1.0.0
 * @author Cosmic Event Manager Team
 */

/**
 * Event tab name pattern (case-insensitive):
 * - MM-DD-YYYY or M-D-YYYY (1-2 digits for month/day)
 * - MM-DDX-YYYY (X = letter, any case)
 * Each may optionally have a trailing -N duplicate suffix.
 *
 * Examples that match:
 *   01-10-2025, 1-10-2025, 09-6C-2025, 10-4c-2025, 11-1C-2025
 */
const EVENT_NAME_REGEX = /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})(-\d+)?$/;

/**
 * Parses a Date from an event tab name.
 * TRIMS whitespace and handles case-insensitive letter suffixes.
 * Accepts 1-2 digits for month and day.
 *
 * Examples:
 *   "01-10-2025"      → Date(2025, 0, 10)
 *   "1-10-2025"       → Date(2025, 0, 10)   // Single-digit month!
 *   "09-6C-2025"      → Date(2025, 8, 6)    // Single-digit day!
 *   "10-4c-2025"      → Date(2025, 9, 4)    // Lowercase letter!
 *   "01-10C-2025"     → Date(2025, 0, 10)
 *   "01-10-2025 "     → Date(2025, 0, 10)   // Trimmed!
 *   " 01-10-2025"     → Date(2025, 0, 10)   // Trimmed!
 *   "PreferredNames"  → null
 *
 * @param {string} name - The sheet name to parse
 * @return {Date|null} Parsed Date object or null if not an event tab
 * @private
 */
function parseEventTabDate_(name) {
  // CRITICAL FIX: Trim whitespace before matching
  const trimmed = name.trim();
  const match = trimmed.match(EVENT_NAME_REGEX);

  if (!match) {
    return null;
  }

  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);
  const year = parseInt(match[4], 10);

  // Validate date ranges
  if (month < 1 || month > 12) {
    Logger.log('WARNING: Invalid month in: ' + name + ' (month=' + month + ')');
    return null;
  }
  if (day < 1 || day > 31) {
    Logger.log('WARNING: Invalid day in: ' + name + ' (day=' + day + ')');
    return null;
  }

  const date = new Date(year, month - 1, day);
  if (isNaN(date.getTime())) {
    Logger.log('WARNING: Invalid date calculation for: ' + name);
    return null;
  }

  return date;
}

/**
 * Normalizes a sheet name for alphabetical sorting:
 * - Trims whitespace
 * - Strips leading non-alphanumeric symbols (emojis, bullets, etc.)
 * - Lowercases the remaining string
 *
 * @param {string} name - Sheet name to normalize
 * @return {string} Normalized name for sorting
 * @private
 */
function normalizeNameForSort_(name) {
  // Trim first
  const trimmed = name.trim();
  // Remove leading non-alphanumeric chars like emojis, punctuation, spaces
  const stripped = trimmed.replace(/^[^A-Za-z0-9]+/, '');
  return stripped.toLowerCase();
}

/**
 * Applies the desired sheet order by moving sheets to their target positions.
 *
 * Works BACKWARDS (right to left) and moves every sheet unconditionally.
 * This avoids issues with checking current positions during index shifts.
 *
 * @param {Array} finalOrder - Array of sheet entry objects with .sheet property
 * @private
 */
function applySheetOrder_(finalOrder) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log('Applying new sheet order (backwards iteration)...');

  // Work BACKWARDS from rightmost to leftmost position
  for (var i = finalOrder.length - 1; i >= 0; i--) {
    const entry = finalOrder[i];
    const sheet = entry.sheet;
    const targetPosition = i + 1; // Apps Script uses 1-based indexing

    Logger.log('  Moving "' + sheet.getName() + '" to position ' + targetPosition);

    // Always move, don't check if already in position (indices shift during moves)
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(targetPosition);
  }

  Logger.log('Sheet reordering complete');
}

/**
 * Organizes all tabs in the spreadsheet into two groups:
 * 1. Non-event tabs (left, sorted alphabetically)
 * 2. Event tabs (right, sorted by date ascending)
 *
 * This is the core service function called by UI handlers.
 * Logs all operations to both console and integrity log.
 *
 * @return {Object} Summary object with counts and message
 */
function organizeTabsLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  Logger.log('===============================================');
  Logger.log('STARTING TAB ORGANIZATION');
  Logger.log('===============================================');
  Logger.log('Total sheets found: ' + sheets.length);
  Logger.log('');

  const eventSheets = [];
  const nonEventSheets = [];
  const stragglers = []; // Track potential problem tabs

  sheets.forEach(function(sheet) {
    const name = sheet.getName();
    const date = parseEventTabDate_(name);

    if (date !== null) {
      // It's an event tab
      eventSheets.push({
        sheet: sheet,
        name: name,
        date: date
      });
      Logger.log('EVENT: "' + name + '" -> ' + date.toISOString().split('T')[0]);
    } else {
      // Not an event tab
      nonEventSheets.push({
        sheet: sheet,
        name: name
      });
      Logger.log('NON-EVENT: "' + name + '"');

      // Check if it LOOKS like an event tab but didn't match
      const trimmed = name.trim();
      if (/\d{2}-\d{2}/.test(trimmed)) {
        stragglers.push(name);
        Logger.log('   WARNING: Contains date pattern but did not match regex!');
      }
    }
  });

  Logger.log('');
  Logger.log('---------------------------------------------');
  Logger.log('CLASSIFICATION RESULTS:');
  Logger.log('   Non-event tabs: ' + nonEventSheets.length);
  Logger.log('   Event tabs: ' + eventSheets.length);
  if (stragglers.length > 0) {
    Logger.log('   WARNING: Potential stragglers: ' + stragglers.length);
    stragglers.forEach(function(s) {
      Logger.log('      -> "' + s + '"');
    });
  }
  Logger.log('---------------------------------------------');
  Logger.log('');

  // 1) Non-events: alphabetical by normalized name
  Logger.log('Sorting non-event tabs alphabetically...');
  nonEventSheets.sort(function(a, b) {
    const an = normalizeNameForSort_(a.name);
    const bn = normalizeNameForSort_(b.name);
    return an.localeCompare(bn);
  });

  // 2) Events: date ascending, then name
  Logger.log('Sorting event tabs by date...');
  eventSheets.sort(function(a, b) {
    const dateDiff = a.date - b.date;
    if (dateDiff !== 0) {
      return dateDiff;
    }
    return a.name.localeCompare(b.name);
  });

  // Final layout: NON-EVENTS -> EVENTS
  const finalOrder = nonEventSheets.concat(eventSheets);

  Logger.log('');
  Logger.log('FINAL ORDER:');
  finalOrder.forEach(function(entry, idx) {
    const pos = idx + 1;
    const isEvent = entry.date !== undefined;
    const prefix = isEvent ? 'EVENT' : 'NON-EVENT';
    Logger.log('   ' + pos + '. ' + prefix + ' "' + entry.name + '"');
  });
  Logger.log('');

  // Apply the new order
  applySheetOrder_(finalOrder);

  // Log to integrity log
  try {
    logIntegrityAction(
      'TAB_ORGANIZE',
      'SYSTEM',
      'Organized ' + nonEventSheets.length + ' non-event tab(s) and ' +
      eventSheets.length + ' event tab(s)' +
      (stragglers.length > 0 ? '; ' + stragglers.length + ' stragglers found' : '')
    );
  } catch (e) {
    Logger.log('WARNING: Could not log to integrity log: ' + e.message);
  }

  Logger.log('===============================================');
  Logger.log('ORGANIZATION COMPLETE');
  Logger.log('===============================================');

  return {
    success: true,
    nonEventCount: nonEventSheets.length,
    eventCount: eventSheets.length,
    stragglerCount: stragglers.length,
    stragglers: stragglers,
    message: 'Organized ' + nonEventSheets.length + ' non-event tab(s) and ' +
             eventSheets.length + ' event tab(s).' +
             (stragglers.length > 0 ? '\n\nWARNING: Found ' + stragglers.length + ' potential straggler(s). Check logs.' : '')
  };
}

/**
 * Gets current tab organization status without making changes.
 * Used for diagnostic purposes and health checks.
 *
 * @return {Object} Status object with classification details
 */
function getTabOrganizationStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const eventSheets = [];
  const nonEventSheets = [];
  const stragglers = [];

  sheets.forEach(function(sheet) {
    const name = sheet.getName();
    const date = parseEventTabDate_(name);

    if (date !== null) {
      eventSheets.push({
        name: name,
        date: date.toISOString().split('T')[0]
      });
    } else {
      nonEventSheets.push({
        name: name
      });

      // Check for stragglers
      const trimmed = name.trim();
      if (/\d{2}-\d{2}/.test(trimmed)) {
        stragglers.push(name);
      }
    }
  });

  return {
    totalSheets: sheets.length,
    nonEventCount: nonEventSheets.length,
    eventCount: eventSheets.length,
    stragglerCount: stragglers.length,
    nonEventSheets: nonEventSheets,
    eventSheets: eventSheets,
    stragglers: stragglers
  };
}

/**
 * DEBUG: Shows which tabs are events vs non-events without moving them.
 * Use this to identify problem tabs before organizing.
 *
 * @return {string} Classification report text
 */
function debugShowTabClassification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let report = 'TAB CLASSIFICATION REPORT\n';
  report += '=============================\n\n';

  let eventCount = 0;
  let nonEventCount = 0;
  let stragglers = [];

  sheets.forEach(function(sheet, idx) {
    const name = sheet.getName();
    const date = parseEventTabDate_(name);

    if (date !== null) {
      report += (idx + 1) + '. EVENT: "' + name + '"\n';
      report += '   -> ' + date.toISOString().split('T')[0] + '\n';
      eventCount++;
    } else {
      report += (idx + 1) + '. NON-EVENT: "' + name + '"\n';
      nonEventCount++;

      // Check if it looks like an event but didn't match
      const trimmed = name.trim();
      if (/\d{2}-\d{2}/.test(trimmed)) {
        report += '   WARNING: Contains date pattern but did not match!\n';
        stragglers.push(name);
      }
    }
  });

  report += '\n-----------------------------\n';
  report += 'SUMMARY:\n';
  report += '  Non-events: ' + nonEventCount + '\n';
  report += '  Events: ' + eventCount + '\n';
  report += '  Stragglers: ' + stragglers.length + '\n';

  if (stragglers.length > 0) {
    report += '\nWARNING: STRAGGLERS (check these):\n';
    stragglers.forEach(function(s) {
      report += '  - "' + s + '"\n';
    });
  }

  Logger.log(report);
  return report;
}

/**
 * Validates tab structure health for Ship-Gates integration.
 * Checks for common issues that might prevent proper organization.
 *
 * @return {Object} Health check result with pass/fail status
 */
function validateTabStructure() {
  const status = getTabOrganizationStatus();
  const issues = [];

  // Check for stragglers
  if (status.stragglerCount > 0) {
    issues.push('Found ' + status.stragglerCount + ' tab(s) with date patterns that do not match naming convention: ' +
                status.stragglers.join(', '));
  }

  // Check for required system sheets
  const requiredSheets = [
    'Prize_Catalog',
    'Prize_Throttle',
    'Integrity_Log',
    'Spent_Pool',
    'Key_Tracker',
    'BP_Total'
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  requiredSheets.forEach(function(sheetName) {
    if (!ss.getSheetByName(sheetName)) {
      issues.push('Required system sheet missing: ' + sheetName);
    }
  });

  return {
    pass: issues.length === 0,
    issues: issues,
    status: status
  };
}