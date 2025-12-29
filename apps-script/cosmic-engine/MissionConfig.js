/**
 * ════════════════════════════════════════════════════════════════════════
 * COSMIC MISSION & ATTENDANCE SYSTEM - CONFIGURATION
 *
 * Central configuration for all mission and attendance tracking systems
 * Compatible with: Engine v7.9.6+
 * Version: 1.0.0
 * ════════════════════════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════════════
// MISSION CONSTANTS
// ══════════════════════════════════════════════════════════════════════

/**
 * Canonical mission IDs for suffix-based attendance missions
 * These MUST match the column headers in MissionLog sheet
 */
var MISSION_IDS = {
  ATTEND_CMD_CASUAL: 'ATTEND_CMD_CASUAL',
  ATTEND_CMD_TRANSITION: 'ATTEND_CMD_TRANSITION',
  ATTEND_CMD_CEDH: 'ATTEND_CMD_CEDH',
  ATTEND_LIMITED_EVENT: 'ATTEND_LIMITED_EVENT',
  ATTEND_ACADEMY: 'ATTEND_ACADEMY',
  ATTEND_OUTREACH: 'ATTEND_OUTREACH'
};

/**
 * Get all mission IDs
 * @return {Array<string>} Array of mission IDs
 */
function getAllMissionIds_() {
  return Object.keys(MISSION_IDS).map(function(key) { return MISSION_IDS[key]; });
}

// ══════════════════════════════════════════════════════════════════════
// ATTENDANCE TRACKING CONFIGURATION
// ══════════════════════════════════════════════════════════════════════

var ATTENDANCE_CONFIG = {
  SHEETS: {
    CALENDAR: 'Attendance_Calendar',
    MISSIONS: 'Attendance_Missions',
    MISSION_LOG_1: 'MissionLog_1',
    MISSION_LOG_2: 'MissionLog_2',
    PREFERRED_NAMES: 'PreferredNames',
    INTEGRITY_LOG: 'Integrity_Log',
    MISSION_LOG: 'MissionLog'
  },

  // Event sheet name pattern: MM-DD-YYYY or MM-DD-[SUFFIX]-YYYY
  EVENT_PATTERN: /^(\d{2})-(\d{2})(?:-([A-Z]))?\-(\d{4})$/,

  // Format suffix legend
  FORMAT_LEGEND: {
    'C': 'Commander',
    'D': 'Draft',
    'S': 'Sealed',
    'P': 'Prerelease',
    'M': 'Modern',
    'L': 'Legacy',
    'V': 'Vintage',
    'N': 'Night Event',
    'B': 'Brawl',
    'T': 'Two-Headed Giant',
    'E': 'Outreach',
    'A': 'Academy'
  },

  // Column headers expected in event sheets
  EXPECTED_HEADERS: {
    PLAYER: 'player',
    STANDING: 'final standing'
  },

  // Performance settings
  BATCH_SIZE: 1000, // Max rows to write at once
  MAX_WEEKS_FOR_STREAK: 52 // Maximum consecutive weeks to check
};

/**
 * Suffix metadata for mission evaluation
 * Used by Mission Suffix Service to determine which missions to award
 */
var SUFFIX_META = {
  'B': {
    name: 'Commander Casual',
    requiresKitPrompt: false,
    missionIds: ['ATTEND_CMD_CASUAL']
  },
  'C': {
    name: 'Commander Transition',
    requiresKitPrompt: false,
    missionIds: ['ATTEND_CMD_TRANSITION']
  },
  'T': {
    name: 'Commander cEDH',
    requiresKitPrompt: false,
    missionIds: ['ATTEND_CMD_CEDH']
  },
  'D': {
    name: 'Draft',
    requiresKitPrompt: true,
    missionIds: ['ATTEND_LIMITED_EVENT']
  },
  'P': {
    name: 'Proxy/Cube Draft',
    requiresKitPrompt: true,
    missionIds: ['ATTEND_LIMITED_EVENT']
  },
  'R': {
    name: 'Prerelease',
    requiresKitPrompt: true,
    missionIds: ['ATTEND_LIMITED_EVENT']
  },
  'S': {
    name: 'Sealed',
    requiresKitPrompt: true,
    missionIds: ['ATTEND_LIMITED_EVENT']
  },
  'A': {
    name: 'Academy',
    requiresKitPrompt: false,
    missionIds: ['ATTEND_ACADEMY']
  },
  'E': {
    name: 'Outreach',
    requiresKitPrompt: false,
    missionIds: ['ATTEND_OUTREACH']
  }
};

/**
 * Get suffix metadata
 * @param {string} suffix - Event suffix (e.g., 'B', 'C', 'D')
 * @return {Object|null} Suffix metadata or null
 */
function getSuffixMeta_(suffix) {
  if (!suffix) return null;
  return SUFFIX_META[suffix] || null;
}

/**
 * Extract suffix from event ID (sheet/tab name)
 * @param {string} eventId - Event sheet name (e.g., "11-23-B-2025")
 * @return {string} Suffix or empty string
 */
function getSuffixFromEventId_(eventId) {
  if (!eventId) return '';
  var match = eventId.match(ATTENDANCE_CONFIG.EVENT_PATTERN);
  return match && match[3] ? match[3] : '';
}