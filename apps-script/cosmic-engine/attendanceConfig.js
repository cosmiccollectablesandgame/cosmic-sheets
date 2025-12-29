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
 * Canonical mission IDs
 * These MUST match the Mission_ID column in MissionLog sheet
 */
var MISSION_IDS = {
  // Attendance missions (auto-computed)
  ATTEND_001: 'ATTEND_001',
  ATTEND_010: 'ATTEND_010',
  ATTEND_020: 'ATTEND_020',
  ATTEND_030: 'ATTEND_030',
  ATTEND_040: 'ATTEND_040',
  ATTEND_050: 'ATTEND_050',
  ATTEND_060: 'ATTEND_060',
  ATTEND_070: 'ATTEND_070',
  ATTEND_080: 'ATTEND_080',
  ATTEND_090: 'ATTEND_090',

  // Flag missions (staff-verified)
  FLAG_010: 'FLAG_010',
  FLAG_020: 'FLAG_020',
  FLAG_030: 'FLAG_030',
  FLAG_040: 'FLAG_040',
  FLAG_050: 'FLAG_050',
  FLAG_060: 'FLAG_060',
  FLAG_070: 'FLAG_070',
  FLAG_100: 'FLAG_100'
};

/**
 * Get all mission IDs
 * @return {Array<string>} Array of mission IDs
 */
function getAllMissionIds_() {
  return Object.keys(MISSION_IDS).map(function(key) { return MISSION_IDS[key]; });
}

/**
 * Default mission definitions (embedded in code)
 * Matches official Cosmic Games mission spec
 * These are used if MissionLog_1/MissionLog_2 sheets don't exist
 */
var MISSION_DEFINITIONS = {
  // ═══════════════════════════════════════════════════════════════════
  // ATTENDANCE MISSIONS (Auto-computed from event sheets)
  // ═══════════════════════════════════════════════════════════════════

  'ATTEND_001': {
    id: 'ATTEND_001',
    name: 'First Contact',
    type: 'first_event',
    criteria: '1',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_010': {
    id: 'ATTEND_010',
    name: 'Stellar Explorer',
    type: 'event_count',
    criteria: '5',
    pointsValue: 2,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_020': {
    id: 'ATTEND_020',
    name: 'Deck Diver',
    type: 'format_diversity',
    criteria: '', // Award 1 BP per unique suffix played
    pointsValue: 1,
    cap: 0, // Unlimited
    oneTime: false,
    active: true
  },

  'ATTEND_030': {
    id: 'ATTEND_030',
    name: 'Lunar Loyalty',
    type: 'monthly_attendance',
    criteria: '4', // 4 events in one month
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_040': {
    id: 'ATTEND_040',
    name: 'Sealed Voyager',
    type: 'format_specific',
    criteria: 'R,S', // Prerelease Sealed or Sealed
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_050': {
    id: 'ATTEND_050',
    name: 'Draft Navigator',
    type: 'format_specific',
    criteria: 'D', // Draft
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_060': {
    id: 'ATTEND_060',
    name: 'Stellar Scholar',
    type: 'format_specific',
    criteria: 'W', // Workshop / Hobby Night
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_070': {
    id: 'ATTEND_070',
    name: 'Meteor Shower',
    type: 'weekly_attendance',
    criteria: '2', // 2 events in one week
    pointsValue: 2,
    cap: 1,
    oneTime: true,
    active: true
  },

  'ATTEND_080': {
    id: 'ATTEND_080',
    name: 'Interstellar Strategist',
    type: 'placement',
    criteria: '4', // Top 4 finish
    pointsValue: 1,
    cap: 0, // Unlimited
    oneTime: false,
    active: true
  },

  'ATTEND_090': {
    id: 'ATTEND_090',
    name: 'Black Hole Survivor',
    type: 'last_place',
    criteria: '', // Last place finish
    pointsValue: 1,
    cap: 0, // Unlimited
    oneTime: false,
    active: true
  },

  // ═══════════════════════════════════════════════════════════════════
  // FLAG MISSIONS (Staff-verified, not auto-computed)
  // ═══════════════════════════════════════════════════════════════════
  // These are tracked in MissionLog but not computed by attendance scanner

  'FLAG_010': {
    id: 'FLAG_010',
    name: 'Cosmic Selfie',
    type: 'flag',
    criteria: 'manual', // Staff-verified
    pointsValue: 2,
    cap: 1,
    oneTime: true,
    active: false // Not computed by scanner
  },

  'FLAG_020': {
    id: 'FLAG_020',
    name: 'Social Media Star',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: false
  },

  'FLAG_030': {
    id: 'FLAG_030',
    name: 'App Explorer',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 2,
    cap: 1,
    oneTime: true,
    active: false
  },

  'FLAG_040': {
    id: 'FLAG_040',
    name: 'Cosmic Merchant',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: false
  },

  'FLAG_050': {
    id: 'FLAG_050',
    name: 'Precon Pioneer',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: false
  },

  'FLAG_060': {
    id: 'FLAG_060',
    name: 'Gravitational Pull',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 0, // Unlimited
    oneTime: false,
    active: false
  },

  'FLAG_070': {
    id: 'FLAG_070',
    name: 'Rogue Planet',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: false
  },

  'FLAG_100': {
    id: 'FLAG_100',
    name: 'Quantum Collector',
    type: 'flag',
    criteria: 'manual',
    pointsValue: 1,
    cap: 1,
    oneTime: true,
    active: false
  }
};

/**
 * Get default mission definitions
 * @return {Object} Mission definitions object
 */
function getDefaultMissionDefinitions_() {
  // Return only active missions
  var activeMissions = {};
  for (var id in MISSION_DEFINITIONS) {
    if (MISSION_DEFINITIONS[id].active) {
      activeMissions[id] = MISSION_DEFINITIONS[id];
    }
  }
  return activeMissions;
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
    MISSION_LOG: 'MissionLog',
    PREFERRED_NAMES: 'PreferredNames',
    INTEGRITY_LOG: 'Integrity_Log'
  },

  // Event sheet name pattern: MM-DD-YYYY or MM-DDX-YYYY
  // Examples: 11-23-2025 (standard) or 11-23B-2025 (Commander Casual)
  EVENT_PATTERN: /^(\d{2})-(\d{2})([A-Z])?\-(\d{4})$/,

  // Format suffix legend (A-Z complete, v7.9.6+ official standard)
  FORMAT_LEGEND: {
    'A': 'Academy / Learn to Play',
    'B': 'Casual Commander (Brackets 1-2)',
    'C': 'Transitional Commander (Brackets 3-4)',
    'D': 'Booster Draft',
    'E': 'External / Outreach',
    'F': 'Free Play Event',
    'G': 'Gundam / Gunpla',
    'H': 'Historic / Legacy MTG',
    'I': 'Yu-Gi-Oh! TCG',
    'J': 'Junior / Youth Events',
    'K': 'Kill Team',
    'L': 'Commander League',
    'M': 'Modern Constructed',
    'N': 'Pokémon TCG',
    'O': 'One Piece TCG',
    'P': 'Proxy / Cube Draft',
    'Q': 'Precon Event',
    'R': 'Prerelease Sealed',
    'S': 'Sealed',
    'T': 'cEDH / High-Power Commander (Bracket 5)',
    'U': 'Star Wars: Unlimited',
    'V': 'Rift (custom / special format)',
    'W': 'Workshop / Hobby Night',
    'X': 'Multi-Event Day',
    'Y': 'Lorcana TCG',
    'Z': 'Staff / Internal Use'
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
 * Complete A-Z legend (v7.9.6+ official standard)
 */
var SUFFIX_META = {
  'A': { name: 'Academy / Learn to Play', category: 'educational' },
  'B': { name: 'Casual Commander (Brackets 1-2)', category: 'commander' },
  'C': { name: 'Transitional Commander (Brackets 3-4)', category: 'commander' },
  'D': { name: 'Booster Draft', category: 'limited' },
  'E': { name: 'External / Outreach', category: 'special' },
  'F': { name: 'Free Play Event', category: 'casual' },
  'G': { name: 'Gundam / Gunpla', category: 'other_tcg' },
  'H': { name: 'Historic / Legacy MTG', category: 'constructed' },
  'I': { name: 'Yu-Gi-Oh! TCG', category: 'other_tcg' },
  'J': { name: 'Junior / Youth Events', category: 'educational' },
  'K': { name: 'Kill Team', category: 'miniatures' },
  'L': { name: 'Commander League', category: 'commander' },
  'M': { name: 'Modern Constructed', category: 'constructed' },
  'N': { name: 'Pokémon TCG', category: 'other_tcg' },
  'O': { name: 'One Piece TCG', category: 'other_tcg' },
  'P': { name: 'Proxy / Cube Draft', category: 'limited' },
  'Q': { name: 'Precon Event', category: 'commander' },
  'R': { name: 'Prerelease Sealed', category: 'limited' },
  'S': { name: 'Sealed', category: 'limited' },
  'T': { name: 'cEDH / High-Power Commander (Bracket 5)', category: 'commander' },
  'U': { name: 'Star Wars: Unlimited', category: 'other_tcg' },
  'V': { name: 'Riftbound ', category: 'Riftbound' },
  'W': { name: 'Workshop / Hobby Night', category: 'educational' },
  'X': { name: 'Multi-Event Day', category: 'special' },
  'Y': { name: 'Lorcana TCG', category: 'other_tcg' },
  'Z': { name: 'Staff / Internal Use', category: 'internal' }
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