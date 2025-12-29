/**
 * ════════════════════════════════════════════════════════════════════════════
 * MISSION SCANNER SERVICE v1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Unified mission scanning and attendance tracking
 * 
 * FIXES:
 *   - Correct event pattern: MM-DD[suffix]-YYYY (e.g., "11-23C-2025")
 *   - Correct column headers: "Rank" + "PreferredName" or "preferred_name_id"
 *   - Code-defined missions (no sheet dependencies)
 *   - Provides scanAttendanceForRange_() for MissionGateService
 *   - Integrates with SUFFIX_MAP from SuffixService
 *
 * OUTPUTS:
 *   - Attendance_Missions: Player × Mission progress matrix
 *   - MissionLog: Canonical mission tracking (if enabled)
 *   - Integrity_Log: Audit trail
 *
 * Compatible with: Engine v7.9.6+, SuffixService v7.9.7+
 * ════════════════════════════════════════════════════════════════════════════
 */

// ════════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ════════════════════════════════════════════════════════════════════════════

const MISSION_SCANNER_CONFIG = {
  // Event sheet pattern: MM-DD-YYYY or MM-DDX-YYYY (suffix attached to day)
  // Examples: "11-23-2025", "11-23C-2025", "05-01B-2025"
  EVENT_PATTERN: /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/,
  
  // Acceptable player column headers (case-insensitive)
  PLAYER_COLUMNS: ['preferredname', 'preferred_name_id', 'player', 'name', 'player name'],
  
  // Acceptable rank column headers (case-insensitive)  
  RANK_COLUMNS: ['rank', 'standing', 'final standing', 'place', 'placement'],
  
  // Output sheets
  SHEETS: {
    ATTENDANCE_MISSIONS: 'Attendance_Missions',
    MISSION_LOG: 'MissionLog',
    PREFERRED_NAMES: 'PreferredNames',
    INTEGRITY_LOG: 'Integrity_Log'
  }
};

// ════════════════════════════════════════════════════════════════════════════
// CODE-DEFINED MISSION REGISTRY
// ════════════════════════════════════════════════════════════════════════════

/**
 * Mission definitions - Single source of truth (no sheet dependency)
 * 
 * Structure:
 *   id: Unique mission identifier (column header)
 *   name: Display name
 *   type: 'attendance' | 'placement' | 'format' | 'streak' | 'one_time'
 *   description: Human-readable description
 *   pointValue: BP awarded per completion
 *   cap: Max times earnable (0 = unlimited)
 *   trigger: Function or config that determines when mission is earned
 */
const MISSION_REGISTRY = {
  // ── ONE-TIME MISSIONS (earned once ever) ──────────────────────────────────
  FIRST_CONTACT: {
    id: 'FIRST_CONTACT',
    name: 'First Contact',
    type: 'one_time',
    description: 'Attend any event',
    pointValue: 1,
    cap: 1,
    trigger: { minEvents: 1 }
  },
  
  STELLAR_EXPLORER: {
    id: 'STELLAR_EXPLORER',
    name: 'Stellar Explorer',
    type: 'one_time',
    description: 'Attend 5 distinct events',
    pointValue: 2,
    cap: 1,
    trigger: { minEvents: 5 }
  },
  
  DECK_DIVER: {
    id: 'DECK_DIVER',
    name: 'Deck Diver',
    type: 'one_time',
    description: 'Play across 2+ event types (different suffixes)',
    pointValue: 1,
    cap: 1,
    trigger: { minFormats: 2 }
  },
  
  LUNAR_LOYALTY: {
    id: 'LUNAR_LOYALTY',
    name: 'Lunar Loyalty',
    type: 'one_time',
    description: '4 events within any 31-day window',
    pointValue: 2,
    cap: 1,
    trigger: { eventsInDays: { count: 4, days: 31 } }
  },
  
  METEOR_SHOWER: {
    id: 'METEOR_SHOWER',
    name: 'Meteor Shower',
    type: 'one_time',
    description: '2 events within 7 days',
    pointValue: 1,
    cap: 1,
    trigger: { eventsInDays: { count: 2, days: 7 } }
  },
  
  // ── FORMAT-SPECIFIC MISSIONS (one-time per format) ────────────────────────
  SEALED_VOYAGER: {
    id: 'SEALED_VOYAGER',
    name: 'Sealed Voyager',
    type: 'format',
    description: 'Play Sealed or Prerelease Sealed',
    pointValue: 1,
    cap: 1,
    trigger: { suffixes: ['S', 'R','D'] }
  },
  
  DRAFT_NAVIGATOR: {
    id: 'DRAFT_NAVIGATOR',
    name: 'Draft Navigator',
    type: 'format',
    description: 'Play Booster Draft',
    pointValue: 1,
    cap: 1,
    trigger: { suffixes: ['D'] }
  },
  
  STELLAR_SCHOLAR: {
    id: 'STELLAR_SCHOLAR',
    name: 'Stellar Scholar',
    type: 'format',
    description: 'Attend Workshop/Hobby Night or Academy',
    pointValue: 1,
    cap: 1,
    trigger: { suffixes: ['W', 'A'] }
  },
  
  // ── COMMANDER ATTENDANCE (cumulative) ─────────────────────────────────────
  ATTEND_CMD_CASUAL: {
    id: 'ATTEND_CMD_CASUAL',
    name: 'Casual Commander Events',
    type: 'attendance',
    description: 'Attend Casual Commander (Bracket 1-2) events',
    pointValue: 0, // Tracked but not auto-awarded BP
    cap: 0,
    trigger: { suffixes: ['B'] }
  },
  
  ATTEND_CMD_TRANSITION: {
    id: 'ATTEND_CMD_TRANSITION',
    name: 'Transitional Commander Events',
    type: 'attendance',
    description: 'Attend Transitional Commander (Bracket 3-4) events',
    pointValue: 0,
    cap: 0,
    trigger: { suffixes: ['C'] }
  },
  
  ATTEND_CMD_CEDH: {
    id: 'ATTEND_CMD_CEDH',
    name: 'cEDH Events',
    type: 'attendance',
    description: 'Attend cEDH (Bracket 5) events',
    pointValue: 0,
    cap: 0,
    trigger: { suffixes: ['T'] }
  },
  
  // ── LIMITED FORMAT ATTENDANCE (cumulative) ────────────────────────────────
  ATTEND_LIMITED_EVENT: {
    id: 'ATTEND_LIMITED_EVENT',
    name: 'Limited Events',
    type: 'attendance',
    description: 'Attend Limited format events (Draft, Sealed, Prerelease)',
    pointValue: 0,
    cap: 0,
    trigger: { suffixes: ['D', 'P', 'R', 'S'] }
  },
  
  // ── SPECIAL PROGRAM ATTENDANCE ────────────────────────────────────────────
  ATTEND_ACADEMY: {
    id: 'ATTEND_ACADEMY',
    name: 'Academy Events',
    type: 'attendance',
    description: 'Attend Academy / Learn to Play events',
    pointValue: 0,
    cap: 0,
    trigger: { suffixes: ['A'] }
  },
  
  ATTEND_OUTREACH: {
    id: 'ATTEND_OUTREACH',
    name: 'Outreach Events',
    type: 'attendance',
    description: 'Attend External / Outreach events',
    pointValue: 0,
    cap: 0,
    trigger: { suffixes: ['E'] }
  },
  
  ATTEND_FREE_PLAY: {
    id: 'ATTEND_FREE_PLAY',
    name: 'Free Play Events',
    type: 'attendance',
    description: 'Attend Free Play events',
    pointValue: 1, // 1 BP per free play attendance
    cap: 0,
    trigger: { suffixes: ['F'] }
  },
  
  // ── PLACEMENT MISSIONS (cumulative) ───────────────────────────────────────
  INTERSTELLAR_STRATEGIST: {
    id: 'INTERSTELLAR_STRATEGIST',
    name: 'Interstellar Strategist',
    type: 'placement',
    description: 'Finish Top 4 in an event',
    pointValue: 1,
    cap: 0, // Unlimited - 1 BP per Top 4 finish
    trigger: { maxRank: 4 }
  },
  
  BLACK_HOLE_SURVIVOR: {
    id: 'BLACK_HOLE_SURVIVOR',
    name: 'Black Hole Survivor',
    type: 'placement',
    description: 'Finish last in an event (when no ranks used)',
    pointValue: 1,
    cap: 0,
    trigger: { isLast: true }
  },
  
  // ── TOTAL ATTENDANCE (for tracking) ───────────────────────────────────────
  TOTAL_EVENTS: {
    id: 'TOTAL_EVENTS',
    name: 'Points',
    type: 'attendance',
    description: 'Total number of events attended',
    pointValue: 0,
    cap: 0,
    trigger: { all: true }
  }
};

/**
 * Get all mission IDs for sheet headers
 * @return {Array<string>} Mission IDs
 */
function getAllMissionIds_() {
  return Object.keys(MISSION_REGISTRY);
}

/**
 * Get mission definition by ID
 * @param {string} missionId - Mission ID
 * @return {Object|null} Mission definition
 */
function getMissionDef_(missionId) {
  return MISSION_REGISTRY[missionId] || null;
}

// ════════════════════════════════════════════════════════════════════════════
// MAIN SCANNING FUNCTION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Run full mission scan - Main entry point
 * 
 * Workflow:
 *   1. Discover all event sheets
 *   2. Extract attendance records (player, event, rank, suffix)
 *   3. Compute mission progress for all players
 *   4. Write to Attendance_Missions sheet
 *   5. Log to Integrity_Log
 *
 * @return {Object} Scan results {eventsScanned, playersTracked, missionsComputed}
 */
function runMissionScan() {
  const startTime = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    console.log('═══════════════════════════════════════════════════════');
    console.log('MISSION SCANNER v1.0.0 - Starting scan...');
    console.log('═══════════════════════════════════════════════════════');
    
    // Step 1: Scan all event sheets
    const scanData = scanAllEvents_(ss);
    
    if (scanData.events.length === 0) {
      console.log('No event sheets found matching pattern.');
      SpreadsheetApp.getUi().alert(
        'No Events Found',
        'No event sheets matching MM-DD-YYYY or MM-DDX-YYYY format were found.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return { eventsScanned: 0, playersTracked: 0, missionsComputed: 0 };
    }
    
    console.log(`Found ${scanData.events.length} events, ${scanData.players.size} unique players`);
    
    // Step 2: Compute mission progress for all players
    const missionProgress = computeAllMissionProgress_(scanData);
    
    // Step 3: Write to Attendance_Missions
    writeAttendanceMissions_(ss, missionProgress);
    
    // Step 4: Log to Integrity_Log
    const duration = (new Date() - startTime) / 1000;
    logMissionScan_(ss, {
      eventsScanned: scanData.events.length,
      playersTracked: scanData.players.size,
      missionsComputed: Object.keys(MISSION_REGISTRY).length,
      duration: duration
    });
    
    // Show success
    SpreadsheetApp.getUi().alert(
      '✅ Mission Scan Complete',
      `Events Scanned: ${scanData.events.length}\n` +
      `Players Tracked: ${scanData.players.size}\n` +
      `Missions Evaluated: ${Object.keys(MISSION_REGISTRY).length}\n` +
      `Duration: ${duration.toFixed(2)}s`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return {
      eventsScanned: scanData.events.length,
      playersTracked: scanData.players.size,
      missionsComputed: Object.keys(MISSION_REGISTRY).length
    };
    
  } catch (error) {
    console.error('Mission scan failed:', error);
    SpreadsheetApp.getUi().alert(
      '❌ Mission Scan Error',
      error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// ════════════════════════════════════════════════════════════════════════════
// EVENT SCANNING
// ════════════════════════════════════════════════════════════════════════════

/**
 * Scan all event sheets and extract attendance data
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Object} {events, players, playerHistory}
 * @private
 */
function scanAllEvents_(ss) {
  const sheets = ss.getSheets();
  const events = [];
  const players = new Set();
  const playerHistory = new Map(); // playerId -> [{eventId, date, suffix, rank}, ...]
  
  // Load PreferredNames for canonical resolution
  const preferredNames = loadPreferredNamesSet_(ss);
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const parsed = parseEventSheetName_(sheetName);
    
    if (!parsed) return; // Not an event sheet
    
    console.log(`Scanning event: ${sheetName} (suffix: ${parsed.suffix || 'none'})`);
    
    const eventData = extractEventData_(sheet, preferredNames);
    
    if (eventData.players.length === 0) {
      console.log(`  No players found in ${sheetName}`);
      return;
    }
    
    const event = {
      eventId: sheetName,
      date: parsed.date,
      suffix: parsed.suffix,
      players: eventData.players,
      placements: eventData.placements,
      playerCount: eventData.players.length
    };
    
    events.push(event);
    
    // Build player history
    eventData.players.forEach(playerId => {
      players.add(playerId);
      
      if (!playerHistory.has(playerId)) {
        playerHistory.set(playerId, []);
      }
      
      playerHistory.get(playerId).push({
        eventId: sheetName,
        date: parsed.date,
        suffix: parsed.suffix,
        rank: eventData.placements[playerId] || null
      });
    });
  });
  
  // Sort events chronologically
  events.sort((a, b) => a.date - b.date);
  
  // Sort each player's history chronologically
  playerHistory.forEach(history => {
    history.sort((a, b) => a.date - b.date);
  });
  
  return { events, players, playerHistory };
}

/**
 * Parse event sheet name into components
 * @param {string} sheetName - Sheet name
 * @return {Object|null} {date, suffix} or null if not valid event
 * @private
 */
function parseEventSheetName_(sheetName) {
  const match = sheetName.match(MISSION_SCANNER_CONFIG.EVENT_PATTERN);
  if (!match) return null;
  
  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);
  const suffix = match[3] ? match[3].toUpperCase() : null;
  const year = parseInt(match[4], 10);
  
  // Validate date
  const date = new Date(year, month - 1, day);
  if (isNaN(date.getTime()) || 
      date.getMonth() !== month - 1 || 
      date.getDate() !== day) {
    console.warn(`Invalid date in sheet name: ${sheetName}`);
    return null;
  }
  
  return { date, suffix };
}

/**
 * Extract player and placement data from event sheet
 * @param {Sheet} sheet - Event sheet
 * @param {Set<string>} preferredNames - Canonical name set
 * @return {Object} {players, placements}
 * @private
 */
function extractEventData_(sheet, preferredNames) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { players: [], placements: {} };
  
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  // Find player column (flexible matching)
  let playerCol = -1;
  for (const colName of MISSION_SCANNER_CONFIG.PLAYER_COLUMNS) {
    const idx = headers.indexOf(colName);
    if (idx !== -1) {
      playerCol = idx;
      break;
    }
  }
  
  // Find rank column (flexible matching)
  let rankCol = -1;
  for (const colName of MISSION_SCANNER_CONFIG.RANK_COLUMNS) {
    const idx = headers.indexOf(colName);
    if (idx !== -1) {
      rankCol = idx;
      break;
    }
  }
  
  if (playerCol === -1) {
    console.warn(`No player column found in ${sheet.getName()}. Headers: ${headers.join(', ')}`);
    return { players: [], placements: {} };
  }
  
  const players = [];
  const placements = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawName = String(row[playerCol] || '').trim();
    
    if (!rawName) continue;
    
    // Resolve to canonical name
    const playerId = resolveToCanonical_(rawName, preferredNames);
    if (!playerId) continue;
    
    players.push(playerId);
    
    // Extract rank if available
    if (rankCol !== -1) {
      const rank = parseInt(row[rankCol], 10);
      if (!isNaN(rank) && rank > 0) {
        placements[playerId] = rank;
      }
    }
  }
  
  // Determine implicit ranks if not provided (list order)
  if (Object.keys(placements).length === 0 && players.length > 0) {
    players.forEach((playerId, idx) => {
      placements[playerId] = idx + 1;
    });
  }
  
  return { players, placements };
}

/**
 * Load PreferredNames as a Set for fast lookup
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Set<string>} Set of canonical names (lowercase for matching)
 * @private
 */
function loadPreferredNamesSet_(ss) {
  const sheet = ss.getSheetByName(MISSION_SCANNER_CONFIG.SHEETS.PREFERRED_NAMES);
  const names = new Map(); // lowercase -> canonical
  
  if (!sheet) {
    console.warn('PreferredNames sheet not found. Name resolution may be inconsistent.');
    return names;
  }
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) {
      names.set(name.toLowerCase(), name);
    }
  }
  
  console.log(`Loaded ${names.size} preferred names`);
  return names;
}

/**
 * Resolve raw name to canonical PreferredName
 * @param {string} rawName - Input name
 * @param {Map<string,string>} preferredNames - lowercase -> canonical map
 * @return {string|null} Canonical name or null
 * @private
 */
function resolveToCanonical_(rawName, preferredNames) {
  if (!rawName) return null;
  
  const lower = rawName.toLowerCase();
  
  // Exact match (case-insensitive)
  if (preferredNames.has(lower)) {
    return preferredNames.get(lower);
  }
  
  // If PreferredNames is empty, accept as-is
  if (preferredNames.size === 0) {
    return rawName;
  }
  
  // Not found - return as-is but log warning
  console.warn(`Name "${rawName}" not in PreferredNames`);
  return rawName;
}

// ════════════════════════════════════════════════════════════════════════════
// MISSION PROGRESS COMPUTATION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Compute mission progress for all players
 * @param {Object} scanData - {events, players, playerHistory}
 * @return {Object} playerId -> {missionId: progress, ...}
 * @private
 */
function computeAllMissionProgress_(scanData) {
  const progress = {};
  
  // Initialize all players with zero progress
  scanData.players.forEach(playerId => {
    progress[playerId] = {};
    Object.keys(MISSION_REGISTRY).forEach(missionId => {
      progress[playerId][missionId] = 0;
    });
  });
  
  // Compute each mission type
  Object.keys(MISSION_REGISTRY).forEach(missionId => {
    const mission = MISSION_REGISTRY[missionId];
    
    switch (mission.type) {
      case 'one_time':
        computeOneTimeMission_(mission, scanData, progress);
        break;
      case 'format':
        computeFormatMission_(mission, scanData, progress);
        break;
      case 'attendance':
        computeAttendanceMission_(mission, scanData, progress);
        break;
      case 'placement':
        computePlacementMission_(mission, scanData, progress);
        break;
    }
  });
  
  return progress;
}

/**
 * Compute one-time milestone missions
 * @private
 */
function computeOneTimeMission_(mission, scanData, progress) {
  const trigger = mission.trigger;
  
  scanData.players.forEach(playerId => {
    const history = scanData.playerHistory.get(playerId) || [];
    let earned = false;
    
    // Min events threshold
    if (trigger.minEvents && history.length >= trigger.minEvents) {
      earned = true;
    }
    
    // Min formats (different suffixes)
    if (trigger.minFormats) {
      const suffixes = new Set(history.map(e => e.suffix || 'MAIN').filter(s => s));
      if (suffixes.size >= trigger.minFormats) {
        earned = true;
      }
    }
    
    // Events within X days
    if (trigger.eventsInDays) {
      const { count, days } = trigger.eventsInDays;
      earned = checkEventsInWindow_(history, count, days);
    }
    
    progress[playerId][mission.id] = earned ? 1 : 0;
  });
}

/**
 * Compute format-specific one-time missions
 * @private
 */
function computeFormatMission_(mission, scanData, progress) {
  const targetSuffixes = mission.trigger.suffixes || [];
  
  scanData.players.forEach(playerId => {
    const history = scanData.playerHistory.get(playerId) || [];
    const attended = history.some(e => targetSuffixes.includes(e.suffix));
    progress[playerId][mission.id] = attended ? 1 : 0;
  });
}

/**
 * Compute cumulative attendance missions
 * @private
 */
function computeAttendanceMission_(mission, scanData, progress) {
  const trigger = mission.trigger;
  
  scanData.players.forEach(playerId => {
    const history = scanData.playerHistory.get(playerId) || [];
    let count = 0;
    
    if (trigger.all) {
      // Count all events
      count = history.length;
    } else if (trigger.suffixes) {
      // Count events matching specific suffixes
      count = history.filter(e => trigger.suffixes.includes(e.suffix)).length;
    }
    
    // Apply cap if set
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }
    
    progress[playerId][mission.id] = count;
  });
}

/**
 * Compute placement-based missions
 * @private
 */
function computePlacementMission_(mission, scanData, progress) {
  const trigger = mission.trigger;
  
  scanData.players.forEach(playerId => {
    const history = scanData.playerHistory.get(playerId) || [];
    let count = 0;
    
    history.forEach(event => {
      if (trigger.maxRank && event.rank && event.rank <= trigger.maxRank) {
        count++;
      }
      
      if (trigger.isLast) {
        // Check if player was last in this event
        const eventData = scanData.events.find(e => e.eventId === event.eventId);
        if (eventData && event.rank === eventData.playerCount) {
          count++;
        }
      }
    });
    
    // Apply cap if set
    if (mission.cap > 0) {
      count = Math.min(count, mission.cap);
    }
    
    progress[playerId][mission.id] = count;
  });
}

/**
 * Check if player has X events within Y days
 * @private
 */
function checkEventsInWindow_(history, targetCount, windowDays) {
  if (history.length < targetCount) return false;
  
  const windowMs = windowDays * 24 * 60 * 60 * 1000;
  
  // Sliding window check
  for (let i = 0; i <= history.length - targetCount; i++) {
    const windowStart = history[i].date.getTime();
    let eventsInWindow = 1;
    
    for (let j = i + 1; j < history.length; j++) {
      const eventTime = history[j].date.getTime();
      if (eventTime - windowStart <= windowMs) {
        eventsInWindow++;
        if (eventsInWindow >= targetCount) return true;
      } else {
        break;
      }
    }
  }
  
  return false;
}

// ════════════════════════════════════════════════════════════════════════════
// OUTPUT WRITING
// ════════════════════════════════════════════════════════════════════════════

/**
 * Write mission progress to Attendance_Missions sheet
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Object} progress - playerId -> {missionId: value, ...}
 * @private
 */
function writeAttendanceMissions_(ss, progress) {
  let sheet = ss.getSheetByName(MISSION_SCANNER_CONFIG.SHEETS.ATTENDANCE_MISSIONS);
  
  if (!sheet) {
    sheet = ss.insertSheet(MISSION_SCANNER_CONFIG.SHEETS.ATTENDANCE_MISSIONS);
  }
  
  // Clear existing data
  sheet.clear();
  
  // Build headers
  const missionIds = Object.keys(MISSION_REGISTRY);
  const headers = ['PreferredName', ...missionIds.map(id => MISSION_REGISTRY[id].name)];
  
  // Build data rows
  const playerIds = Object.keys(progress).sort();
  const rows = [headers];
  
  playerIds.forEach(playerId => {
    const row = [playerId];
    missionIds.forEach(missionId => {
      row.push(progress[playerId][missionId] || 0);
    });
    rows.push(row);
  });
  
  // Write all data in one batch
  if (rows.length > 0) {
    sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Format header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  
  // Format data
  if (playerIds.length > 0) {
    sheet.getRange(2, 2, playerIds.length, missionIds.length)
      .setNumberFormat('0')
      .setHorizontalAlignment('center');
  }
  
  // Freeze header and name column
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  console.log(`Wrote ${playerIds.length} player rows to Attendance_Missions`);
}

// ════════════════════════════════════════════════════════════════════════════
// MISSIONGATESERVICE COMPATIBILITY
// ════════════════════════════════════════════════════════════════════════════

/**
 * Scan attendance for a date range (required by MissionGateService)
 * @param {Date} startDate - Start of range
 * @param {Date} endDate - End of range
 * @return {Array<Object>} [{playerId, eventId, rank, suffix}, ...]
 */
function scanAttendanceForRange_(startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanData = scanAllEvents_(ss);
  const records = [];
  
  scanData.events.forEach(event => {
    // Filter by date range
    if (event.date < startDate || event.date > endDate) return;
    
    event.players.forEach(playerId => {
      records.push({
        playerId: playerId,
        eventId: event.eventId,
        rank: event.placements[playerId] || null,
        suffix: event.suffix,
        date: event.date
      });
    });
  });
  
  return records;
}

// ════════════════════════════════════════════════════════════════════════════
// LOGGING
// ════════════════════════════════════════════════════════════════════════════

/**
 * Log mission scan to Integrity_Log
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Object} summary - Scan summary
 * @private
 */
function logMissionScan_(ss, summary) {
  let sheet = ss.getSheetByName(MISSION_SCANNER_CONFIG.SHEETS.INTEGRITY_LOG);
  
  if (!sheet) {
    sheet = ss.insertSheet(MISSION_SCANNER_CONFIG.SHEETS.INTEGRITY_LOG);
    sheet.appendRow(['Timestamp', 'User', 'Action', 'Target', 'Details']);
    sheet.setFrozenRows(1);
  }
  
  const timestamp = new Date().toISOString();
  const user = Session.getActiveUser().getEmail() || 'System';
  
  sheet.appendRow([
    timestamp,
    user,
    'MISSION_SCAN',
    'System',
    JSON.stringify(summary)
  ]);
}

// ════════════════════════════════════════════════════════════════════════════
// MENU INTEGRATION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Menu handler for mission scan
 * Call this from your menu: .addItem('Scan + Update Missions', 'onScanMissions')
 */
function onScanMissions() {
  runMissionScan();
}

// ════════════════════════════════════════════════════════════════════════════
// HELPER: isValidSuffix_ (for suffixTests.gs compatibility)
// ════════════════════════════════════════════════════════════════════════════

/**
 * Check if suffix code is valid
 * @param {string} code - Suffix code
 * @return {boolean} True if valid A-Z
 */
function isValidSuffix_(code) {
  if (!code || typeof code !== 'string') return false;
  return /^[A-Z]$/.test(code) && SUFFIX_MAP && SUFFIX_MAP[code];
}

/**
 * Get suffixes filtered by criteria (for suffixTests.gs)
 * @param {Object} filter - {requiresKitPrompt: true, etc.}
 * @return {Array<string>} Matching suffix codes
 */
function getFilteredSuffixes_(filter) {
  if (!SUFFIX_MAP) return [];
  
  return Object.keys(SUFFIX_MAP).filter(code => {
    const meta = SUFFIX_MAP[code];
    if (filter.requiresKitPrompt !== undefined) {
      return meta.requiresKitPrompt === filter.requiresKitPrompt;
    }
    return true;
  });
}

// ════════════════════════════════════════════════════════════════════════════
// END OF MISSION SCANNER SERVICE v1.0.0
// ════════════════════════════════════════════════════════════════════════════