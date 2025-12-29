/**
 * Mission Points Service - Points Pipeline Consolidation
 * @fileoverview Manages the three mission point sources (Flag_Missions, Attendance_Missions, Dice_Points)
 *               and consolidates them into BP_Total
 */

// ============================================================================
// CONSTANTS - Mission Point Categories
// ============================================================================

/**
 * Flag Mission categories (discretionary awards)
 */
const FLAG_MISSION_CATEGORIES = [
  'Cosmic_Selfie',
  'Review_Writer',
  'Social_Media_Star',
  'App_Explorer',
  'Cosmic_Merchant',
  'Precon_Pioneer',
  'Gravitational_Pull',
  'Rogue_Planet',
  'Quantum_Collector'
];

/**
 * Default points per flag mission category
 */
const FLAG_MISSION_DEFAULTS = {
  'Cosmic_Selfie': 5,
  'Review_Writer': 10,
  'Social_Media_Star': 5,
  'App_Explorer': 5,
  'Cosmic_Merchant': 15,
  'Precon_Pioneer': 10,
  'Gravitational_Pull': 10,
  'Rogue_Planet': 20,
  'Quantum_Collector': 25
};

/**
 * Attendance milestone thresholds and point awards
 */
const ATTENDANCE_MILESTONES = {
  5: 5,     // 5 events = 5 bonus points
  10: 10,   // 10 events = 10 bonus points
  25: 15,   // 25 events = 15 bonus points
  50: 25,   // 50 events = 25 bonus points
  100: 50   // 100 events = 50 bonus points
};

// ============================================================================
// PREFERRED NAMES - Canonical Player Registry
// ============================================================================

/**
 * Ensures PreferredNames sheet exists with proper schema
 */
function ensurePreferredNamesSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PreferredNames');

  const requiredHeaders = [
    'preferred_name_id',
    'display_name',
    'created_at',
    'last_active',
    'status'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('PreferredNames');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Gets all canonical player names from PreferredNames
 * @return {Array<Object>} Array of player objects
 */
function getAllPreferredNames() {
  ensurePreferredNamesSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');

  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  return toObjects(data).filter(p => p.preferred_name_id && p.status !== 'INACTIVE');
}

/**
 * Gets canonical player names as a Set for quick lookup
 * @return {Set<string>} Set of preferred_name_id values
 */
function getPreferredNameSet() {
  const players = getAllPreferredNames();
  return new Set(players.map(p => p.preferred_name_id));
}

/**
 * Adds a new player to PreferredNames (and provisions all source sheets)
 * @param {string} preferredNameId - Unique player identifier
 * @param {string} displayName - Display name (optional, defaults to preferredNameId)
 * @return {Object} Result {success, message}
 */
function registerPlayer(preferredNameId, displayName = null) {
  if (!preferredNameId || typeof preferredNameId !== 'string') {
    throwError('Invalid player ID', 'INVALID_INPUT', 'preferred_name_id is required');
  }

  preferredNameId = preferredNameId.trim();
  displayName = displayName ? displayName.trim() : preferredNameId;

  ensurePreferredNamesSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');

  // Check if player already exists
  const existing = getPreferredNameSet();
  if (existing.has(preferredNameId)) {
    return { success: false, message: 'Player already exists' };
  }

  // Add to PreferredNames
  sheet.appendRow([
    preferredNameId,
    displayName,
    dateISO(),
    dateISO(),
    'ACTIVE'
  ]);

  // Provision all source sheets
  provisionPlayerInSourceSheets_(preferredNameId);

  logIntegrityAction('PLAYER_REGISTER', {
    preferredName: preferredNameId,
    details: `Registered new player: ${displayName}`,
    dfTags: ['DF-PROVISION'],
    status: 'SUCCESS'
  });

  return { success: true, message: `Player ${preferredNameId} registered successfully` };
}

/**
 * Provisions a player in all source sheets (Flag_Missions, Attendance_Missions, Dice_Points, BP_Total)
 * @param {string} preferredNameId - Player ID
 * @private
 */
function provisionPlayerInSourceSheets_(preferredNameId) {
  // Provision in Flag_Missions
  ensureFlagMissionsSchema();
  const flagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flag_Missions');
  if (flagSheet && !playerExistsInSheet_(flagSheet, preferredNameId, 'preferred_name_id')) {
    const newRow = [preferredNameId, 0]; // preferred_name_id, Flag_Points (will be computed)
    FLAG_MISSION_CATEGORIES.forEach(() => newRow.push(0));
    newRow.push(dateISO()); // LastUpdated
    flagSheet.appendRow(newRow);
  }

  // Provision in Attendance_Missions
  ensureAttendanceMissionsSchema();
  const attSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance_Missions');
  if (attSheet && !playerExistsInSheet_(attSheet, preferredNameId, 'preferred_name_id')) {
    attSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Events_Attended
      0,                // Current_Streak
      0,                // Best_Streak
      0,                // Attendance_Points
      '',               // Last_Event_Date
      '',               // Milestones_Achieved (comma-separated)
      dateISO()         // LastUpdated
    ]);
  }

  // Provision in Dice_Points
  ensureDicePointsSchema();
  const diceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dice_Points');
  if (diceSheet && !playerExistsInSheet_(diceSheet, preferredNameId, 'preferred_name_id')) {
    diceSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Points
      false,            // Add_Point_Checkbox
      dateISO()         // LastUpdated
    ]);
  }

  // Provision in BP_Total (uses consolidated schema)
  ensureBPTotalConsolidatedSchema();
  const bpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BP_Total');
  if (bpSheet && !playerExistsInSheet_(bpSheet, preferredNameId, 'preferred_name_id')) {
    bpSheet.appendRow([
      preferredNameId,  // preferred_name_id
      0,                // Flag_Mission_Points
      0,                // Attendance_Mission_Points
      0,                // Dice_Points
      0,                // Raw_Total
      0,                // Capped_BP
      0,                // Overflow_to_Prestige
      dateISO()         // LastUpdated
    ]);
  }
}

/**
 * Checks if a player exists in a sheet
 * @param {Sheet} sheet - The sheet to check
 * @param {string} preferredNameId - Player ID
 * @param {string} colName - Column name to check
 * @return {boolean} True if player exists
 * @private
 */
function playerExistsInSheet_(sheet, preferredNameId, colName) {
  if (!sheet || sheet.getLastRow() <= 1) return false;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf(colName);

  if (nameCol === -1) return false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      return true;
    }
  }

  return false;
}

// ============================================================================
// FLAG MISSIONS
// ============================================================================

/**
 * Ensures Flag_Missions sheet exists with proper schema
 */
function ensureFlagMissionsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Flag_Missions');

  const requiredHeaders = [
    'preferred_name_id',
    'Flag_Points'  // Computed sum of all mission columns
  ].concat(FLAG_MISSION_CATEGORIES).concat(['LastUpdated']);

  if (!sheet) {
    sheet = ss.insertSheet('Flag_Missions');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);

    // Set up ARRAYFORMULA for Flag_Points column (B)
    // Formula: sum of columns C through K for each row
    sheet.getRange('B2').setFormula(
      '=ARRAYFORMULA(IF(A2:A="","",SUMIF(ROW(A2:A),"="&ROW(A2:A),C2:C)+SUMIF(ROW(A2:A),"="&ROW(A2:A),D2:D)+SUMIF(ROW(A2:A),"="&ROW(A2:A),E2:E)+SUMIF(ROW(A2:A),"="&ROW(A2:A),F2:F)+SUMIF(ROW(A2:A),"="&ROW(A2:A),G2:G)+SUMIF(ROW(A2:A),"="&ROW(A2:A),H2:H)+SUMIF(ROW(A2:A),"="&ROW(A2:A),I2:I)+SUMIF(ROW(A2:A),"="&ROW(A2:A),J2:J)+SUMIF(ROW(A2:A),"="&ROW(A2:A),K2:K)))'
    );

    // Style header
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setFontWeight('bold')
      .setBackground('#9c27b0')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Awards a flag mission to a player
 * @param {string} preferredNameId - Player ID
 * @param {string} missionCategory - Mission category (e.g., 'Cosmic_Selfie')
 * @param {number} points - Points to award (default: from FLAG_MISSION_DEFAULTS)
 * @return {Object} Result {before, after, awarded}
 */
function awardFlagMission(preferredNameId, missionCategory, points = null) {
  // Validate category
  if (!FLAG_MISSION_CATEGORIES.includes(missionCategory)) {
    throwError('Invalid mission category', 'INVALID_CATEGORY',
      `Must be one of: ${FLAG_MISSION_CATEGORIES.join(', ')}`);
  }

  // Use default points if not specified
  if (points === null) {
    points = FLAG_MISSION_DEFAULTS[missionCategory] || 5;
  }

  ensureFlagMissionsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet) {
    throwError('Flag_Missions sheet not found', 'SHEET_MISSING');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const categoryCol = headers.indexOf(missionCategory);
  const updatedCol = headers.indexOf('LastUpdated');

  if (nameCol === -1 || categoryCol === -1) {
    throwError('Invalid Flag_Missions schema', 'SCHEMA_INVALID');
  }

  // Find player row
  let playerRow = -1;
  let currentPoints = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      playerRow = i;
      currentPoints = coerceNumber(data[i][categoryCol], 0);
      break;
    }
  }

  if (playerRow === -1) {
    // Auto-provision if player doesn't exist
    const canonicalNames = getPreferredNameSet();
    if (!canonicalNames.has(preferredNameId)) {
      throwError('Player not in PreferredNames registry', 'PLAYER_NOT_FOUND',
        `Register player first using registerPlayer("${preferredNameId}")`);
    }
    provisionPlayerInSourceSheets_(preferredNameId);
    // Re-read data
    return awardFlagMission(preferredNameId, missionCategory, points);
  }

  const newPoints = currentPoints + points;
  sheet.getRange(playerRow + 1, categoryCol + 1).setValue(newPoints);
  if (updatedCol !== -1) {
    sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
  }

  logIntegrityAction('FLAG_MISSION_AWARD', {
    preferredName: preferredNameId,
    details: `${missionCategory}: ${currentPoints} → ${newPoints} (+${points})`,
    dfTags: ['DF-MISSION'],
    status: 'SUCCESS'
  });

  return {
    before: currentPoints,
    after: newPoints,
    awarded: points
  };
}

/**
 * Gets a player's flag mission points
 * @param {string} preferredNameId - Player ID
 * @return {Object|null} Player's flag missions data
 */
function getPlayerFlagMissions(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet || sheet.getLastRow() <= 1) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');

  if (nameCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      const result = {};
      headers.forEach((header, idx) => {
        result[header] = data[i][idx];
      });
      return result;
    }
  }

  return null;
}

// ============================================================================
// ATTENDANCE MISSIONS
// ============================================================================

/**
 * Ensures Attendance_Missions sheet exists with proper schema
 */
function ensureAttendanceMissionsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Missions');

  const requiredHeaders = [
    'preferred_name_id',
    'Events_Attended',
    'Current_Streak',
    'Best_Streak',
    'Attendance_Points',
    'Last_Event_Date',
    'Milestones_Achieved',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#ff9800')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Records attendance for a player and awards milestone points
 * @param {string} preferredNameId - Player ID
 * @param {string} eventDate - Event date (ISO format or Date object)
 * @return {Object} Result {eventsAttended, milestonesAwarded, pointsAwarded}
 */
function recordAttendance(preferredNameId, eventDate = null) {
  ensureAttendanceMissionsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet) {
    throwError('Attendance_Missions sheet not found', 'SHEET_MISSING');
  }

  eventDate = eventDate || dateISO();
  if (eventDate instanceof Date) {
    eventDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const eventsCol = headers.indexOf('Events_Attended');
  const streakCol = headers.indexOf('Current_Streak');
  const bestStreakCol = headers.indexOf('Best_Streak');
  const pointsCol = headers.indexOf('Attendance_Points');
  const lastEventCol = headers.indexOf('Last_Event_Date');
  const milestonesCol = headers.indexOf('Milestones_Achieved');
  const updatedCol = headers.indexOf('LastUpdated');

  // Find player row
  let playerRow = -1;
  let currentData = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      playerRow = i;
      currentData = {
        events: coerceNumber(data[i][eventsCol], 0),
        streak: coerceNumber(data[i][streakCol], 0),
        bestStreak: coerceNumber(data[i][bestStreakCol], 0),
        points: coerceNumber(data[i][pointsCol], 0),
        lastEvent: data[i][lastEventCol],
        milestones: data[i][milestonesCol] ? String(data[i][milestonesCol]).split(',').map(m => parseInt(m, 10)).filter(m => !isNaN(m)) : []
      };
      break;
    }
  }

  if (playerRow === -1) {
    // Auto-provision
    const canonicalNames = getPreferredNameSet();
    if (!canonicalNames.has(preferredNameId)) {
      throwError('Player not in PreferredNames registry', 'PLAYER_NOT_FOUND');
    }
    provisionPlayerInSourceSheets_(preferredNameId);
    return recordAttendance(preferredNameId, eventDate);
  }

  // Increment events attended
  const newEvents = currentData.events + 1;

  // Update streak (simplified: just increment for now)
  const newStreak = currentData.streak + 1;
  const newBestStreak = Math.max(currentData.bestStreak, newStreak);

  // Check for new milestones
  const milestonesAwarded = [];
  let pointsAwarded = 0;

  Object.keys(ATTENDANCE_MILESTONES).forEach(threshold => {
    const thresholdNum = parseInt(threshold, 10);
    if (newEvents >= thresholdNum && !currentData.milestones.includes(thresholdNum)) {
      milestonesAwarded.push(thresholdNum);
      pointsAwarded += ATTENDANCE_MILESTONES[threshold];
    }
  });

  const newPoints = currentData.points + pointsAwarded;
  const newMilestones = currentData.milestones.concat(milestonesAwarded).sort((a, b) => a - b);

  // Update sheet
  sheet.getRange(playerRow + 1, eventsCol + 1).setValue(newEvents);
  sheet.getRange(playerRow + 1, streakCol + 1).setValue(newStreak);
  sheet.getRange(playerRow + 1, bestStreakCol + 1).setValue(newBestStreak);
  sheet.getRange(playerRow + 1, pointsCol + 1).setValue(newPoints);
  sheet.getRange(playerRow + 1, lastEventCol + 1).setValue(eventDate);
  sheet.getRange(playerRow + 1, milestonesCol + 1).setValue(newMilestones.join(','));
  sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());

  if (milestonesAwarded.length > 0) {
    logIntegrityAction('ATTENDANCE_MILESTONE', {
      preferredName: preferredNameId,
      details: `Events: ${newEvents}. Milestones: ${milestonesAwarded.join(', ')}. Points: +${pointsAwarded}`,
      dfTags: ['DF-MILESTONE'],
      status: 'SUCCESS'
    });
  }

  return {
    eventsAttended: newEvents,
    milestonesAwarded,
    pointsAwarded,
    totalPoints: newPoints
  };
}

/**
 * Gets a player's attendance data
 * @param {string} preferredNameId - Player ID
 * @return {Object|null} Player's attendance data
 */
function getPlayerAttendance(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet || sheet.getLastRow() <= 1) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');

  if (nameCol === -1) return null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      const result = {};
      headers.forEach((header, idx) => {
        result[header] = data[i][idx];
      });
      return result;
    }
  }

  return null;
}

// ============================================================================
// DICE POINTS
// ============================================================================

/**
 * Ensures Dice_Points sheet exists with proper schema
 */
function ensureDicePointsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Dice_Points');

  const requiredHeaders = [
    'preferred_name_id',
    'Points',
    'Add_Point_Checkbox',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Dice_Points');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#4caf50')
      .setFontColor('#ffffff');

    // Set checkbox data validation for column C
    // (Applied per-row when players are added)
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Awards dice points to a player
 * @param {string} preferredNameId - Player ID
 * @param {number} points - Points to add (default: 1)
 * @return {Object} Result {before, after, awarded}
 */
function awardDicePoints(preferredNameId, points = 1) {
  ensureDicePointsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dice_Points');

  if (!sheet) {
    throwError('Dice_Points sheet not found', 'SHEET_MISSING');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const pointsCol = headers.indexOf('Points');
  const updatedCol = headers.indexOf('LastUpdated');

  // Find player row
  let playerRow = -1;
  let currentPoints = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      playerRow = i;
      currentPoints = coerceNumber(data[i][pointsCol], 0);
      break;
    }
  }

  if (playerRow === -1) {
    // Auto-provision
    const canonicalNames = getPreferredNameSet();
    if (!canonicalNames.has(preferredNameId)) {
      throwError('Player not in PreferredNames registry', 'PLAYER_NOT_FOUND');
    }
    provisionPlayerInSourceSheets_(preferredNameId);
    return awardDicePoints(preferredNameId, points);
  }

  const newPoints = currentPoints + points;
  sheet.getRange(playerRow + 1, pointsCol + 1).setValue(newPoints);
  if (updatedCol !== -1) {
    sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
  }

  logIntegrityAction('DICE_POINTS_AWARD', {
    preferredName: preferredNameId,
    details: `Dice Points: ${currentPoints} → ${newPoints} (+${points})`,
    dfTags: ['DF-DICE'],
    status: 'SUCCESS'
  });

  return {
    before: currentPoints,
    after: newPoints,
    awarded: points
  };
}

/**
 * Gets a player's dice points
 * @param {string} preferredNameId - Player ID
 * @return {number} Current dice points (0 if not found)
 */
function getPlayerDicePoints(preferredNameId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dice_Points');

  if (!sheet || sheet.getLastRow() <= 1) return 0;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('preferred_name_id');
  const pointsCol = headers.indexOf('Points');

  if (nameCol === -1 || pointsCol === -1) return 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredNameId) {
      return coerceNumber(data[i][pointsCol], 0);
    }
  }

  return 0;
}

/**
 * Handles checkbox click for dice point increment (for onEdit trigger)
 * @param {Object} e - Edit event object
 */
function onDicePointCheckboxEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Dice_Points') return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const checkboxCol = headers.indexOf('Add_Point_Checkbox') + 1;

  if (e.range.getColumn() !== checkboxCol) return;
  if (e.value !== 'TRUE' && e.value !== true) return;

  const row = e.range.getRow();
  if (row <= 1) return; // Header row

  const nameCol = headers.indexOf('preferred_name_id') + 1;
  const pointsCol = headers.indexOf('Points') + 1;
  const updatedCol = headers.indexOf('LastUpdated') + 1;

  const preferredNameId = sheet.getRange(row, nameCol).getValue();
  const currentPoints = coerceNumber(sheet.getRange(row, pointsCol).getValue(), 0);

  // Increment points
  sheet.getRange(row, pointsCol).setValue(currentPoints + 1);
  if (updatedCol > 0) {
    sheet.getRange(row, updatedCol).setValue(dateISO());
  }

  // Reset checkbox
  e.range.setValue(false);

  logIntegrityAction('DICE_POINTS_CHECKBOX', {
    preferredName: preferredNameId,
    details: `Quick-add: ${currentPoints} → ${currentPoints + 1}`,
    status: 'SUCCESS'
  });
}

// ============================================================================
// BP_TOTAL CONSOLIDATION
// ============================================================================

/**
 * Ensures BP_Total has the consolidated schema with all source columns
 */
function ensureBPTotalConsolidatedSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  const requiredHeaders = [
    'preferred_name_id',
    'Flag_Mission_Points',
    'Attendance_Mission_Points',
    'Dice_Points',
    'Raw_Total',
    'Capped_BP',
    'Overflow_to_Prestige',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#2196f3')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check for old schema (PreferredName, BP_Current, LastUpdated)
  // and migrate if needed
  if (headers.includes('PreferredName') && headers.includes('BP_Current')) {
    migrateBPTotalSchema_(sheet);
  }

  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Migrates old BP_Total schema to new consolidated schema
 * @param {Sheet} sheet - BP_Total sheet
 * @private
 */
function migrateBPTotalSchema_(sheet) {
  const data = sheet.getDataRange().getValues();
  const oldHeaders = data[0];

  const nameCol = oldHeaders.indexOf('PreferredName');
  const bpCol = oldHeaders.indexOf('BP_Current');
  const updatedCol = oldHeaders.indexOf('LastUpdated');

  if (nameCol === -1 || bpCol === -1) return;

  // Build new data
  const newHeaders = [
    'preferred_name_id',
    'Flag_Mission_Points',
    'Attendance_Mission_Points',
    'Dice_Points',
    'Raw_Total',
    'Capped_BP',
    'Overflow_to_Prestige',
    'LastUpdated'
  ];

  const newData = [newHeaders];

  for (let i = 1; i < data.length; i++) {
    const oldName = data[i][nameCol];
    const oldBP = coerceNumber(data[i][bpCol], 0);
    const oldUpdated = updatedCol !== -1 ? data[i][updatedCol] : dateISO();

    if (oldName) {
      const throttle = getThrottleKV();
      const globalCap = coerceNumber(throttle.BP_Global_Cap, 100);
      const capped = Math.min(oldBP, globalCap);
      const overflow = Math.max(0, oldBP - globalCap);

      newData.push([
        oldName,      // preferred_name_id
        0,            // Flag_Mission_Points (will be synced)
        0,            // Attendance_Mission_Points (will be synced)
        0,            // Dice_Points (will be synced)
        oldBP,        // Raw_Total (preserve old BP as starting point)
        capped,       // Capped_BP
        overflow,     // Overflow_to_Prestige
        oldUpdated    // LastUpdated
      ]);
    }
  }

  // Clear and rewrite
  sheet.clear();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  sheet.setFrozenRows(1);
  sheet.getRange('A1:H1')
    .setFontWeight('bold')
    .setBackground('#2196f3')
    .setFontColor('#ffffff');

  logIntegrityAction('BP_TOTAL_MIGRATE', {
    details: `Migrated ${newData.length - 1} players to consolidated schema`,
    dfTags: ['DF-MIGRATE'],
    status: 'SUCCESS'
  });
}

/**
 * Syncs BP_Total from all three source sheets
 * @param {boolean} dryRun - If true, returns preview without writing
 * @return {Object} Sync result {synced, errors, preview}
 */
function syncBPTotals(dryRun = false) {
  ensureBPTotalConsolidatedSchema();
  ensureFlagMissionsSchema();
  ensureAttendanceMissionsSchema();
  ensureDicePointsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpSheet = ss.getSheetByName('BP_Total');
  const flagSheet = ss.getSheetByName('Flag_Missions');
  const attSheet = ss.getSheetByName('Attendance_Missions');
  const diceSheet = ss.getSheetByName('Dice_Points');

  // Get canonical names
  const canonicalNames = getPreferredNameSet();

  // Build lookup maps from source sheets
  const flagMap = buildSourceMap_(flagSheet, 'preferred_name_id', 'Flag_Points');
  const attMap = buildSourceMap_(attSheet, 'preferred_name_id', 'Attendance_Points');
  const diceMap = buildSourceMap_(diceSheet, 'preferred_name_id', 'Points');

  // Get throttle config
  const throttle = getThrottleKV();
  const globalCap = coerceNumber(throttle.BP_Global_Cap, 100);

  // Process BP_Total
  const bpData = bpSheet.getDataRange().getValues();
  const bpHeaders = bpData[0];

  const nameCol = bpHeaders.indexOf('preferred_name_id');
  const flagCol = bpHeaders.indexOf('Flag_Mission_Points');
  const attCol = bpHeaders.indexOf('Attendance_Mission_Points');
  const diceCol = bpHeaders.indexOf('Dice_Points');
  const rawCol = bpHeaders.indexOf('Raw_Total');
  const cappedCol = bpHeaders.indexOf('Capped_BP');
  const overflowCol = bpHeaders.indexOf('Overflow_to_Prestige');
  const updatedCol = bpHeaders.indexOf('LastUpdated');

  const errors = [];
  const preview = [];
  let syncedCount = 0;

  // Track which players exist in BP_Total
  const existingPlayers = new Set();

  for (let i = 1; i < bpData.length; i++) {
    const playerId = bpData[i][nameCol];
    if (!playerId) continue;

    existingPlayers.add(playerId);

    // Validate against canonical names
    if (!canonicalNames.has(playerId)) {
      errors.push({ player: playerId, error: 'Not in PreferredNames registry' });
      continue;
    }

    const flagPoints = coerceNumber(flagMap.get(playerId), 0);
    const attPoints = coerceNumber(attMap.get(playerId), 0);
    const dicePoints = coerceNumber(diceMap.get(playerId), 0);
    const rawTotal = flagPoints + attPoints + dicePoints;
    const cappedBP = Math.min(rawTotal, globalCap);
    const overflow = Math.max(0, rawTotal - globalCap);

    const change = {
      player: playerId,
      flag: flagPoints,
      attendance: attPoints,
      dice: dicePoints,
      raw: rawTotal,
      capped: cappedBP,
      overflow: overflow
    };
    preview.push(change);

    if (!dryRun) {
      bpSheet.getRange(i + 1, flagCol + 1).setValue(flagPoints);
      bpSheet.getRange(i + 1, attCol + 1).setValue(attPoints);
      bpSheet.getRange(i + 1, diceCol + 1).setValue(dicePoints);
      bpSheet.getRange(i + 1, rawCol + 1).setValue(rawTotal);
      bpSheet.getRange(i + 1, cappedCol + 1).setValue(cappedBP);
      bpSheet.getRange(i + 1, overflowCol + 1).setValue(overflow);
      bpSheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());

      // Route overflow to Prestige_Overflow
      if (overflow > 0) {
        routeOverflowToPrestige_(playerId, overflow);
      }
    }

    syncedCount++;
  }

  // Check for players in source sheets not in BP_Total
  const allSourcePlayers = new Set([...flagMap.keys(), ...attMap.keys(), ...diceMap.keys()]);
  allSourcePlayers.forEach(playerId => {
    if (!existingPlayers.has(playerId)) {
      if (canonicalNames.has(playerId)) {
        errors.push({ player: playerId, error: 'In source sheet but not in BP_Total (run provision)' });
      } else {
        errors.push({ player: playerId, error: 'In source sheet but not in PreferredNames registry' });
      }
    }
  });

  // Validate no negative values
  preview.forEach(p => {
    if (p.flag < 0 || p.attendance < 0 || p.dice < 0) {
      errors.push({ player: p.player, error: 'Negative point values detected' });
    }
  });

  if (!dryRun) {
    const checksum = computeChecksum(bpSheet.getDataRange().getValues());
    logIntegrityAction('BP_SYNC', {
      checksumAfter: checksum,
      details: `Synced ${syncedCount} players. Errors: ${errors.length}`,
      dfTags: ['DF-SYNC'],
      status: errors.length === 0 ? 'SUCCESS' : 'WARNING'
    });
  }

  return {
    synced: syncedCount,
    errors,
    preview
  };
}

/**
 * Builds a lookup map from a source sheet
 * @param {Sheet} sheet - Source sheet
 * @param {string} keyCol - Key column name
 * @param {string} valueCol - Value column name
 * @return {Map} Map of key -> value
 * @private
 */
function buildSourceMap_(sheet, keyCol, valueCol) {
  const map = new Map();

  if (!sheet || sheet.getLastRow() <= 1) return map;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keyIdx = headers.indexOf(keyCol);
  const valueIdx = headers.indexOf(valueCol);

  if (keyIdx === -1 || valueIdx === -1) return map;

  for (let i = 1; i < data.length; i++) {
    const key = data[i][keyIdx];
    const value = data[i][valueIdx];
    if (key) {
      map.set(key, coerceNumber(value, 0));
    }
  }

  return map;
}

/**
 * Routes overflow points to Prestige_Overflow sheet
 * @param {string} preferredNameId - Player ID
 * @param {number} overflow - Overflow amount
 * @private
 */
function routeOverflowToPrestige_(preferredNameId, overflow) {
  // This uses the existing addPrestigeOverflow_ from bpService.gs
  // But we need to track it separately for the sync
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prestige_Overflow');

  if (!sheet) {
    sheet = ss.insertSheet('Prestige_Overflow');
    sheet.appendRow(['PreferredName', 'Total_Overflow', 'Last_Updated', 'Prestige_Tier']);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  let playerRow = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === preferredNameId) {
      playerRow = i;
      break;
    }
  }

  if (playerRow === -1) {
    sheet.appendRow([preferredNameId, overflow, dateISO(), computePrestigeTier_(overflow)]);
  } else {
    // Update existing - set to current overflow (not cumulative from sync)
    sheet.getRange(playerRow + 1, 2).setValue(overflow);
    sheet.getRange(playerRow + 1, 3).setValue(dateISO());
    sheet.getRange(playerRow + 1, 4).setValue(computePrestigeTier_(overflow));
  }
}

/**
 * Computes prestige tier based on overflow amount
 * @param {number} overflow - Total overflow
 * @return {string} Tier name
 * @private
 */
function computePrestigeTier_(overflow) {
  if (overflow >= 500) return 'Diamond';
  if (overflow >= 250) return 'Platinum';
  if (overflow >= 100) return 'Gold';
  if (overflow >= 50) return 'Silver';
  return 'Bronze';
}

// ============================================================================
// VALIDATION & INTEGRITY
// ============================================================================

/**
 * Validates all mission point data for integrity
 * @return {Object} Validation result {pass, issues}
 */
function validateMissionPointsIntegrity() {
  const issues = [];

  // Get canonical names
  const canonicalNames = getPreferredNameSet();

  // Check Flag_Missions
  const flagSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flag_Missions');
  if (flagSheet && flagSheet.getLastRow() > 1) {
    const data = flagSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Flag_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      // Check for negative values
      FLAG_MISSION_CATEGORIES.forEach(cat => {
        const catCol = headers.indexOf(cat);
        if (catCol !== -1 && coerceNumber(data[i][catCol], 0) < 0) {
          issues.push({
            sheet: 'Flag_Missions',
            row: i + 1,
            player: playerId,
            issue: `Negative value in ${cat}`
          });
        }
      });
    }
  }

  // Check Attendance_Missions
  const attSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance_Missions');
  if (attSheet && attSheet.getLastRow() > 1) {
    const data = attSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');
    const pointsCol = headers.indexOf('Attendance_Points');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Attendance_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      if (pointsCol !== -1 && coerceNumber(data[i][pointsCol], 0) < 0) {
        issues.push({
          sheet: 'Attendance_Missions',
          row: i + 1,
          player: playerId,
          issue: 'Negative attendance points'
        });
      }
    }
  }

  // Check Dice_Points
  const diceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dice_Points');
  if (diceSheet && diceSheet.getLastRow() > 1) {
    const data = diceSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');
    const pointsCol = headers.indexOf('Points');

    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][nameCol];
      if (playerId && !canonicalNames.has(playerId)) {
        issues.push({
          sheet: 'Dice_Points',
          row: i + 1,
          player: playerId,
          issue: 'Not in PreferredNames registry'
        });
      }

      if (pointsCol !== -1 && coerceNumber(data[i][pointsCol], 0) < 0) {
        issues.push({
          sheet: 'Dice_Points',
          row: i + 1,
          player: playerId,
          issue: 'Negative dice points'
        });
      }
    }
  }

  return {
    pass: issues.length === 0,
    issues
  };
}

/**
 * Provisions all existing PreferredNames players in all source sheets
 * @return {Object} Result {provisioned, alreadyExisted}
 */
function provisionAllPlayers() {
  const players = getAllPreferredNames();
  let provisioned = 0;
  let alreadyExisted = 0;

  players.forEach(player => {
    const playerId = player.preferred_name_id;

    // Check if already provisioned in all sheets
    const flagExists = playerExistsInSheet_(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flag_Missions'),
      playerId, 'preferred_name_id'
    );
    const attExists = playerExistsInSheet_(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Attendance_Missions'),
      playerId, 'preferred_name_id'
    );
    const diceExists = playerExistsInSheet_(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dice_Points'),
      playerId, 'preferred_name_id'
    );
    const bpExists = playerExistsInSheet_(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('BP_Total'),
      playerId, 'preferred_name_id'
    );

    if (flagExists && attExists && diceExists && bpExists) {
      alreadyExisted++;
    } else {
      provisionPlayerInSourceSheets_(playerId);
      provisioned++;
    }
  });

  logIntegrityAction('PROVISION_ALL', {
    details: `Provisioned: ${provisioned}, Already existed: ${alreadyExisted}`,
    dfTags: ['DF-PROVISION'],
    status: 'SUCCESS'
  });

  return { provisioned, alreadyExisted };
}

// ============================================================================
// COMPATIBILITY LAYER - Bridges to existing getCanonicalNames()
// ============================================================================

/**
 * Enhanced getCanonicalNames that includes PreferredNames sheet
 * This overrides/extends the existing implementation
 * @return {Array<string>} Sorted array of canonical names
 */
function getCanonicalNames() {
  const names = new Set();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Source 1: PreferredNames (canonical source)
  const prefSheet = ss.getSheetByName('PreferredNames');
  if (prefSheet && prefSheet.getLastRow() > 1) {
    const data = prefSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');
    const statusCol = headers.indexOf('status');

    if (nameCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        const name = data[i][nameCol];
        const status = statusCol !== -1 ? data[i][statusCol] : 'ACTIVE';
        if (name && status !== 'INACTIVE') {
          names.add(name);
        }
      }
    }
  }

  // Source 2: Key_Tracker (legacy compatibility)
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet && keySheet.getLastRow() > 1) {
    const data = keySheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) names.add(data[i][0]); // PreferredName in column A
    }
  }

  // Source 3: BP_Total (legacy compatibility)
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const data = bpSheet.getDataRange().getValues();
    const headers = data[0];
    // Check both old and new column names
    const nameCol = headers.indexOf('preferred_name_id') !== -1
      ? headers.indexOf('preferred_name_id')
      : headers.indexOf('PreferredName');

    if (nameCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][nameCol]) names.add(data[i][nameCol]);
      }
    }
  }

  return Array.from(names).sort();
}

// ============================================================================
// PLAYER LOOKUP - Unified Player Data
// ============================================================================

/**
 * Gets comprehensive player data from all sources
 * @param {string} preferredNameId - Player ID
 * @return {Object} Complete player profile
 */
function getPlayerProfile(preferredNameId) {
  const profile = {
    id: preferredNameId,
    registered: false,
    flag_missions: null,
    attendance: null,
    dice_points: 0,
    bp_total: null,
    keys: null
  };

  // Check PreferredNames
  const canonicalNames = getPreferredNameSet();
  profile.registered = canonicalNames.has(preferredNameId);

  // Get Flag Missions
  profile.flag_missions = getPlayerFlagMissions(preferredNameId);

  // Get Attendance
  profile.attendance = getPlayerAttendance(preferredNameId);

  // Get Dice Points
  profile.dice_points = getPlayerDicePoints(preferredNameId);

  // Get BP Total
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const data = bpSheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('preferred_name_id');

    if (nameCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][nameCol] === preferredNameId) {
          profile.bp_total = {};
          headers.forEach((h, idx) => {
            profile.bp_total[h] = data[i][idx];
          });
          break;
        }
      }
    }
  }

  // Get Keys
  profile.keys = getPlayerKeys(preferredNameId);

  return profile;
}