/**
 * Player Provisioning Service - Cosmic Tournament Manager v7.9.7
 * @fileoverview Idempotent, header-driven player provisioning pipeline
 *
 * Ensures each player in PreferredNames has a "profile" row in all
 * required data sheets. This is the canonical provisioning engine.
 */

// ============================================================================
// CONSTANTS & CONFIGURATION
// ============================================================================

/**
 * Canonical sheet name for the source of truth on player names
 * @const {string}
 */
const PREFERRED_NAMES_SHEET = 'PreferredNames';

/**
 * Header name synonyms for player key columns (case-insensitive matching)
 * @const {string[]}
 */
const PLAYER_KEY_HEADER_NAMES = [
  'PreferredName',
  'Preferred Name',
  'preferred_name_id',
  'Player',
  'Player Name'
];

/**
 * Configuration for all sheets where players need profile rows.
 * - sheetName: Exact name of the target sheet
 * - keyHeaders: Array of header synonyms to find the key column
 * - optional: If true, missing sheet is tolerated; if false/undefined, missing sheet throws
 *
 * NOTE: Store_Credit_Ledger is intentionally excluded - it's transactional, not per-player.
 * @const {Array<Object>}
 */
const PLAYER_PROVISION_TARGETS = [
  // Core mission/points sheets (required)
  { sheetName: 'Attendance_Missions', keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: false },
  { sheetName: 'Flag_Missions',       keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: false },
  { sheetName: 'BP_Total',            keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: false },

  // Dice points (may be named either way)
  { sheetName: 'Dice Roll Points',    keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },
  { sheetName: 'Dice_Points',         keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },

  // Key tracker (may be named either way)
  { sheetName: 'Key Tracker',         keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },
  { sheetName: 'Key_Tracker',         keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },

  // Prize wall points
  { sheetName: "Player's Prize-Wall-Points", keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },

  // Attendance calendar
  { sheetName: 'Attendance_Calendar', keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },

  // Achievements / awards (optional backend sheets)
  { sheetName: 'Player Achievements', keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true },
  { sheetName: 'players_awards_lists', keyHeaders: PLAYER_KEY_HEADER_NAMES, optional: true }
];

// ============================================================================
// LOW-LEVEL HELPERS
// ============================================================================

/**
 * Finds the first column index whose header matches one of the provided names (case-insensitive).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to search
 * @param {string[]} headerNames - Array of candidate header names to match
 * @return {number} 1-based column index, or -1 if not found
 * @private
 */
function findColumnIndexByHeader_(sheet, headerNames) {
  if (!sheet) return -1;

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return -1;

  // Read header row (row 1)
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Normalize candidate names to lowercase for comparison
  const normalizedCandidates = headerNames.map(h => String(h).toLowerCase().trim());

  // Find first matching header
  for (let i = 0; i < headers.length; i++) {
    const headerNormalized = String(headers[i]).toLowerCase().trim();

    for (const candidate of normalizedCandidates) {
      if (headerNormalized === candidate) {
        return i + 1; // Convert to 1-based index
      }
    }
  }

  return -1;
}

/**
 * Gets all values from a specific column (starting at row 2).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet
 * @param {number} colIndex - 1-based column index
 * @return {string[]} Array of trimmed string values
 * @private
 */
function getColumnValues_(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  return values.map(row => String(row[0]).trim());
}

// ============================================================================
// CORE PROVISIONING FUNCTIONS
// ============================================================================

/**
 * Ensures a given player has a row in the specified sheet.
 *
 * This is the core idempotent operation: if the player already exists,
 * nothing changes; if they don't exist, a new row is appended with the
 * player name in the key column and blank values elsewhere.
 *
 * @param {string} sheetName - Name of the target sheet
 * @param {string} preferredName - The player's canonical name
 * @param {string[]} headerNames - Candidate key header names
 * @param {Object} [options] - Options
 * @param {boolean} [options.optional=false] - If true, missing sheet is tolerated
 * @return {Object} result
 *   - sheetName: string - The sheet name processed
 *   - created: boolean - true if a new row was created
 *   - existed: boolean - true if player already had a row
 *   - skipped: boolean - true if sheet was missing but optional
 */
function ensurePlayerRowInSheet_(sheetName, preferredName, headerNames, options) {
  options = options || {};
  const optional = !!options.optional;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  // Handle missing sheet
  if (!sheet) {
    if (optional) {
      return {
        sheetName: sheetName,
        created: false,
        existed: false,
        skipped: true
      };
    }
    throw new Error('Required sheet "' + sheetName + '" not found.');
  }

  // Find the key column by header
  const keyColIndex = findColumnIndexByHeader_(sheet, headerNames);

  // Handle missing key column
  if (keyColIndex === -1) {
    if (optional) {
      return {
        sheetName: sheetName,
        created: false,
        existed: false,
        skipped: true
      };
    }
    throw new Error('Required key column not found in sheet "' + sheetName + '". Expected one of: ' + headerNames.join(', '));
  }

  // Check if player already exists (case-sensitive match)
  const existingValues = getColumnValues_(sheet, keyColIndex);
  const playerExists = existingValues.some(val => val === preferredName);

  if (playerExists) {
    return {
      sheetName: sheetName,
      created: false,
      existed: true,
      skipped: false
    };
  }

  // Player not found - create new row
  const lastCol = sheet.getLastColumn();
  const rowValues = new Array(Math.max(lastCol, keyColIndex)).fill('');
  rowValues[keyColIndex - 1] = preferredName;

  sheet.appendRow(rowValues);

  return {
    sheetName: sheetName,
    created: true,
    existed: false,
    skipped: false
  };
}

/**
 * Ensures the given player has a row in all provision targets.
 *
 * This is the main "provision this one person everywhere" call.
 * It iterates through PLAYER_PROVISION_TARGETS and ensures the
 * player has a row in each sheet.
 *
 * @param {string} preferredName - The player's canonical name
 * @return {Object} result
 *   - playerName: string - The processed player name
 *   - createdBySheet: Object.<string, boolean> - true if row created per sheet
 *   - existedBySheet: Object.<string, boolean> - true if row existed per sheet
 *   - skippedSheets: string[] - names of sheets that were skipped (missing but optional)
 */
function provisionSinglePlayerProfile_(preferredName) {
  // Validate input
  if (!preferredName || !String(preferredName).trim()) {
    throw new Error('preferredName is required.');
  }

  const name = String(preferredName).trim();

  const createdBySheet = {};
  const existedBySheet = {};
  const skippedSheets = [];

  // Process each target sheet
  PLAYER_PROVISION_TARGETS.forEach(function(target) {
    const result = ensurePlayerRowInSheet_(
      target.sheetName,
      name,
      target.keyHeaders || PLAYER_KEY_HEADER_NAMES,
      { optional: !!target.optional }
    );

    createdBySheet[target.sheetName] = !!result.created;
    existedBySheet[target.sheetName] = !!result.existed;

    if (result.skipped) {
      skippedSheets.push(target.sheetName);
    }
  });

  return {
    playerName: name,
    createdBySheet: createdBySheet,
    existedBySheet: existedBySheet,
    skippedSheets: skippedSheets
  };
}

/**
 * Ensures that all players in PreferredNames exist in all provision targets.
 *
 * This is the bulk operation for provisioning all players at once.
 * It reads the PreferredNames sheet and calls provisionSinglePlayerProfile_
 * for each player.
 *
 * @return {Object} result
 *   - totalPlayers: number - Total players processed
 *   - provisionedNewRows: number - Total new rows created across all sheets
 *   - playersWithAnyNewRows: number - Count of players for whom at least one row was created
 *   - skippedSheets: string[] - Union of all skipped sheet names
 */
function provisionAllPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefSheet = ss.getSheetByName(PREFERRED_NAMES_SHEET);

  if (!prefSheet) {
    throw new Error('PreferredNames sheet not found.');
  }

  const lastRow = prefSheet.getLastRow();

  // Handle empty or header-only sheet
  if (lastRow < 2) {
    return {
      totalPlayers: 0,
      provisionedNewRows: 0,
      playersWithAnyNewRows: 0,
      skippedSheets: []
    };
  }

  // Read all player names (column A, rows 2 to lastRow)
  const names = prefSheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(function(row) { return String(row[0]).trim(); })
    .filter(function(name) { return name.length > 0; });

  let provisionedNewRows = 0;
  let playersWithAnyNewRows = 0;
  const allSkippedSheets = new Set();

  // Process each player
  names.forEach(function(name) {
    const result = provisionSinglePlayerProfile_(name);

    // Count new rows created for this player
    let playerNewRows = 0;
    Object.keys(result.createdBySheet).forEach(function(sheetName) {
      if (result.createdBySheet[sheetName]) {
        playerNewRows++;
        provisionedNewRows++;
      }
    });

    if (playerNewRows > 0) {
      playersWithAnyNewRows++;
    }

    // Track skipped sheets
    result.skippedSheets.forEach(function(sheetName) {
      allSkippedSheets.add(sheetName);
    });
  });

  return {
    totalPlayers: names.length,
    provisionedNewRows: provisionedNewRows,
    playersWithAnyNewRows: playersWithAnyNewRows,
    skippedSheets: Array.from(allSkippedSheets)
  };
}

// ============================================================================
// PUBLIC API: ADD NEW PLAYER
// ============================================================================

/**
 * Adds a new player to the system and provisions their profile across all sheets.
 *
 * This is the main entry point for adding players. It:
 * 1. Checks if the player already exists in PreferredNames (case-insensitive)
 * 2. If new, adds them to PreferredNames
 * 3. Provisions their profile row in all target sheets
 *
 * @param {string} name - The player's name
 * @return {Object} result
 *   - success: boolean - Always true if no error thrown
 *   - alreadyExisted: boolean - true if player was already in PreferredNames
 *   - playerName: string - The canonical player name used
 *   - profileResult: Object - Result from provisionSinglePlayerProfile_
 *   - message: string - Human-readable result message
 */
function addNewPlayer(name) {
  if (!name) {
    throw new Error('Player name is required.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PREFERRED_NAMES_SHEET);

  if (!sheet) {
    throw new Error('PreferredNames sheet not found.');
  }

  const trimmed = String(name).trim();
  if (!trimmed) {
    throw new Error('Player name is required.');
  }

  // Check for existing player (case-insensitive)
  const lastRow = sheet.getLastRow();
  let existingCanonical = null;

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      const existing = String(values[i][0] || '').trim();
      if (existing && existing.toLowerCase() === trimmed.toLowerCase()) {
        existingCanonical = existing;
        break;
      }
    }
  }

  let playerNameToUse = trimmed;
  let alreadyExisted = false;

  if (existingCanonical) {
    // Use existing canonical name (preserves original casing)
    playerNameToUse = existingCanonical;
    alreadyExisted = true;
  } else {
    // Add new player to PreferredNames
    sheet.appendRow([trimmed]);
  }

  // Provision this player across all target sheets
  const profileResult = provisionSinglePlayerProfile_(playerNameToUse);

  return {
    success: true,
    alreadyExisted: alreadyExisted,
    playerName: playerNameToUse,
    profileResult: profileResult,
    message: alreadyExisted
      ? 'Player already existed; profile provisioning refreshed.'
      : 'Player added and profile provisioning completed.'
  };
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets all preferred names from the PreferredNames sheet.
 * This is the canonical source of truth for player names.
 *
 * @return {string[]} Array of preferred names
 */
function getAllPreferredNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PREFERRED_NAMES_SHEET);

  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, 1).getValues()
    .map(function(row) { return String(row[0]).trim(); })
    .filter(function(name) { return name.length > 0; })
    .sort();
}

/**
 * Ensures the PreferredNames sheet exists with proper schema.
 * Creates it if missing.
 */
function ensurePreferredNamesSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PREFERRED_NAMES_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(PREFERRED_NAMES_SHEET);
    sheet.appendRow(['PreferredName']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(['PreferredName']);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Diagnoses provisioning status for a single player.
 * Returns which sheets have the player and which don't.
 *
 * @param {string} preferredName - The player's name
 * @return {Object} diagnosis
 *   - playerName: string
 *   - presentIn: string[] - sheets where player exists
 *   - missingFrom: string[] - sheets where player is missing
 *   - skipped: string[] - sheets that don't exist or have no key column
 */
function diagnosePlayerProvisioning(preferredName) {
  if (!preferredName || !String(preferredName).trim()) {
    throw new Error('preferredName is required.');
  }

  const name = String(preferredName).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const presentIn = [];
  const missingFrom = [];
  const skipped = [];

  PLAYER_PROVISION_TARGETS.forEach(function(target) {
    const sheet = ss.getSheetByName(target.sheetName);

    if (!sheet) {
      skipped.push(target.sheetName);
      return;
    }

    const keyColIndex = findColumnIndexByHeader_(sheet, target.keyHeaders || PLAYER_KEY_HEADER_NAMES);

    if (keyColIndex === -1) {
      skipped.push(target.sheetName);
      return;
    }

    const existingValues = getColumnValues_(sheet, keyColIndex);
    const playerExists = existingValues.some(function(val) { return val === name; });

    if (playerExists) {
      presentIn.push(target.sheetName);
    } else {
      missingFrom.push(target.sheetName);
    }
  });

  return {
    playerName: name,
    presentIn: presentIn,
    missingFrom: missingFrom,
    skipped: skipped
  };
}