/**
 * Name Hygiene Service - Cosmic Event Manager v7.9.7
 * @fileoverview Detects unknown player names and maps them to canonical names
 *
 * Provides:
 * - scanForUnknownPlayers(): Scans event sheets for names not in PreferredNames
 * - mapPlayerName(oldName, newName): Replaces bad names with canonical names
 */

// ============================================================================
// MAIN SERVICE FUNCTIONS
// ============================================================================

/**
 * Scans event sheets for player names that are not in PreferredNames.
 *
 * @return {Object} result
 *   {
 *     unknowns: Array<UnknownNameRecord>,
 *     canonicalCount: number,
 *     eventSheetsScanned: number
 *   }
 */
function scanForUnknownPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Get canonical names
  const { canonicalSet, canonicalList } = getCanonicalNamesForHygiene_();

  // Step 2: Find event sheets
  const eventSheets = getEventSheets_(ss);

  // Step 3: Scan for unknown names
  const unknowns = [];

  eventSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) return; // No data rows

    // Get header row and find player name column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const playerColIndex = findPlayerNameColumn_(headers);

    if (playerColIndex === -1) return; // No player column found

    // Get all player names in that column
    const playerData = sheet.getRange(2, playerColIndex + 1, lastRow - 1, 1).getValues();

    playerData.forEach((row, rowIdx) => {
      const rawName = String(row[0] || '').trim();

      if (!rawName) return; // Skip blanks

      const nameLower = rawName.toLowerCase();

      // Check if name is in canonical set
      if (!canonicalSet.has(nameLower)) {
        // Get suggestions for this unknown name
        const { suggestedMatch, suggestions } = suggestMatches_(rawName, canonicalList);

        unknowns.push({
          rawName: rawName,
          sheetName: sheetName,
          rowIndex: rowIdx + 2, // 1-based, accounting for header
          colIndex: playerColIndex + 1, // 1-based
          suggestedMatch: suggestedMatch,
          suggestions: suggestions
        });
      }
    });
  });

  return {
    unknowns: unknowns,
    canonicalCount: canonicalList.length,
    eventSheetsScanned: eventSheets.length
  };
}

/**
 * Maps a "bad" player name to a canonical name across sheets.
 *
 * @param {string} oldName - The raw/bad name to replace.
 * @param {string} newName - The desired canonical name.
 * @return {Object} result
 *   {
 *     success: boolean,
 *     message: string,
 *     replacements: Array<{ sheetName: string, count: number }>,
 *     createdNewPlayer?: boolean
 *   }
 */
function mapPlayerName(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Step 1: Validate inputs
  oldName = String(oldName || '').trim();
  newName = String(newName || '').trim();

  if (!oldName || !newName) {
    return {
      success: false,
      message: 'Both old name and new name are required.',
      replacements: [],
      createdNewPlayer: false
    };
  }

  if (oldName.toLowerCase() === newName.toLowerCase()) {
    return {
      success: false,
      message: 'Old and new name are the same.',
      replacements: [],
      createdNewPlayer: false
    };
  }

  // Step 2: Check if newName is canonical, or create it
  const { canonicalSet, canonicalList, canonicalMap } = getCanonicalNamesForHygiene_();
  let createdNewPlayer = false;
  let canonicalNewName = newName;

  const newNameLower = newName.toLowerCase();
  if (canonicalMap.has(newNameLower)) {
    // Use exact canonical spelling
    canonicalNewName = canonicalMap.get(newNameLower);
  } else {
    // Need to add this player
    if (typeof addNewPlayer === 'function') {
      try {
        const res = addNewPlayer(newName);
        if (res && res.success) {
          createdNewPlayer = true;
          canonicalNewName = newName;
        } else {
          // addNewPlayer failed, try direct fallback
          addNewPlayerFallback_(ss, newName);
          createdNewPlayer = true;
          canonicalNewName = newName;
        }
      } catch (e) {
        // addNewPlayer threw an error, try fallback
        addNewPlayerFallback_(ss, newName);
        createdNewPlayer = true;
        canonicalNewName = newName;
      }
    } else {
      // No addNewPlayer function available, use fallback
      addNewPlayerFallback_(ss, newName);
      createdNewPlayer = true;
      canonicalNewName = newName;
    }
  }

  // Step 3: Get target sheets (event sheets + optional mission sheets)
  const targetSheets = getTargetSheetsForMapping_(ss);

  // Step 4: Replace occurrences
  const replacements = [];
  let totalReplacements = 0;

  targetSheets.forEach(sheet => {
    const count = replaceNameInSheet_(sheet, oldName, canonicalNewName);
    if (count > 0) {
      replacements.push({
        sheetName: sheet.getName(),
        count: count
      });
      totalReplacements += count;
    }
  });

  // Step 5: Log the action
  if (totalReplacements > 0) {
    logIntegrityAction('NAME_MAPPING', {
      eventId: 'SYSTEM',
      details: `Mapped "${oldName}" -> "${canonicalNewName}" (${totalReplacements} occurrences)`,
      status: 'SUCCESS'
    });
  }

  // Step 6: Return result
  if (totalReplacements === 0) {
    return {
      success: false,
      message: `No occurrences of "${oldName}" found to replace.`,
      replacements: [],
      createdNewPlayer: createdNewPlayer
    };
  }

  return {
    success: true,
    message: `Replaced ${totalReplacements} occurrence(s) of "${oldName}" with "${canonicalNewName}".`,
    replacements: replacements,
    createdNewPlayer: createdNewPlayer
  };
}

// ============================================================================
// HELPER FUNCTIONS - CANONICAL NAMES
// ============================================================================

/**
 * Gets canonical names from PreferredNames sheet or fallback to Key_Tracker/BP_Total.
 * @return {Object} { canonicalSet: Set, canonicalList: Array, canonicalMap: Map }
 * @private
 */
function getCanonicalNamesForHygiene_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = [];
  const nameSet = new Set();
  const nameMap = new Map(); // lowercase -> original case

  // Try PreferredNames sheet first
  const prefSheet = ss.getSheetByName('PreferredNames');
  if (prefSheet && prefSheet.getLastRow() > 1) {
    const headers = prefSheet.getRange(1, 1, 1, prefSheet.getLastColumn()).getValues()[0];
    const prefCol = headers.findIndex(h => String(h).trim() === 'PreferredName');

    if (prefCol !== -1) {
      const data = prefSheet.getRange(2, prefCol + 1, prefSheet.getLastRow() - 1, 1).getValues();
      data.forEach(row => {
        const name = String(row[0] || '').trim();
        if (name && !nameSet.has(name.toLowerCase())) {
          names.push(name);
          nameSet.add(name.toLowerCase());
          nameMap.set(name.toLowerCase(), name);
        }
      });
    }
  }

  // Fallback: Key_Tracker
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet && keySheet.getLastRow() > 1) {
    const headers = keySheet.getRange(1, 1, 1, keySheet.getLastColumn()).getValues()[0];
    const nameCol = headers.findIndex(h => String(h).trim() === 'PreferredName');
    const colIdx = nameCol !== -1 ? nameCol : 0;

    const data = keySheet.getRange(2, colIdx + 1, keySheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      const name = String(row[0] || '').trim();
      if (name && !nameSet.has(name.toLowerCase())) {
        names.push(name);
        nameSet.add(name.toLowerCase());
        nameMap.set(name.toLowerCase(), name);
      }
    });
  }

  // Fallback: BP_Total
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const headers = bpSheet.getRange(1, 1, 1, bpSheet.getLastColumn()).getValues()[0];
    const nameCol = headers.findIndex(h => String(h).trim() === 'PreferredName');
    const colIdx = nameCol !== -1 ? nameCol : 0;

    const data = bpSheet.getRange(2, colIdx + 1, bpSheet.getLastRow() - 1, 1).getValues();
    data.forEach(row => {
      const name = String(row[0] || '').trim();
      if (name && !nameSet.has(name.toLowerCase())) {
        names.push(name);
        nameSet.add(name.toLowerCase());
        nameMap.set(name.toLowerCase(), name);
      }
    });
  }

  return {
    canonicalSet: nameSet,
    canonicalList: names.sort(),
    canonicalMap: nameMap
  };
}

/**
 * Adds a new player as a fallback when addNewPlayer() is not available.
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} name - Player name to add
 * @private
 */
function addNewPlayerFallback_(ss, name) {
  const timestamp = dateISO ? dateISO() :
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");

  // Add to PreferredNames if it exists
  let prefSheet = ss.getSheetByName('PreferredNames');
  if (prefSheet) {
    const headers = prefSheet.getRange(1, 1, 1, prefSheet.getLastColumn()).getValues()[0];
    const prefCol = headers.findIndex(h => String(h).trim() === 'PreferredName');
    if (prefCol !== -1) {
      const newRow = new Array(headers.length).fill('');
      newRow[prefCol] = name;
      prefSheet.appendRow(newRow);
    }
  }

  // Add to Key_Tracker
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet) {
    keySheet.appendRow([name, 0, 0, 0, 0, 0, 0, timestamp]);
  }

  // Add to BP_Total
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet) {
    bpSheet.appendRow([name, 0, timestamp]);
  }
}

// ============================================================================
// HELPER FUNCTIONS - SHEET DISCOVERY
// ============================================================================

/**
 * Gets all event sheets matching the MM-DD-YYYY or MM-DDx-YYYY pattern.
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array<Sheet>} Event sheets
 * @private
 */
function getEventSheets_(ss) {
  const sheets = ss.getSheets();
  // Pattern: MM-DD-YYYY or MM-DDx-YYYY (x = letter suffix like A, B, etc.)
  const eventPattern = /^\d{2}-\d{2}[A-Z]?-?\d{4}$/i;

  return sheets.filter(sheet => eventPattern.test(sheet.getName()));
}

/**
 * Gets all target sheets for name mapping (event sheets + optional mission sheets).
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array<Sheet>} Target sheets
 * @private
 */
function getTargetSheetsForMapping_(ss) {
  const targets = [];

  // Add event sheets
  const eventSheets = getEventSheets_(ss);
  targets.push(...eventSheets);

  // Add optional mission/tracking sheets
  const optionalSheetNames = [
    'Attendance_Missions',
    'Flag_Missions',
    'Dice_Points',
    'BP_Total',
    'Key_Tracker',
    'Players_Prize-Wall-Points',
    'Event_Outcomes',
    'Prestige_Overflow'
  ];

  optionalSheetNames.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      targets.push(sheet);
    }
  });

  return targets;
}

/**
 * Finds the player name column index (0-based) from headers.
 * @param {Array<string>} headers - Header row values
 * @return {number} Column index (0-based) or -1 if not found
 * @private
 */
function findPlayerNameColumn_(headers) {
  // Try exact matches first (case-insensitive)
  const targetHeaders = ['PreferredName', 'Player', 'Player Name'];

  for (const target of targetHeaders) {
    const idx = headers.findIndex(h =>
      String(h).trim().toLowerCase() === target.toLowerCase()
    );
    if (idx !== -1) return idx;
  }

  // Fallback to column B (index 1)
  if (headers.length > 1) {
    return 1;
  }

  return -1;
}

// ============================================================================
// HELPER FUNCTIONS - NAME REPLACEMENT
// ============================================================================

/**
 * Replaces all occurrences of oldName with newName in a sheet's player column.
 * @param {Sheet} sheet - Target sheet
 * @param {string} oldName - Name to replace
 * @param {string} newName - Replacement name
 * @return {number} Number of replacements made
 * @private
 */
function replaceNameInSheet_(sheet, oldName, newName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0; // No data rows

  // Get headers and find player column
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const playerColIndex = findPlayerNameColumn_(headers);

  if (playerColIndex === -1) return 0;

  // Get all values in player column
  const range = sheet.getRange(2, playerColIndex + 1, lastRow - 1, 1);
  const values = range.getValues();
  let replaceCount = 0;

  // Case-insensitive comparison for matching
  const oldNameLower = oldName.toLowerCase();

  values.forEach((row, idx) => {
    const cellValue = String(row[0] || '').trim();
    if (cellValue.toLowerCase() === oldNameLower) {
      values[idx][0] = newName;
      replaceCount++;
    }
  });

  // Write back if any changes were made
  if (replaceCount > 0) {
    range.setValues(values);
  }

  return replaceCount;
}

// ============================================================================
// HELPER FUNCTIONS - SUGGESTION ENGINE
// ============================================================================

/**
 * Suggests matching canonical names for an unknown name.
 * @param {string} rawName - The unknown name to match
 * @param {Array<string>} canonicalList - List of canonical names
 * @return {Object} { suggestedMatch: string, suggestions: Array<string> }
 * @private
 */
function suggestMatches_(rawName, canonicalList) {
  if (!canonicalList || canonicalList.length === 0) {
    return { suggestedMatch: '', suggestions: [] };
  }

  const rawLower = rawName.toLowerCase();
  const scores = [];

  canonicalList.forEach(canonical => {
    const canonLower = canonical.toLowerCase();

    // Calculate similarity score
    let score = 0;

    // Exact match (shouldn't happen, but just in case)
    if (rawLower === canonLower) {
      score = 1000;
    }
    // Substring match
    else if (canonLower.includes(rawLower) || rawLower.includes(canonLower)) {
      score = 500 + Math.min(rawLower.length, canonLower.length);
    }
    // Levenshtein distance
    else {
      const distance = levenshteinDistance_(rawLower, canonLower);
      const maxLen = Math.max(rawLower.length, canonLower.length);
      // Convert distance to similarity (lower distance = higher score)
      score = Math.max(0, 100 - (distance / maxLen) * 100);
    }

    scores.push({ name: canonical, score: score });
  });

  // Sort by score descending
  scores.sort((a, b) => b.score - a.score);

  // Get top 3 suggestions with reasonable scores
  const threshold = 30; // Minimum similarity score
  const suggestions = scores
    .filter(s => s.score >= threshold)
    .slice(0, 3)
    .map(s => s.name);

  // Best match is first suggestion if score is high enough
  const suggestedMatch = (scores[0] && scores[0].score >= 50) ? scores[0].name : '';

  return {
    suggestedMatch: suggestedMatch,
    suggestions: suggestions
  };
}

/**
 * Computes Levenshtein distance between two strings.
 * @param {string} a - First string
 * @param {string} b - Second string
 * @return {number} Edit distance
 * @private
 */
function levenshteinDistance_(a, b) {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;

  const matrix = [];

  // Initialize matrix
  for (let i = 0; i <= b.length; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= a.length; j++) {
    matrix[0][j] = j;
  }

  // Fill matrix
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1, // substitution
          matrix[i][j - 1] + 1,     // insertion
          matrix[i - 1][j] + 1      // deletion
        );
      }
    }
  }

  return matrix[b.length][a.length];
}

// ============================================================================
// UI ENTRY POINT
// ============================================================================

/**
 * Opens the Detect & Fix Player Names sidebar.
 * Wired from menu: Players -> Detect & Fix Player Names
 */
function onDetectNewPlayers() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/detect_fix_players')
      .setWidth(420);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to open Detect & Fix Players sidebar.\n\nError: ' + e.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    console.error('onDetectNewPlayers error:', e);
  }
}