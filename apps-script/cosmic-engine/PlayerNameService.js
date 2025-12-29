/**
 * ════════════════════════════════════════════════════════════════════════════
 * PLAYER NAME SERVICE v1.0.0
 * ════════════════════════════════════════════════════════════════════════════
 *
 * @fileoverview Player name validation, spell checking, and correction
 *
 * Features:
 *   - Scans all event sheets for unrecognized player names
 *   - Fuzzy matching to suggest likely corrections
 *   - Levenshtein distance + phonetic matching
 *   - Bulk correction across all sheets
 *   - UndiscoveredNames tracking with suggestions
 *   - HTML dialog for review and approval
 *
 * Public API:
 *   - runPlayerNameCheck() - Main scan, returns mismatches with suggestions
 *   - applyNameCorrection(oldName, newName) - Fix name across all sheets
 *   - addToPreferredNames(name) - Add new canonical name
 *   - getUndiscoveredNames() - Get pending names for review
 *   - onPlayerNameChecker() - Open HTML dialog
 *
 * Compatible with: Engine v7.9.6+
 * ════════════════════════════════════════════════════════════════════════════
 */

// ════════════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ════════════════════════════════════════════════════════════════════════════

const NAME_SERVICE_CONFIG = {
  // Event sheet pattern (same as MissionScannerService)
  EVENT_PATTERN: /^(\d{1,2})-(\d{1,2})([A-Za-z])?-(\d{4})$/,
  
  // Player column headers to search (case-insensitive)
  PLAYER_COLUMNS: ['preferredname', 'preferred_name_id', 'player', 'name', 'player name'],
  
  // Sheets
  SHEETS: {
    PREFERRED_NAMES: 'PreferredNames',
    UNDISCOVERED: 'UndiscoveredNames',
    KEY_TRACKER: 'Key_Tracker',
    BP_TOTAL: 'BP_Total',
    INTEGRITY_LOG: 'Integrity_Log'
  },
  
  // Fuzzy match settings
  FUZZY: {
    MAX_DISTANCE: 3,           // Max Levenshtein distance for suggestions
    MIN_SIMILARITY: 0.6,       // Min similarity score (0-1) for suggestions
    MAX_SUGGESTIONS: 5         // Max suggestions per unknown name
  }
};

// ════════════════════════════════════════════════════════════════════════════
// MAIN SCAN FUNCTION
// ════════════════════════════════════════════════════════════════════════════

/**
 * Run full player name check across all event sheets
 * @return {Object} {totalNames, matched, unmatched, suggestions}
 */
function runPlayerNameCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('═══════════════════════════════════════════════════════');
  console.log('PLAYER NAME SERVICE - Starting scan...');
  console.log('═══════════════════════════════════════════════════════');
  
  // Load canonical names
  const preferredNames = loadPreferredNamesList_(ss);
  console.log(`Loaded ${preferredNames.length} preferred names`);
  
  // Scan all event sheets
  const allNames = scanAllEventNames_(ss);
  console.log(`Found ${allNames.size} unique names across events`);
  
  // Categorize: matched vs unmatched
  const matched = [];
  const unmatched = [];
  
  allNames.forEach((occurrences, rawName) => {
    const canonical = findExactMatch_(rawName, preferredNames);
    if (canonical) {
      matched.push({ raw: rawName, canonical: canonical, count: occurrences.length });
    } else {
      // Find fuzzy suggestions
      const suggestions = findFuzzySuggestions_(rawName, preferredNames);
      unmatched.push({
        raw: rawName,
        occurrences: occurrences, // [{sheet, row}, ...]
        count: occurrences.length,
        suggestions: suggestions
      });
    }
  });
  
  // Update UndiscoveredNames sheet
  updateUndiscoveredNames_(ss, unmatched);
  
  // Log results
  logIntegrityAction('NAME_CHECK', {
    details: `Scanned ${allNames.size} names: ${matched.length} matched, ${unmatched.length} unmatched`,
    status: unmatched.length === 0 ? 'SUCCESS' : 'NEEDS_REVIEW'
  });
  
  console.log(`Results: ${matched.length} matched, ${unmatched.length} need review`);
  
  return {
    totalNames: allNames.size,
    matched: matched.length,
    unmatched: unmatched,
    suggestions: unmatched.map(u => ({
      name: u.raw,
      count: u.count,
      suggestions: u.suggestions.map(s => s.name)
    }))
  };
}

/**
 * Scan all event sheets and collect player names with locations
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Map<string, Array>} Map of rawName -> [{sheet, row}, ...]
 * @private
 */
function scanAllEventNames_(ss) {
  const sheets = ss.getSheets();
  const nameMap = new Map(); // rawName -> [{sheet, row}, ...]
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Check if event sheet
    if (!NAME_SERVICE_CONFIG.EVENT_PATTERN.test(sheetName)) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    // Find player column
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    let playerCol = -1;
    
    for (const colName of NAME_SERVICE_CONFIG.PLAYER_COLUMNS) {
      const idx = headers.indexOf(colName);
      if (idx !== -1) {
        playerCol = idx;
        break;
      }
    }
    
    if (playerCol === -1) return;
    
    // Extract names
    for (let i = 1; i < data.length; i++) {
      const rawName = String(data[i][playerCol] || '').trim();
      if (!rawName) continue;
      
      if (!nameMap.has(rawName)) {
        nameMap.set(rawName, []);
      }
      nameMap.get(rawName).push({
        sheet: sheetName,
        row: i + 1 // 1-indexed for user display
      });
    }
  });
  
  return nameMap;
}

/**
 * Load PreferredNames as array
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Array<string>} List of canonical names
 * @private
 */
function loadPreferredNamesList_(ss) {
  const sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.PREFERRED_NAMES);
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const names = [];
  
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][0] || '').trim();
    if (name) names.push(name);
  }
  
  return names;
}

/**
 * Find exact match (case-insensitive)
 * @param {string} rawName - Name to match
 * @param {Array<string>} preferredNames - Canonical names
 * @return {string|null} Matched canonical name or null
 * @private
 */
function findExactMatch_(rawName, preferredNames) {
  const lower = rawName.toLowerCase();
  
  for (const canonical of preferredNames) {
    if (canonical.toLowerCase() === lower) {
      return canonical;
    }
  }
  
  return null;
}

// ════════════════════════════════════════════════════════════════════════════
// FUZZY MATCHING
// ════════════════════════════════════════════════════════════════════════════

/**
 * Find fuzzy suggestions for unknown name
 * @param {string} rawName - Unknown name
 * @param {Array<string>} preferredNames - Canonical names
 * @return {Array<Object>} [{name, distance, similarity, phonetic}, ...]
 * @private
 */
function findFuzzySuggestions_(rawName, preferredNames) {
  const suggestions = [];
  const config = NAME_SERVICE_CONFIG.FUZZY;
  
  const rawLower = rawName.toLowerCase();
  const rawSoundex = soundex_(rawName);
  const rawParts = rawLower.split(/\s+/);
  
  preferredNames.forEach(canonical => {
    const canonLower = canonical.toLowerCase();
    const canonSoundex = soundex_(canonical);
    const canonParts = canonLower.split(/\s+/);
    
    // Calculate Levenshtein distance
    const distance = levenshteinDistance_(rawLower, canonLower);
    
    // Calculate similarity (0-1)
    const maxLen = Math.max(rawLower.length, canonLower.length);
    const similarity = 1 - (distance / maxLen);
    
    // Check phonetic match
    const phoneticMatch = rawSoundex === canonSoundex;
    
    // Check partial matches (first/last name swaps, etc.)
    const partialScore = calculatePartialMatch_(rawParts, canonParts);
    
    // Combined score
    const combinedScore = (similarity * 0.5) + (phoneticMatch ? 0.25 : 0) + (partialScore * 0.25);
    
    if (distance <= config.MAX_DISTANCE || similarity >= config.MIN_SIMILARITY || phoneticMatch || partialScore > 0.5) {
      suggestions.push({
        name: canonical,
        distance: distance,
        similarity: similarity,
        phonetic: phoneticMatch,
        partialScore: partialScore,
        score: combinedScore
      });
    }
  });
  
  // Sort by combined score (descending) and limit
  suggestions.sort((a, b) => b.score - a.score);
  return suggestions.slice(0, config.MAX_SUGGESTIONS);
}

/**
 * Calculate Levenshtein distance between two strings
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

/**
 * Calculate Soundex phonetic code
 * @param {string} name - Name to encode
 * @return {string} Soundex code
 * @private
 */
function soundex_(name) {
  const a = name.toLowerCase().split('');
  const f = a.shift();
  
  const codes = {
    a: '', e: '', i: '', o: '', u: '',
    b: 1, f: 1, p: 1, v: 1,
    c: 2, g: 2, j: 2, k: 2, q: 2, s: 2, x: 2, z: 2,
    d: 3, t: 3,
    l: 4,
    m: 5, n: 5,
    r: 6
  };
  
  const result = a
    .map(c => codes[c])
    .filter((c, i, arr) => c !== '' && c !== arr[i - 1])
    .join('');
  
  return (f + result + '000').slice(0, 4).toUpperCase();
}

/**
 * Calculate partial match score for name parts
 * Handles first/last name swaps, initials, etc.
 * @param {Array<string>} parts1 - First name parts
 * @param {Array<string>} parts2 - Second name parts
 * @return {number} Score 0-1
 * @private
 */
function calculatePartialMatch_(parts1, parts2) {
  if (parts1.length === 0 || parts2.length === 0) return 0;
  
  let matchedParts = 0;
  const totalParts = Math.max(parts1.length, parts2.length);
  
  parts1.forEach(p1 => {
    for (const p2 of parts2) {
      // Exact match
      if (p1 === p2) {
        matchedParts++;
        break;
      }
      // Initial match (e.g., "J" matches "John")
      if (p1.length === 1 && p2.startsWith(p1)) {
        matchedParts += 0.5;
        break;
      }
      if (p2.length === 1 && p1.startsWith(p2)) {
        matchedParts += 0.5;
        break;
      }
      // Prefix match (e.g., "Jon" matches "Jonathan")
      if (p1.length >= 3 && p2.startsWith(p1)) {
        matchedParts += 0.75;
        break;
      }
      if (p2.length >= 3 && p1.startsWith(p2)) {
        matchedParts += 0.75;
        break;
      }
    }
  });
  
  return matchedParts / totalParts;
}

// ════════════════════════════════════════════════════════════════════════════
// UNDISCOVERED NAMES MANAGEMENT
// ════════════════════════════════════════════════════════════════════════════

/**
 * Update UndiscoveredNames sheet with unmatched names and suggestions
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {Array<Object>} unmatched - Unmatched names with suggestions
 * @private
 */
function updateUndiscoveredNames_(ss, unmatched) {
  let sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.UNDISCOVERED);
  
  if (!sheet) {
    sheet = ss.insertSheet(NAME_SERVICE_CONFIG.SHEETS.UNDISCOVERED);
  }
  
  // Clear and rebuild
  sheet.clear();
  
  // Headers
  const headers = [
    'Potential_Name',
    'Occurrences',
    'First_Seen_Sheet',
    'Best_Match',
    'Match_Score',
    'All_Suggestions',
    'Status',
    'Resolved_To',
    'Timestamp'
  ];
  
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);
  sheet.getRange('A1:I1')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Add unmatched names
  unmatched.forEach(item => {
    const bestMatch = item.suggestions.length > 0 ? item.suggestions[0] : null;
    const allSuggestions = item.suggestions.map(s => s.name).join(', ');
    
    sheet.appendRow([
      item.raw,
      item.count,
      item.occurrences[0]?.sheet || '',
      bestMatch ? bestMatch.name : '',
      bestMatch ? (bestMatch.score * 100).toFixed(0) + '%' : '',
      allSuggestions,
      'PENDING',
      '',
      new Date().toISOString()
    ]);
  });
  
  // Format
  sheet.autoResizeColumns(1, headers.length);
  
  // Add data validation for Status column
  if (unmatched.length > 0) {
    const statusRange = sheet.getRange(2, 7, unmatched.length, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['PENDING', 'APPROVED', 'CORRECTED', 'NEW_PLAYER', 'IGNORED'])
      .build();
    statusRange.setDataValidation(statusRule);
  }
  
  console.log(`Updated UndiscoveredNames with ${unmatched.length} entries`);
}

/**
 * Get undiscovered names for HTML dialog
 * @return {Array<Object>} Names with suggestions and status
 */
function getUndiscoveredNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.UNDISCOVERED);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[6] === 'PENDING') { // Only pending items
      results.push({
        row: i + 1,
        name: row[0],
        occurrences: row[1],
        firstSheet: row[2],
        bestMatch: row[3],
        matchScore: row[4],
        allSuggestions: row[5] ? row[5].split(', ') : [],
        status: row[6]
      });
    }
  }
  
  return results;
}

// ════════════════════════════════════════════════════════════════════════════
// CORRECTION FUNCTIONS
// ════════════════════════════════════════════════════════════════════════════

/**
 * Apply name correction across all event sheets
 * @param {string} oldName - Name to replace
 * @param {string} newName - Canonical name to use
 * @return {Object} {sheetsFixed, cellsFixed}
 */
function applyNameCorrection(oldName, newName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  let sheetsFixed = 0;
  let cellsFixed = 0;
  
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Only process event sheets
    if (!NAME_SERVICE_CONFIG.EVENT_PATTERN.test(sheetName)) return;
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    // Find player column
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    let playerCol = -1;
    
    for (const colName of NAME_SERVICE_CONFIG.PLAYER_COLUMNS) {
      const idx = headers.indexOf(colName);
      if (idx !== -1) {
        playerCol = idx;
        break;
      }
    }
    
    if (playerCol === -1) return;
    
    let sheetModified = false;
    
    // Find and replace
    for (let i = 1; i < data.length; i++) {
      const cellValue = String(data[i][playerCol] || '').trim();
      
      // Case-insensitive match
      if (cellValue.toLowerCase() === oldName.toLowerCase()) {
        sheet.getRange(i + 1, playerCol + 1).setValue(newName);
        cellsFixed++;
        sheetModified = true;
      }
    }
    
    if (sheetModified) sheetsFixed++;
  });
  
  // Update UndiscoveredNames status
  markNameResolved_(ss, oldName, newName, 'CORRECTED');
  
  // Log the correction
  logIntegrityAction('NAME_CORRECTION', {
    details: `"${oldName}" → "${newName}" in ${sheetsFixed} sheets (${cellsFixed} cells)`,
    status: 'SUCCESS'
  });
  
  console.log(`Corrected "${oldName}" → "${newName}": ${sheetsFixed} sheets, ${cellsFixed} cells`);
  
  return { sheetsFixed, cellsFixed };
}

/**
 * Add a new name to PreferredNames
 * @param {string} name - Name to add
 * @return {boolean} Success
 */
function addToPreferredNames(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.PREFERRED_NAMES);
  
  if (!sheet) {
    sheet = ss.insertSheet(NAME_SERVICE_CONFIG.SHEETS.PREFERRED_NAMES);
    sheet.appendRow(['PreferredName']);
    sheet.setFrozenRows(1);
  }
  
  // Check if already exists
  const existingNames = loadPreferredNamesList_(ss);
  const lower = name.toLowerCase();
  
  for (const existing of existingNames) {
    if (existing.toLowerCase() === lower) {
      console.log(`Name "${name}" already exists in PreferredNames`);
      return false;
    }
  }
  
  // Add new name
  sheet.appendRow([name]);
  
  // Update UndiscoveredNames status
  markNameResolved_(ss, name, name, 'NEW_PLAYER');
  
  // Log
  logIntegrityAction('NAME_ADD', {
    details: `Added "${name}" to PreferredNames`,
    status: 'SUCCESS'
  });
  
  console.log(`Added "${name}" to PreferredNames`);
  return true;
}

/**
 * Mark a name as resolved in UndiscoveredNames
 * @param {Spreadsheet} ss - Active spreadsheet
 * @param {string} originalName - Original name
 * @param {string} resolvedTo - Resolution
 * @param {string} status - Status code
 * @private
 */
function markNameResolved_(ss, originalName, resolvedTo, status) {
  const sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.UNDISCOVERED);
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === originalName) {
      sheet.getRange(i + 1, 7).setValue(status);      // Status column
      sheet.getRange(i + 1, 8).setValue(resolvedTo);  // Resolved_To column
      break;
    }
  }
}

/**
 * Ignore a name (mark as not needing resolution)
 * @param {string} name - Name to ignore
 * @return {boolean} Success
 */
function ignoreUndiscoveredName(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  markNameResolved_(ss, name, '', 'IGNORED');
  
  logIntegrityAction('NAME_IGNORE', {
    details: `Ignored "${name}"`,
    status: 'SUCCESS'
  });
  
  return true;
}

// ════════════════════════════════════════════════════════════════════════════
// BULK OPERATIONS
// ════════════════════════════════════════════════════════════════════════════

/**
 * Auto-correct names with high-confidence matches
 * @param {number} minScore - Minimum match score (0-100) to auto-correct
 * @return {Object} {corrected, skipped}
 */
function autoCorrectHighConfidence(minScore = 90) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.UNDISCOVERED);
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { corrected: 0, skipped: 0 };
  }
  
  const data = sheet.getDataRange().getValues();
  let corrected = 0;
  let skipped = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[6];
    const matchScore = parseInt(row[4]) || 0;
    const bestMatch = row[3];
    
    if (status !== 'PENDING') continue;
    
    if (matchScore >= minScore && bestMatch) {
      // Auto-correct
      const oldName = row[0];
      applyNameCorrection(oldName, bestMatch);
      corrected++;
    } else {
      skipped++;
    }
  }
  
  return { corrected, skipped };
}

/**
 * Get name statistics
 * @return {Object} Stats summary
 */
function getNameStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const preferredNames = loadPreferredNamesList_(ss);
  const allNames = scanAllEventNames_(ss);
  
  let matched = 0;
  let unmatched = 0;
  
  allNames.forEach((occurrences, rawName) => {
    const canonical = findExactMatch_(rawName, preferredNames);
    if (canonical) {
      matched++;
    } else {
      unmatched++;
    }
  });
  
  return {
    totalPreferredNames: preferredNames.length,
    totalUniqueNamesInEvents: allNames.size,
    matched: matched,
    unmatched: unmatched,
    matchRate: allNames.size > 0 ? ((matched / allNames.size) * 100).toFixed(1) + '%' : '100%'
  };
}

// ════════════════════════════════════════════════════════════════════════════
// HTML DIALOG SUPPORT
// ════════════════════════════════════════════════════════════════════════════

/**
 * Open Player Name Checker dialog
 * Menu handler: Players > Check Player Names
 */
function onPlayerNameChecker() {
  const html = HtmlService.createHtmlOutputFromFile('PlayerNameChecker')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Player Name Checker');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Player Name Checker');
}

/**
 * Run scan and return results for HTML dialog
 * @return {Object} Scan results
 */
function runNameScanForDialog() {
  const results = runPlayerNameCheck();
  const stats = getNameStats();
  
  return {
    stats: stats,
    unmatched: results.unmatched.map(item => ({
      name: item.raw,
      count: item.count,
      firstSheet: item.occurrences[0]?.sheet || '',
      suggestions: item.suggestions.map(s => ({
        name: s.name,
        score: (s.score * 100).toFixed(0)
      }))
    }))
  };
}

/**
 * Apply correction from HTML dialog
 * @param {string} oldName - Original name
 * @param {string} newName - Corrected name
 * @return {Object} Result
 */
function applyNameCorrectionFromDialog(oldName, newName) {
  try {
    const result = applyNameCorrection(oldName, newName);
    return { success: true, ...result };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Add new player from HTML dialog
 * @param {string} name - Name to add
 * @return {Object} Result
 */
function addNewPlayerFromDialog(name) {
  try {
    const success = addToPreferredNames(name);
    return { success: success };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Ignore name from HTML dialog
 * @param {string} name - Name to ignore
 * @return {Object} Result
 */
function ignoreNameFromDialog(name) {
  try {
    ignoreUndiscoveredName(name);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════════════════════
// UTILITY
// ════════════════════════════════════════════════════════════════════════════

/**
 * Log integrity action (shared with other services)
 * @param {string} action - Action code
 * @param {Object} data - Action data
 */
function logIntegrityAction(action, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let log = ss.getSheetByName(NAME_SERVICE_CONFIG.SHEETS.INTEGRITY_LOG);
    
    if (!log) {
      log = ss.insertSheet(NAME_SERVICE_CONFIG.SHEETS.INTEGRITY_LOG);
      log.appendRow(['Timestamp', 'User', 'Action', 'Target', 'Details', 'Status']);
      log.setFrozenRows(1);
    }
    
    const timestamp = new Date().toISOString();
    const user = Session.getActiveUser().getEmail() || 'System';
    
    log.appendRow([
      timestamp,
      user,
      action,
      data.target || '',
      data.details || '',
      data.status || ''
    ]);
  } catch (e) {
    console.error('Failed to log integrity action:', e);
  }
}

// ════════════════════════════════════════════════════════════════════════════
// END OF PLAYER NAME SERVICE v1.0.0
// ════════════════════════════════════════════════════════════════════════════