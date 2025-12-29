/**
 * Event Service - Event Creation and Management
 * @fileoverview Creates events, manages rosters, metadata, and event tabs
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
  SHEETS: {
    ATTENDANCE: 'Attendance_Missions',
    FLAG: 'Flag_Missions',
    DICE: 'Dice_Points',
    BP_TOTAL: 'BP_Total',
    PREFERRED: 'PreferredNames',
    LOG: 'Integrity_Log'
  },
  
  EVENT_PATTERN: /^\d{1,2}-\d{1,2}[a-z]?-\d{4}$/i,
  
  SUFFIXES: {
  'A': 'Academy / Learn to Play',
  'B': 'Casual Commander (Brackets 1–2)',
  'C': 'Transitional Commander (Brackets 3–4)',
  'D': 'Booster Draft',
  'E': 'External / Outreach (offsite / partner events)',
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
  'V': 'Riftbound',
  'W': 'Workshop / Hobby Night',
  'X': 'Multi-Event Day / Multi-Flight',
  'Y': 'Lorcana TCG',
  'Z': 'Staff / Internal Use'
  }
};

// ============================================================================
// DIALOG SUPPORT FUNCTIONS (for HTML dialogs)
// ============================================================================

/**
 * Test connection for HTML dialog
 * @return {string} Connection confirmation
 */
function testConnection() {
  return '✅ Backend connected: ' + new Date().toLocaleString();
}

/**
 * Gets suffix legend for dialogs
 * @return {Array<Object>} Suffix legend
 */
function getSuffixLegend() {
  return Object.entries(CONFIG.SUFFIXES).map(([code, label]) => ({ 
    code, 
    label 
  }));
}

/**
 * Creates event sheet from simple name (for HTML dialog)
 * @param {string} sheetName - Name like "11-12c-2025"
 * @return {string} Created sheet name
 */
function createEventSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Validate format
  if (!CONFIG.EVENT_PATTERN.test(sheetName)) {
    throw new Error('Invalid sheet name format. Expected MM-DD[suffix]-YYYY');
  }
  
  // Check if exists
  if (ss.getSheetByName(sheetName)) {
    throw new Error(`Sheet "${sheetName}" already exists!`);
  }
  
  // Parse event type from name
  const type = parseEventType(sheetName);
  const eventType = CONFIG.SUFFIXES[type] || 'Unknown Event';
  
  // Create sheet
  const sheet = ss.insertSheet(sheetName);
  
  // Set up headers
  const headers = ['Rank', 'PreferredName', 'R1_Prize', 'R2_Prize', 'R3_Prize', 'End_Prizes'];
  sheet.getRange(1, 1, 1, 6).setValues([headers]);
  
  // Format header row
  sheet.getRange('A1:F1')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  // Set column widths
  sheet.setColumnWidth(1, 80);   // Rank
  sheet.setColumnWidth(2, 200);  // PreferredName
  sheet.setColumnWidths(3, 4, 120); // Prize columns
  
  sheet.setFrozenRows(1);
  
  // Add event type label
  sheet.getRange('H1').setValue('Event Type:');
  sheet.getRange('I1').setValue(eventType);
  sheet.getRange('H1').setFontWeight('bold');
  
  // Store metadata
  const seed = generateSeed();
  const eventDate = parseEventDate(sheetName);
  
  setEventProps(sheet, {
    event_type: eventType,
    event_type_code: type,
    event_date: eventDate ? eventDate.toISOString() : new Date().toISOString(),
    event_seed: seed,
    entry: 5, // Default entry fee
    kit_cost_per_player: 0
  });
  
  // Apply row banding
  try {
    sheet.getRange('A2:F100').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  } catch (e) {
    // Banding failed, not critical
  }
  
  // Log creation
  logIntegrityAction('EVENT_CREATE', {
    eventId: sheetName,
    seed,
    details: `Type: ${eventType}`,
    status: 'SUCCESS'
  });
  
  // Activate the new sheet
  ss.setActiveSheet(sheet);
  
  return sheetName;
}

// ============================================================================
// SMART ROSTER PARSER (for HTML dialog)
// ============================================================================

/**
 * Smart parses roster text into clean player names
 * @param {string} text - Raw pasted roster text
 * @return {Array<string>} Clean player names
 */
function smartParseRoster(text) {
  if (!text) return [];
  const lines = text.trim().split(/\r?\n/);
  const players = [];
  
  for (let raw of lines) {
    let s = (raw || "").trim();
    
    // Skip empty lines
    if (!s) continue;
    
    // Skip lines with these keywords (contains, not exact match)
    const skipKeywords = ['eventlink', 'copyright', 'wizards', 'coast', 'report:', 
                          'event:', 'event date:', 'event information:', 'format:', 
                          'structure:', 'opponent', 'match win', 'game win', 
                          'rank', 'omw%', 'gw%', 'ogw%', 'points', '---', '==='];
    const lowerLine = s.toLowerCase();
    if (skipKeywords.some(keyword => lowerLine.includes(keyword))) continue;
    
    // Skip lines with dates/times (contains slashes or colons with PM/AM)
    if (/\d{1,2}\/\d{1,2}\/\d{4}/.test(s)) continue;  // Date format
    if (/\d{1,2}:\d{2}\s*(am|pm)/i.test(s)) continue; // Time format
    
    // Skip lines with URLs or @ symbols
    if (s.includes('http') || s.includes('www.') || s.includes('@')) continue;
    
    // Skip lines that are mostly numbers/symbols
    const letterCount = (s.match(/[a-zA-Z]/g) || []).length;
    if (letterCount < 3) continue; // Must have at least 3 letters
    
    // Now clean the line
    // 1) Strip leading list markers
    s = s.replace(/^\s*(?:[\-\u2022•]+|\(?\d+\)?[.):]?)\s+/, "");
    
    // 2) Collapse tabs/multiple spaces
    s = s.replace(/\s+/g, " ");
    
    // 3) Remove parentheses notes
    s = s.replace(/\s*\([^)]+\)\s*$/g, "");
    
    // 4) Remove trailing W-L-D like "2-1-0" or "9-66-85-69"
    s = s.replace(/\s+\d+(?:-\d+){1,}\s*$/g, "");
    
    // 5) Remove remaining trailing numbers: "John 9 66 85 69" -> "John"
    s = s.replace(/\s+(?:\d+(?:\s+|$))+$/g, "");
    
    // 6) Remove trailing dash notes
    s = s.replace(/\s*[-–—]\s*.*$/g, "");
    
    // 7) Final trim
    s = s.trim();
    if (!s) continue;
    
    // Validate: must be valid name format
    const words = s.split(/\s+/);
    
    // Must have at least 1 word
    if (words.length < 1) continue;
    
    // Each word must be at least 1 char and mostly letters
    // Allow single letter initials (like "Jeremy B")
    let validWords = true;
    for (let word of words) {
      if (word.length === 0) {
        validWords = false;
        break;
      }
      // Word must be either:
      // - All letters (2+ chars): "Jeremy"
      // - Single capital letter (initial): "B"
      // - Has apostrophe or hyphen: "O'Brien", "Jean-Luc"
      if (word.length === 1 && /[A-Z]/.test(word)) continue; // Initial OK
      if (word.includes("'") || word.includes("-")) continue; // O'Brien OK
      if (word.length >= 2 && /^[a-zA-Z]+$/.test(word)) continue; // Normal name OK
      // Otherwise invalid
      validWords = false;
      break;
    }
    if (!validWords) continue;
    
    // Must not contain numbers (after cleaning)
    if (/\d/.test(s)) continue;
    
    // 8) Proper-case
    players.push(properCase(s));
  }
  
  // Deduplicate
  const seen = new Set();
  const unique = [];
  for (const p of players) {
    const lower = p.toLowerCase();
    if (!seen.has(lower)) {
      seen.add(lower);
      unique.push(p);
    }
  }
  return unique;
}

/**
 * Converts name to proper case
 * @param {string} name - Name to convert
 * @return {string} Proper-cased name
 */
function properCase(name) {
  return name.split(' ').map(word => {
    if (!word) return '';
    if (word.includes("'")) {
      return word.split("'").map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join("'");
    }
    if (word.includes('-')) {
      return word.split('-').map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join('-');
    }
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

/**
 * Imports parsed roster to active sheet (for HTML dialog)
 * @param {Array<string>} players - Clean player names
 * @return {Object} Import result
 */
function importParsedRosterToActiveSheet(players) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startRow = Math.max(2, sheet.getLastRow() + 1);
  
  const data = players.map((name, i) => [
    i + 1,       // Rank (clean 1..N)
    name,        // Player Name
    '', '', '', ''  // R1_Prize, R2_Prize, R3_Prize, End_Prizes
  ]);
  
  sheet.getRange(startRow, 1, data.length, 6).setValues(data);
  logIntegrityAction('ROSTER_IMPORT', {
    eventId: sheet.getName(),
    details: `${players.length} players`,
    status: 'SUCCESS'
  });
  
  return { count: players.length };
}

// ============================================================================
// EVENT CREATION (Advanced API)
// ============================================================================

/**
 * Creates a new event tab with full metadata
 * @param {Object} meta - Event metadata
 * @param {string} meta.type - Event type: CONSTRUCTED, LIMITED, HYBRID
 * @param {Date} meta.date - Event date
 * @param {number} meta.entry - Entry fee
 * @param {number} meta.kitCost - Kit cost per player (LIMITED only)
 * @return {Object} {eventId, sheet}
 */
function createEvent(meta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Generate event ID (MM-DD-YYYY format)
  const eventId = formatEventDate(meta.date);

  // Check if tab already exists
  let sheet = ss.getSheetByName(eventId);
  if (sheet) {
    // Add suffix to make unique
    let suffix = 1;
    while (ss.getSheetByName(`${eventId}-${suffix}`)) {
      suffix++;
    }
    sheet = ss.insertSheet(`${eventId}-${suffix}`);
  } else {
    sheet = ss.insertSheet(eventId);
  }

  // Create canonical headers
  const headers = ['Rank', 'PreferredName', 'R1_Prize', 'R2_Prize', 'R3_Prize', 'End_Prizes'];
  sheet.appendRow(headers);

  // Format header row
  sheet.setFrozenRows(1);
  sheet.getRange('A1:F1')
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  // Set column widths
  sheet.setColumnWidth(1, 80);  // Rank
  sheet.setColumnWidth(2, 200); // PreferredName
  sheet.setColumnWidths(3, 4, 120); // Prize columns

  // Store metadata
  const seed = generateSeed();
  const eventProps = {
    event_type: meta.type,
    event_date: meta.date.toISOString(),
    event_seed: seed,
    entry: meta.entry,
    kit_cost_per_player: meta.kitCost || 0
  };

  setEventProps(sheet, eventProps);

  // Log creation
  logIntegrityAction('EVENT_CREATE', {
    eventId: sheet.getName(),
    seed,
    details: `Type: ${meta.type}, Entry: ${formatCurrency(meta.entry)}`,
    status: 'SUCCESS'
  });

  return {
    eventId: sheet.getName(),
    sheet,
    seed
  };
}

// ============================================================================
// EVENT METADATA
// ============================================================================

/**
 * Sets event properties (stored as developer metadata)
 * @param {Sheet} sheet - Event sheet
 * @param {Object} props - Properties to set
 */
function setEventProps(sheet, props) {
  const metadata = sheet.getDeveloperMetadata();

  Object.keys(props).forEach(key => {
    // Remove existing key
    metadata.forEach(m => {
      if (m.getKey() === key) {
        m.remove();
      }
    });

    // Add new value
    sheet.addDeveloperMetadata(key, String(props[key]));
  });
}

/**
 * Gets event properties from metadata
 * @param {Sheet} sheet - Event sheet
 * @return {Object} Event properties
 */
function getEventProps(sheet) {
  const metadata = sheet.getDeveloperMetadata();
  const props = {};

  metadata.forEach(m => {
    const key = m.getKey();
    const value = m.getValue();

    // Coerce types
    if (key === 'entry' || key === 'kit_cost_per_player') {
      props[key] = parseFloat(value) || 0;
    } else {
      props[key] = value;
    }
  });

  return props;
}

/**
 * Gets event property by key
 * @param {Sheet} sheet - Event sheet
 * @param {string} key - Property key
 * @param {*} defaultValue - Default if not found
 * @return {*} Property value
 */
function getEventProp(sheet, key, defaultValue = null) {
  const props = getEventProps(sheet);
  return props[key] !== undefined ? props[key] : defaultValue;
}

// ============================================================================
// ROSTER OPERATIONS (Advanced API)
// ============================================================================

/**
 * Imports roster from paste text with canonical name matching
 * @param {string} eventId - Event tab name
 * @param {string} pasteText - Pasted roster (newline or comma separated)
 * @return {Object} Import result
 */
function rosterImport(eventId, pasteText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    throw new Error(`Event not found: No tab named "${eventId}"`);
  }

  // Parse names
  const lines = pasteText.split(/[\n,]+/).map(line => line.trim()).filter(line => line);
  const canonicalNames = getCanonicalNames();
  const matched = [];
  const unmatched = [];

  lines.forEach(line => {
    const canonical = findCanonicalName_(line, canonicalNames);
    if (canonical) {
      matched.push(canonical);
    } else {
      unmatched.push(line);
    }
  });

  // Write matched to roster
  if (matched.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    const rows = matched.map((name, idx) => [startRow + idx, name, '', '', '', '']);
    sheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
  }

  logIntegrityAction('ROSTER_IMPORT', {
    eventId,
    details: `Imported ${matched.length} matched, ${unmatched.length} unmatched`,
    status: 'SUCCESS'
  });

  return {
    matched,
    unmatched,
    total: lines.length
  };
}

/**
 * Gets canonical names from all player sheets
 * @return {Array<string>} Canonical names
 */
function getCanonicalNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const names = new Set();

  // Key_Tracker
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet) {
    const data = keySheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) names.add(data[i][0]);
    }
  }

  // BP_Total
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet) {
    const data = bpSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) names.add(data[i][0]);
    }
  }

  // PreferredNames (if exists)
  const prefSheet = ss.getSheetByName('PreferredNames');
  if (prefSheet) {
    const data = prefSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) names.add(data[i][0]);
    }
  }

  return Array.from(names).sort();
}

/**
 * Finds canonical name with fuzzy matching
 * @param {string} input - Input name
 * @param {Array<string>} canonicalNames - Canonical names
 * @return {string|null} Canonical name or null
 * @private
 */
function findCanonicalName_(input, canonicalNames) {
  const normalized = input.toLowerCase().trim();

  // Exact match
  for (const name of canonicalNames) {
    if (name.toLowerCase() === normalized) {
      return name;
    }
  }

  // Partial match (contains)
  for (const name of canonicalNames) {
    if (name.toLowerCase().includes(normalized) || normalized.includes(name.toLowerCase())) {
      return name;
    }
  }

  return null;
}

// ============================================================================
// EVENT DISCOVERY
// ============================================================================

/**
 * Validates if a sheet name is an event sheet
 * Supports: MM-DD-YYYY or MM-DDX-YYYY (X = single letter suffix)
 * @param {string} name - Sheet name to validate
 * @return {boolean} True if valid event sheet name
 */
function isEventSheetName_(name) {
  if (!name) return false;

  // Pattern 1: MM-DD-YYYY (e.g., 11-23-2025)
  const plain = /^\d{2}-\d{2}-\d{4}$/;

  // Pattern 2: MM-DDX-YYYY (e.g., 11-23B-2025, where X is single letter)
  const suffixed = /^\d{2}-\d{2}[A-Z]-\d{4}$/;

  return plain.test(name) || suffixed.test(name);
}

/**
 * Parses event ID to extract date, suffix, and id
 * @param {string} sheetName - Event sheet name
 * @return {Object|null} {date: Date, id: string, suffix: string|null}
 */
function parseEventId_(sheetName) {
  if (!isEventSheetName_(sheetName)) {
    return null;
  }

  // Check if suffixed (MM-DDX-YYYY)
  const suffixMatch = sheetName.match(/^(\d{2})-(\d{2})([A-Z])-(\d{4})$/);

  if (suffixMatch) {
    const [, month, day, suffix, year] = suffixMatch;
    const dateStr = `${month}/${day}/${year}`;
    const date = new Date(dateStr);

    return {
      date: date,
      id: sheetName,
      suffix: suffix
    };
  }

  // Plain format (MM-DD-YYYY)
  const plainMatch = sheetName.match(/^(\d{2})-(\d{2})-(\d{4})$/);

  if (plainMatch) {
    const [, month, day, year] = plainMatch;
    const dateStr = `${month}/${day}/${year}`;
    const date = new Date(dateStr);

    return {
      date: date,
      id: sheetName,
      suffix: null
    };
  }

  return null;
}

/**
 * Gets all valid event sheets
 * @return {Array<Sheet>} Array of event sheets
 */
function getAllEventSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().filter(s => isEventSheetName_(s.getName()));
}

/**
 * Lists all event tabs (MM-DD-YYYY and MM-DDX-YYYY formats)
 * @return {Array<string>} Event tab names, sorted
 */
function listEventTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  return sheets
    .map(s => s.getName())
    .filter(name => isEventSheetName_(name))
    .sort();
}

/**
 * Gets event index (all events with metadata)
 * @return {Array<Object>} Event objects
 */
function indexEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventTabs = listEventTabs();
  const events = [];

  eventTabs.forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;

    const props = getEventProps(sheet);
    const playerCount = sheet.getLastRow() - 1; // Exclude header

    events.push({
      eventId: tabName,
      eventDate: props.event_date || '',
      eventType: props.event_type || 'CONSTRUCTED',
      playerCount,
      entry: props.entry || 0,
      kitCost: props.kit_cost_per_player || 0,
      seed: props.event_seed || ''
    });
  });

  return events;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Parses event date from sheet name
 * @param {string} name - Sheet name like "11-12c-2025"
 * @return {Date|null} Parsed date
 */
function parseEventDate(name) {
  const match = name.match(/^(\d{1,2})-(\d{1,2})([a-z])?-(\d{4})$/i);
  if (!match) return null;
  return new Date(parseInt(match[4]), parseInt(match[1]) - 1, parseInt(match[2]));
}

/**
 * Parses event type suffix from sheet name
 * @param {string} name - Sheet name
 * @return {string} Event type code (A-Z or empty)
 */
function parseEventType(name) {
  const match = name.match(/[a-z](?=-\d{4}$)/i);
  return match ? match[0].toUpperCase() : '';
}

/**
 * Formats date as MM-DD-YYYY
 * @param {Date} date - Date to format
 * @return {string} Formatted date
 */
function formatEventDate(date) {
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${mm}-${dd}-${yyyy}`;
}

/**
 * Generates random seed for event
 * @return {string} Random seed
 */
function generateSeed() {
  return Math.random().toString(36).substring(2, 10).toUpperCase();
}

/**
 * Formats currency
 * @param {number} amount - Amount to format
 * @return {string} Formatted currency
 */
function formatCurrency(amount) {
  return '$' + amount.toFixed(2);
}

/**
 * Logs integrity action
 * @param {string} action - Action type
 * @param {Object} data - Action data
 */
function logIntegrityAction(action, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(CONFIG.SHEETS.LOG);
  
  if (!log) {
    log = ss.insertSheet(CONFIG.SHEETS.LOG);
    log.getRange('A1:F1').setValues([
      ['Timestamp', 'StoreID', 'User', 'Action', 'Target', 'Details']
    ]).setBackground('#4285f4').setFontColor('#ffffff').setFontWeight('bold');
  }
  
  const timestamp = new Date();
  const user = Session.getActiveUser().getEmail() || 'System';
  const storeID = ss.getName().match(/\[(.*?)\]/)?.[1] || 'COSMIC';
  const checksum = (timestamp.getTime() % 10000).toString(16);
  
  log.appendRow([
    timestamp, 
    storeID, 
    user, 
    action, 
    data.eventId || data.target || '', 
    `${data.details || ''} [${checksum}]`
  ]);
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures event tab has canonical schema
 * @param {Sheet} sheet - Event sheet
 */
function ensureEventSchema(sheet) {
  const requiredHeaders = ['Rank', 'PreferredName', 'R1_Prize', 'R2_Prize', 'R3_Prize', 'End_Prizes'];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Check for missing headers
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}