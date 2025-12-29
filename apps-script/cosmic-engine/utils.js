/**
 * Utility Functions for Cosmic Event Manager
 * @fileoverview PRNG, hashing, validation, and data transformation utilities
 */

// ============================================================================
// PRNG & SEEDING
// ============================================================================

/**
 * Generates a 10-character base62 alphanumeric seed
 * @return {string} 10-char seed
 */
function generateSeed() {
  const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  let seed = '';
  for (let i = 0; i < 10; i++) {
    seed += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return seed;
}

/**
 * Seeded PRNG using simple LCG (Linear Congruential Generator)
 * @param {string} seed - Alphanumeric seed string
 * @return {Function} PRNG function that returns 0-1
 */
function createSeededRandom(seed) {
  // Convert seed to numeric hash
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = ((hash << 5) - hash) + seed.charCodeAt(i);
    hash = hash & hash; // Convert to 32-bit integer
  }

  // LCG parameters
  let state = Math.abs(hash);

  return function() {
    state = (state * 1664525 + 1013904223) % 4294967296;
    return state / 4294967296;
  };
}

/**
 * Shuffles array using seeded random
 * @param {Array} array - Array to shuffle
 * @param {Function} rng - Seeded random function
 * @return {Array} Shuffled array
 */
function seededShuffle(array, rng) {
  const arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// ============================================================================
// HASHING & CHECKSUMS
// ============================================================================

/**
 * Computes SHA-256 hash of string
 * @param {string} input - String to hash
 * @return {string} Hex hash
 */
function sha256(input) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    input,
    Utilities.Charset.UTF_8
  );

  return rawHash.map(byte => {
    const v = (byte < 0) ? 256 + byte : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

/**
 * Computes checksum for range data
 * @param {Array<Array>} rangeValues - 2D array from getValues()
 * @return {string} Checksum hash
 */
function computeChecksum(rangeValues) {
  const serialized = JSON.stringify(rangeValues);
  return sha256(serialized).substring(0, 12); // First 12 chars
}

/**
 * Computes hash for preview object (deterministic)
 * @param {Object} obj - Object to hash
 * @return {string} Hash string
 */
function computeHash(obj) {
  // Stringify with sorted keys for determinism
  const serialized = JSON.stringify(obj, Object.keys(obj).sort());
  return sha256(serialized).substring(0, 16);
}

// ============================================================================
// DATE & TIME
// ============================================================================

/**
 * Returns current timestamp in ISO format
 * @return {string} ISO timestamp
 */
function dateISO() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ss'Z'"
  );
}

/**
 * Formats date as MM-DD-YYYY
 * @param {Date} date - Date object
 * @return {string} Formatted date
 */
function formatEventDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM-dd-yyyy');
}

/**
 * Parses ISO date string
 * @param {string} isoString - ISO date string
 * @return {Date} Date object
 */
function parseISODate(isoString) {
  return new Date(isoString);
}

// ============================================================================
// DATA PARSING & COERCION
// ============================================================================

/**
 * Parses CSV text
 * @param {string} csvText - CSV content
 * @param {string} delimiter - Delimiter (default: ',')
 * @return {Array<Array<string>>} Parsed rows
 */
function parseCSV(csvText, delimiter = ',') {
  const rows = [];
  const lines = csvText.split(/\r?\n/);

  for (const line of lines) {
    if (!line.trim()) continue;

    const row = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      const next = line[i + 1];

      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          i++; // Skip next quote
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === delimiter && !inQuotes) {
        row.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    row.push(current);
    rows.push(row);
  }

  return rows;
}

/**
 * Converts 2D array to key-value object (2-column format)
 * @param {Array<Array>} data - 2D array [[key, value], ...]
 * @return {Object} KV object
 */
function toKV(data) {
  const kv = {};
  for (let i = 1; i < data.length; i++) { // Skip header
    const [key, value] = data[i];
    if (key) kv[key] = value;
  }
  return kv;
}

/**
 * Converts 2D array to array of objects using first row as headers
 * @param {Array<Array>} data - 2D array with header row
 * @return {Array<Object>} Array of row objects
 */
function toObjects(data) {
  if (data.length === 0) return [];
  const headers = data[0];
  const objects = [];

  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    objects.push(obj);
  }

  return objects;
}

/**
 * Converts array of objects to Map keyed by specified field
 * @param {Array<Object>} objects - Array of objects
 * @param {string} keyField - Field to use as key
 * @return {Map} Map of objects
 */
function toMapByKey(objects, keyField) {
  const map = new Map();
  objects.forEach(obj => {
    if (obj[keyField] !== undefined) {
      map.set(obj[keyField], obj);
    }
  });
  return map;
}

// ============================================================================
// VALIDATION & COERCION
// ============================================================================

/**
 * Coerces value to number, returns default if invalid
 * @param {*} value - Value to coerce
 * @param {number} defaultValue - Default if invalid
 * @return {number} Coerced number
 */
function coerceNumber(value, defaultValue = 0) {
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * Coerces value to boolean
 * @param {*} value - Value to coerce
 * @return {boolean} Boolean value
 */
function coerceBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const upper = value.toUpperCase();
    return upper === 'TRUE' || upper === 'YES' || upper === '1';
  }
  return Boolean(value);
}

/**
 * Validates and parses ratio string (e.g., "3:1")
 * @param {string} ratioStr - Ratio string "X:Y"
 * @return {Object} {num, den} or null if invalid
 */
function parseRatio(ratioStr) {
  if (!ratioStr || typeof ratioStr !== 'string') return null;

  const match = ratioStr.match(/^(\d+):(\d+)$/);
  if (!match) return null;

  const num = parseInt(match[1], 10);
  const den = parseInt(match[2], 10);

  if (den === 0) return null;

  return { num, den };
}

/**
 * Clamps number between min and max
 * @param {number} value - Value to clamp
 * @param {number} min - Minimum
 * @param {number} max - Maximum
 * @return {number} Clamped value
 */
function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

// ============================================================================
// SYNONYM MAPPING
// ============================================================================

/**
 * Synonym map for canonical header names
 * @return {Object} Map of synonym -> canonical
 */
function getHeaderSynonyms() {
  return {
    // Prize_Catalog
    'SKU': 'Code',
    'Product_Name': 'Name',
    'Product': 'Name',
    'Stock_Qty': 'Qty',
    'Quantity': 'Qty',
    'Retail_Value': 'EV_Cost',
    'Cost_Basis': 'COGS',
    'Cost': 'COGS',
    'L_Bucket': 'Level',
    'Tier': 'Level',
    'In_Stock': 'InStock',

    // Event tabs
    'Preferred_Name': 'PreferredName',
    'Player_Name': 'PreferredName',
    'Name': 'PreferredName',
    'Player': 'PreferredName',

    // Key_Tracker
    'Preferred_Name': 'PreferredName',
    'Rainbow_Eligible': 'RainbowEligible',
    'Last_Updated': 'LastUpdated',

    // BP_Total
    'BP': 'Current_BP',
    'Bonus_Points': 'Current_BP',
    'BP_Current': 'Current_BP'  // Legacy mapping
  };
}

/**
 * Normalizes header using synonym map
 * @param {string} header - Original header
 * @return {string} Canonical header
 */
function normalizeHeader(header) {
  const synonyms = getHeaderSynonyms();
  return synonyms[header] || header;
}

// ============================================================================
// ERROR HANDLING
// ============================================================================

/**
 * Throws formatted error with code
 * @param {string} message - Error message
 * @param {string} code - Error code (e.g., 'INVALID_INPUT')
 * @param {string} remediation - Suggested fix
 * @throws {Error} Formatted error
 */
function throwError(message, code, remediation = '') {
  const fullMessage = `[${code}] ${message}${remediation ? '\n\nFix: ' + remediation : ''}`;
  throw new Error(fullMessage);
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Formats number as currency
 * @param {number} value - Numeric value
 * @return {string} Formatted currency
 */
function formatCurrency(value) {
  return '$' + value.toFixed(2);
}

/**
 * Formats percentage
 * @param {number} value - Decimal value (0.95 = 95%)
 * @param {number} decimals - Decimal places
 * @return {string} Formatted percentage
 */
function formatPercent(value, decimals = 0) {
  return (value * 100).toFixed(decimals) + '%';
}

/**
 * Gets current user email
 * @return {string} User email
 */
function currentUser() {
  return Session.getActiveUser().getEmail() || 'unknown';
}

// ============================================================================
// ARRAY UTILITIES
// ============================================================================

/**
 * Groups array by key function
 * @param {Array} array - Array to group
 * @param {Function} keyFn - Function that returns group key
 * @return {Map} Map of key -> array of items
 */
function groupBy(array, keyFn) {
  const groups = new Map();
  array.forEach(item => {
    const key = keyFn(item);
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(item);
  });
  return groups;
}

/**
 * Sums array by value function
 * @param {Array} array - Array to sum
 * @param {Function} valueFn - Function that returns numeric value
 * @return {number} Sum
 */
function sumBy(array, valueFn) {
  return array.reduce((sum, item) => sum + valueFn(item), 0);
}

/**
 * Filters unique values
 * @param {Array} array - Array with duplicates
 * @return {Array} Unique values
 */
function unique(array) {
  return [...new Set(array)];
}

/**
 * Deep clones object/array
 * @param {*} obj - Object to clone
 * @return {*} Cloned object
 */
function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}