/**
 * LockboxService.gs
 * Cosmic Games Tournament Manager v7.9.7
 *
 * Service functions for the Unlock Lockbox flow.
 * Handles player key status retrieval, eligibility verification,
 * and lockbox unlock operations with key reset.
 */

// =============================================================================
// TYPE DEFINITIONS
// =============================================================================

/**
 * @typedef {Object} PlayerKeyStatus
 * @property {string} preferred_name_id  - Player's canonical name
 * @property {number} red                - Count of Red keys
 * @property {number} blue               - Count of Blue keys
 * @property {number} green              - Count of Green keys
 * @property {number} yellow             - Count of Yellow keys
 * @property {number} purple             - Count of Purple keys
 * @property {number} rainbow            - Count of Rainbow keys
 * @property {number} totalKeys          - Sum of all keys
 * @property {string} colorsEarned       - e.g. "Red, Blue, Yellow"
 * @property {string} collectedAll5      - "Yes" or "No"
 * @property {boolean} eligibleForRainbowKey - Whether player can convert to rainbow
 * @property {boolean} eligibleForLockbox    - Computed eligibility for unlock
 * @property {string} eligibilityReason      - Human-readable reason
 */

// =============================================================================
// CONSTANTS
// =============================================================================

const LOCKBOX_SHEET_NAME_ = 'Key_Tracker';
const LOCKBOX_HEADERS_ = {
  NAME: 'preferred_name_id',
  RED: 'Red',
  BLUE: 'Blue',
  GREEN: 'Green',
  YELLOW: 'Yellow',
  PURPLE: 'Purple',
  RAINBOW: 'Rainbow',
  TOTAL: 'Total_Number_of_Keys',
  COLORS_EARNED: 'Colors_of_Keys_Earned',
  COLLECTED_ALL_5: 'Collected_All_5',
  ELIGIBLE_RAINBOW: 'Eligible_for_rainbow_key'
};

// =============================================================================
// PUBLIC FUNCTIONS
// =============================================================================

/**
 * Returns a sorted list of player names present in Key_Tracker.
 *
 * @return {string[]} Sorted player names (A-Z)
 * @throws {Error} If Key_Tracker sheet is missing
 */
function getKeyTrackerPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOCKBOX_SHEET_NAME_);

  if (!sheet) {
    throw new Error('[SHEET_MISSING] Key_Tracker sheet not found. Please run Build/Repair to create it.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return []; // Only header row or empty
  }

  const headers = data[0];
  const nameCol = findColumnIndex_(headers, LOCKBOX_HEADERS_.NAME);

  if (nameCol === -1) {
    throw new Error('[SCHEMA_INVALID] Key_Tracker is missing the preferred_name_id column.');
  }

  // Collect non-empty player names
  const names = [];
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][nameCol] || '').trim();
    if (name) {
      names.push(name);
    }
  }

  // Sort alphabetically (A-Z)
  return names.sort((a, b) => a.localeCompare(b));
}

/**
 * Returns the key status for a single player.
 *
 * @param {string} name - preferred_name_id
 * @return {PlayerKeyStatus} Full key status including computed eligibility
 * @throws {Error} If player not found or sheet missing
 */
function getPlayerKeyStatus(name) {
  if (!name || typeof name !== 'string') {
    throw new Error('[INVALID_INPUT] Player name is required.');
  }

  const trimmedName = name.trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOCKBOX_SHEET_NAME_);

  if (!sheet) {
    throw new Error('[SHEET_MISSING] Key_Tracker sheet not found.');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Get column indices
  const cols = getColumnIndices_(headers);

  // Find player row
  let playerRow = null;
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][cols.name] || '').trim();
    if (rowName === trimmedName) {
      playerRow = data[i];
      break;
    }
  }

  if (!playerRow) {
    throw new Error('[PLAYER_NOT_FOUND] Player "' + trimmedName + '" not found in Key_Tracker.');
  }

  // Extract key counts with NaN coercion
  const red = coerceToNumber_(playerRow[cols.red]);
  const blue = coerceToNumber_(playerRow[cols.blue]);
  const green = coerceToNumber_(playerRow[cols.green]);
  const yellow = coerceToNumber_(playerRow[cols.yellow]);
  const purple = coerceToNumber_(playerRow[cols.purple]);
  const rainbow = coerceToNumber_(playerRow[cols.rainbow]);

  // Calculate total keys
  const totalKeys = red + blue + green + yellow + purple + rainbow;

  // Get colors earned (from column or compute)
  let colorsEarned = '';
  if (cols.colorsEarned !== -1) {
    colorsEarned = String(playerRow[cols.colorsEarned] || '').trim();
  }
  if (!colorsEarned) {
    colorsEarned = buildColorsEarnedString_(red, blue, green, yellow, purple);
  }

  // Get Collected_All_5
  let collectedAll5 = 'No';
  if (cols.collectedAll5 !== -1) {
    const val = String(playerRow[cols.collectedAll5] || '').trim().toLowerCase();
    collectedAll5 = (val === 'yes' || val === 'true') ? 'Yes' : 'No';
  }

  // Get Eligible_for_rainbow_key
  let eligibleForRainbowKey = false;
  if (cols.eligibleRainbow !== -1) {
    eligibleForRainbowKey = coerceToBoolean_(playerRow[cols.eligibleRainbow]);
  }

  // Compute lockbox eligibility
  const eligibility = computeLockboxEligibility_(
    red, blue, green, yellow, purple, rainbow, totalKeys, collectedAll5
  );

  return {
    preferred_name_id: trimmedName,
    red: red,
    blue: blue,
    green: green,
    yellow: yellow,
    purple: purple,
    rainbow: rainbow,
    totalKeys: totalKeys,
    colorsEarned: colorsEarned,
    collectedAll5: collectedAll5,
    eligibleForRainbowKey: eligibleForRainbowKey,
    eligibleForLockbox: eligibility.eligible,
    eligibilityReason: eligibility.reason
  };
}

/**
 * Lightweight wrapper that returns only eligibility info for a player.
 *
 * @param {string} name - preferred_name_id
 * @return {Object} Result object
 *   {
 *     success: boolean,
 *     eligible: boolean,
 *     reason: string,
 *     status?: PlayerKeyStatus
 *   }
 */
function verifyLockboxEligibility(name) {
  try {
    const status = getPlayerKeyStatus(name);
    return {
      success: true,
      eligible: status.eligibleForLockbox,
      reason: status.eligibilityReason,
      status: status
    };
  } catch (e) {
    return {
      success: false,
      eligible: false,
      reason: e.message || String(e)
    };
  }
}

/**
 * Unlocks the Lockbox for a player: logs the unlock and resets their keys.
 *
 * @param {Object} payload
 *   {
 *     preferred_name_id: string,
 *     staff?: string,
 *     prizeNote?: string
 *   }
 * @return {Object} Result object
 *   {
 *     success: boolean,
 *     message: string,
 *     status?: PlayerKeyStatus (post-reset status)
 *   }
 */
function unlockLockbox(payload) {
  // Validate payload
  if (!payload || !payload.preferred_name_id) {
    return {
      success: false,
      message: 'Player name (preferred_name_id) is required.'
    };
  }

  const name = String(payload.preferred_name_id).trim();
  const staff = payload.staff || '';
  const prizeNote = payload.prizeNote || '';

  try {
    // Fetch current status
    const previousStatus = getPlayerKeyStatus(name);

    // Check eligibility
    if (!previousStatus.eligibleForLockbox) {
      return {
        success: false,
        message: 'Player is not eligible: ' + previousStatus.eligibilityReason
      };
    }

    // Perform the unlock: reset keys to zero
    const newStatus = resetPlayerKeys_(name);

    // Log the action if logIntegrityAction exists
    if (typeof logIntegrityAction === 'function') {
      try {
        logIntegrityAction('LOCKBOX_UNLOCK', {
          preferred_name_id: name,
          staff: staff || getCurrentUser_(),
          prizeNote: prizeNote,
          before: previousStatus,
          after: newStatus,
          status: 'SUCCESS'
        });
      } catch (logError) {
        // Don't fail the operation if logging fails
        console.error('Failed to log LOCKBOX_UNLOCK:', logError);
      }
    }

    return {
      success: true,
      message: 'Lockbox unlocked and keys reset for ' + name + '.',
      status: newStatus
    };

  } catch (e) {
    // Log failure if logging exists
    if (typeof logIntegrityAction === 'function') {
      try {
        logIntegrityAction('LOCKBOX_UNLOCK', {
          preferred_name_id: name,
          staff: staff || getCurrentUser_(),
          error: e.message,
          status: 'FAILURE'
        });
      } catch (logError) {
        console.error('Failed to log LOCKBOX_UNLOCK failure:', logError);
      }
    }

    return {
      success: false,
      message: 'Unlock failed: ' + (e.message || String(e))
    };
  }
}

// =============================================================================
// PRIVATE HELPER FUNCTIONS
// =============================================================================

/**
 * Finds the index of a header in the headers array.
 * @private
 */
function findColumnIndex_(headers, headerName) {
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) {
      return i;
    }
  }
  return -1;
}

/**
 * Gets all column indices for Key_Tracker headers.
 * @private
 */
function getColumnIndices_(headers) {
  return {
    name: findColumnIndex_(headers, LOCKBOX_HEADERS_.NAME),
    red: findColumnIndex_(headers, LOCKBOX_HEADERS_.RED),
    blue: findColumnIndex_(headers, LOCKBOX_HEADERS_.BLUE),
    green: findColumnIndex_(headers, LOCKBOX_HEADERS_.GREEN),
    yellow: findColumnIndex_(headers, LOCKBOX_HEADERS_.YELLOW),
    purple: findColumnIndex_(headers, LOCKBOX_HEADERS_.PURPLE),
    rainbow: findColumnIndex_(headers, LOCKBOX_HEADERS_.RAINBOW),
    total: findColumnIndex_(headers, LOCKBOX_HEADERS_.TOTAL),
    colorsEarned: findColumnIndex_(headers, LOCKBOX_HEADERS_.COLORS_EARNED),
    collectedAll5: findColumnIndex_(headers, LOCKBOX_HEADERS_.COLLECTED_ALL_5),
    eligibleRainbow: findColumnIndex_(headers, LOCKBOX_HEADERS_.ELIGIBLE_RAINBOW)
  };
}

/**
 * Coerces a value to a number, defaulting to 0 for NaN.
 * @private
 */
function coerceToNumber_(value) {
  const num = Number(value);
  return isNaN(num) ? 0 : num;
}

/**
 * Coerces a value to a boolean.
 * @private
 */
function coerceToBoolean_(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const lower = value.toLowerCase().trim();
    return lower === 'true' || lower === 'yes' || lower === '1';
  }
  return Boolean(value);
}

/**
 * Builds a comma-separated string of colors earned from key counts.
 * @private
 */
function buildColorsEarnedString_(red, blue, green, yellow, purple) {
  const colors = [];
  if (red > 0) colors.push('Red');
  if (blue > 0) colors.push('Blue');
  if (green > 0) colors.push('Green');
  if (yellow > 0) colors.push('Yellow');
  if (purple > 0) colors.push('Purple');
  return colors.length > 0 ? colors.join(', ') : 'None';
}

/**
 * Computes lockbox eligibility based on key counts and collected status.
 * @private
 * @return {Object} { eligible: boolean, reason: string }
 */
function computeLockboxEligibility_(red, blue, green, yellow, purple, rainbow, totalKeys, collectedAll5) {
  // Count distinct colored keys with value > 0
  let distinctColors = 0;
  if (red > 0) distinctColors++;
  if (blue > 0) distinctColors++;
  if (green > 0) distinctColors++;
  if (yellow > 0) distinctColors++;
  if (purple > 0) distinctColors++;

  // Eligibility rules
  if (collectedAll5 === 'Yes') {
    return {
      eligible: true,
      reason: 'Player has all five colored keys.'
    };
  }

  if (rainbow > 0 && (distinctColors + rainbow) >= 5) {
    return {
      eligible: true,
      reason: 'Rainbow key(s) plus colored keys cover all five colors.'
    };
  }

  if (totalKeys === 0) {
    return {
      eligible: false,
      reason: 'Player has no keys yet.'
    };
  }

  return {
    eligible: false,
    reason: 'Player does not yet have a complete 5-color set.'
  };
}

/**
 * Resets all key counts to zero for a player and returns the new status.
 * @private
 */
function resetPlayerKeys_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOCKBOX_SHEET_NAME_);

  if (!sheet) {
    throw new Error('[SHEET_MISSING] Key_Tracker sheet not found.');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const cols = getColumnIndices_(headers);

  // Find player row index (1-based for sheet, 0-based for array)
  let playerRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][cols.name] || '').trim();
    if (rowName === name) {
      playerRowIndex = i;
      break;
    }
  }

  if (playerRowIndex === -1) {
    throw new Error('[PLAYER_NOT_FOUND] Player "' + name + '" not found in Key_Tracker.');
  }

  // Build the reset values
  // We'll update specific columns to reset values
  const updates = [];

  // Red, Blue, Green, Yellow, Purple, Rainbow -> 0
  if (cols.red !== -1) updates.push({ col: cols.red, value: 0 });
  if (cols.blue !== -1) updates.push({ col: cols.blue, value: 0 });
  if (cols.green !== -1) updates.push({ col: cols.green, value: 0 });
  if (cols.yellow !== -1) updates.push({ col: cols.yellow, value: 0 });
  if (cols.purple !== -1) updates.push({ col: cols.purple, value: 0 });
  if (cols.rainbow !== -1) updates.push({ col: cols.rainbow, value: 0 });

  // Total_Number_of_Keys -> 0 (only if column exists and is not a formula)
  if (cols.total !== -1) updates.push({ col: cols.total, value: 0 });

  // Colors_of_Keys_Earned -> "None" or ""
  if (cols.colorsEarned !== -1) updates.push({ col: cols.colorsEarned, value: '' });

  // Collected_All_5 -> "No"
  if (cols.collectedAll5 !== -1) updates.push({ col: cols.collectedAll5, value: 'No' });

  // Eligible_for_rainbow_key -> FALSE
  if (cols.eligibleRainbow !== -1) updates.push({ col: cols.eligibleRainbow, value: false });

  // Apply updates in batch (each cell individually for safety with mixed formulas)
  const sheetRow = playerRowIndex + 1; // Convert to 1-based
  updates.forEach(function(update) {
    sheet.getRange(sheetRow, update.col + 1).setValue(update.value);
  });

  // Return the new status after reset
  return getPlayerKeyStatus(name);
}

/**
 * Gets the current user's email or a default value.
 * @private
 */
function getCurrentUser_() {
  try {
    return Session.getActiveUser().getEmail() || 'unknown';
  } catch (e) {
    return 'unknown';
  }
}