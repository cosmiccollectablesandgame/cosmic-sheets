/**
 * ============================================================================
 * Bonus Points Service v7.9 - Comprehensive BP Management
 * ============================================================================
 *
 * @fileoverview Centralizes all Bonus Points (BP) logic for Cosmic Tournament Engine v7.9
 *
 * SPEC: BonusPointsService.spec
 *
 * INPUTS (Fishbone Schema):
 *   - PreferredNames: Master player list (canonical names)
 *   - BP_Total: Current_BP (0-100 cap), Historical_BP
 *   - BP_Prestige: Prestige overflow when BP > 100
 *   - Redeemed_BP: BP spending ledger
 *   - Dice_Points: Operational UI for dice-roll BP additions
 *   - Flag_Missions: Mission-based BP awards
 *   - Attendance_Missions: Attendance/achievement BP tracking
 *   - Event tabs (MM-DD-YYYY): Ranked event results for Top-4 and Black Hole Survivor
 *
 * OUTPUTS:
 *   - BP_Total: Updated Current_BP and Historical_BP
 *   - BP_Prestige: Overflow tracking
 *   - Redeemed_BP: Redemption log
 *   - Integrity_Log: Audit trail
 *   - UndiscoveredNames: Unknown player names
 *
 * INVARIANTS:
 *   - 1 BP = $.20-$1 internal accounting value
 *   - Current_BP capped at 100 (configurable via Prize_Throttle)
 *   - Overflow goes to BP_Prestige
 *   - Engine NEVER does tax math
 *   - All name resolution via PreferredNames canonical list
 *
 * PUBLIC API:
 *   - awardBonusPoints(rawName, amount, source, meta)
 *   - redeemBonusPoints(rawName, amount, sink, meta)
 *   - recomputeRankAndSurvivorBonuses()
 *   - recomputeFlagMissionBonuses()
 *   - awardFromD20Roll(rawName, roll, sourceEventId)
 *   - awardFromHybridRoll(rawName, amount, sourceEventId)
 *   - recomputeAllBonusPoints()
 *   - runBonusPointsSelfTest()
 *
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

/**
 * Gets BP configuration from Prize_Throttle or defaults
 * @return {Object} BP config {globalCap, d20Resolver, hybridCap, eventPattern}
 * @private
 */
function getBPConfig_() {
  const throttle = getThrottleKV_();

  return {
    globalCap: coerceNumber(throttle.BP_Global_Cap, 100),
    d20Resolver: {
      20: 'Pack',
      '17-19': 'Key Roll',
      '14-16': '+1 BP',
      '1-13': 'None'
    },
    hybridCap: coerceNumber(throttle.Hybrid_Roll_Cap, 50),
    eventPattern: /^(\d{1,2}-\d{1,2}(-\d{4}|[A-Z]?-\d{4}))$/i
  };
}

/**
 * Gets throttle KV (shared with other services)
 * @return {Object} Key-value throttle settings
 * @private
 */
function getThrottleKV_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Throttle');

  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  return toKV(data);
}

// ============================================================================
// CORE NAME RESOLUTION
// ============================================================================

/**
 * Resolves raw player name to canonical PreferredName
 * @param {string} rawName - Raw player name from any source
 * @return {string|null} Canonical PreferredName or null if not found
 */
function resolvePlayerName(rawName) {
  if (!rawName || typeof rawName !== 'string') {
    return null;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');

  if (!sheet) {
    logUndiscoveredName_(rawName, 'PreferredNames sheet missing');
    return null;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return null;

  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');

  if (nameCol === -1) {
    logUndiscoveredName_(rawName, 'Invalid PreferredNames schema');
    return null;
  }

  // Normalize search term
  const normalized = rawName.trim().toLowerCase();

  // Try exact match (case-insensitive)
  for (let i = 1; i < data.length; i++) {
    const preferredName = data[i][nameCol];
    if (preferredName && preferredName.toString().trim().toLowerCase() === normalized) {
      return preferredName.toString().trim();
    }
  }

  // Not found - log to UndiscoveredNames
  logUndiscoveredName_(rawName, 'Not in PreferredNames');
  return null;
}

/**
 * Logs unknown player name to UndiscoveredNames sheet
 * @param {string} rawName - Unknown name
 * @param {string} reason - Reason for non-resolution
 * @private
 */
function logUndiscoveredName_(rawName, reason) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('UndiscoveredNames');

    if (!sheet) {
      sheet = ss.insertSheet('UndiscoveredNames');
      sheet.appendRow(['Potential_Name', 'First_Seen_Sheet', 'Reason', 'Reviewed', 'Timestamp']);
      sheet.setFrozenRows(1);
    }

    // Check if already logged
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rawName) {
        return; // Already logged
      }
    }

    sheet.appendRow([
      rawName,
      'BonusPointsService',
      reason,
      false,
      dateISO()
    ]);
  } catch (e) {
    console.error('Failed to log undiscovered name:', e);
  }
}

// ============================================================================
// BP BALANCE OPERATIONS
// ============================================================================

/**
 * Gets player's current BP balance
 * @param {string} preferredName - Canonical player name
 * @return {Object} {currentBP, historicalBP, prestige}
 */
function getPlayerBPBalance(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    return { currentBP: 0, historicalBP: 0, prestige: 0 };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const currentCol =
  headers.indexOf('BP_Current') >= 0
    ? headers.indexOf('BP_Current')
    : headers.indexOf('Current_BP');

const historicalCol =
  headers.indexOf('BP_Historical') >= 0
    ? headers.indexOf('BP_Historical')
    : headers.indexOf('Historical_BP');

  if (nameCol === -1 || currentCol === -1) {
    return { currentBP: 0, historicalBP: 0, prestige: 0 };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const currentBP = coerceNumber(data[i][currentCol], 0);
      const historicalBP = historicalCol >= 0 ? coerceNumber(data[i][historicalCol], 0) : 0;
      const prestige = getPlayerPrestige_(preferredName);

      return { currentBP, historicalBP, prestige };
    }
  }

  return { currentBP: 0, historicalBP: 0, prestige: 0 };
}

/**
 * Gets player's prestige overflow points
 * @param {string} preferredName - Canonical player name
 * @return {number} Prestige points
 * @private
 */
function getPlayerPrestige_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Prestige');

  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return 0;

  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName') >= 0 ? headers.indexOf('PreferredName') : 0;
  const prestigeCol = headers.indexOf('Prestige_Points') >= 0 ? headers.indexOf('Prestige_Points') : 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      return coerceNumber(data[i][prestigeCol], 0);
    }
  }

  return 0;
}

/**
 * Sets player BP with cap enforcement and prestige overflow
 * @param {string} preferredName - Canonical player name
 * @param {number} newAmount - New BP amount (before clamping)
 * @param {number} deltaHistorical - Amount to add to Historical_BP (for awards, not redemptions)
 * @return {Object} {currentBP, overflow, prestige}
 * @private
 */
function setPlayerBP_(preferredName, newAmount, deltaHistorical = 0) {
  ensureBPTotalSchema_();
  ensureBPPrestigeSchema_();

  const config = getBPConfig_();
  const globalCap = config.globalCap;

  // Clamp to 0-globalCap
  const clamped = clamp(newAmount, 0, globalCap);
  const overflow = Math.max(0, newAmount - globalCap);

  // Update BP_Total
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    throwError('BP_Total not found', 'SHEET_MISSING');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
 const currentCol =
  headers.indexOf('BP_Current') >= 0
    ? headers.indexOf('BP_Current')
    : headers.indexOf('Current_BP');

const historicalCol =
  headers.indexOf('BP_Historical') >= 0
    ? headers.indexOf('BP_Historical')
    : headers.indexOf('Historical_BP');

  const updatedCol = headers.indexOf('LastUpdated');

  if (nameCol === -1 || currentCol === -1) {
    throwError('Invalid BP_Total schema', 'SCHEMA_INVALID');
  }

  // Find or create player row
  let playerRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      break;
    }
  }

  if (playerRow === -1) {
    // Create new row
    const newRow = [];
    newRow[nameCol] = preferredName;
    newRow[currentCol] = clamped;
    if (historicalCol >= 0) newRow[historicalCol] = deltaHistorical;
    if (updatedCol >= 0) newRow[updatedCol] = dateISO();

    // Pad with empty values to match column count
    while (newRow.length < headers.length) {
      if (newRow[newRow.length] === undefined) newRow.push('');
    }

    sheet.appendRow(newRow);
  } else {
    // Update existing
    sheet.getRange(playerRow + 1, currentCol + 1).setValue(clamped);

    if (historicalCol >= 0 && deltaHistorical !== 0) {
      const oldHistorical = coerceNumber(data[playerRow][historicalCol], 0);
      sheet.getRange(playerRow + 1, historicalCol + 1).setValue(oldHistorical + deltaHistorical);
    }

    if (updatedCol >= 0) {
      sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
    }
  }

  // Handle overflow to Prestige
  let totalPrestige = getPlayerPrestige_(preferredName);

  if (overflow > 0) {
    totalPrestige = addPrestigeOverflow_(preferredName, overflow);
  }

  return {
    currentBP: clamped,
    overflow: overflow,
    prestige: totalPrestige
  };
}

/**
 * Adds prestige overflow points
 * @param {string} preferredName - Canonical player name
 * @param {number} overflow - Overflow amount
 * @return {number} New total prestige
 * @private
 */
function addPrestigeOverflow_(preferredName, overflow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Prestige');

  if (!sheet) {
    ensureBPPrestigeSchema_();
    sheet = ss.getSheetByName('BP_Prestige');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = 0;
  const prestigeCol = headers.indexOf('Prestige_Points') >= 0 ? headers.indexOf('Prestige_Points') : 1;
  const totalEarnedCol = headers.indexOf('Total_BP_Earned_Lifetime');
  const updatedCol = headers.indexOf('Last_Updated');

  let playerRow = -1;
  let currentPrestige = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      currentPrestige = coerceNumber(data[i][prestigeCol], 0);
      break;
    }
  }

  const newPrestige = currentPrestige + overflow;

  if (playerRow === -1) {
    // Create new row
    const newRow = [preferredName, newPrestige];
    if (totalEarnedCol >= 0) newRow.push(0);
    if (headers.indexOf('Total_BP_Redeemed_Lifetime') >= 0) newRow.push(0);
    if (headers.indexOf('Prestige_Milestone') >= 0) newRow.push('Bronze');
    if (updatedCol >= 0) newRow.push(dateISO());

    sheet.appendRow(newRow);
  } else {
    sheet.getRange(playerRow + 1, prestigeCol + 1).setValue(newPrestige);
    if (updatedCol >= 0) {
      sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
    }
  }

  logIntegrityAction('PRESTIGE_OVERFLOW', {
    preferredName,
    details: `Overflow: ${currentPrestige} → ${newPrestige} (+${overflow})`,
    status: 'SUCCESS'
  });

  return newPrestige;
}

// ============================================================================
// PUBLIC API: AWARD & REDEEM
// ============================================================================

/**
 * Awards BP to a player, respecting cap 0-100 and prestige overflow
 * @param {string} rawName - Raw player name
 * @param {number} amount - BP to award
 * @param {string} source - Source (e.g., 'TOP4', 'BLACK_HOLE', 'D20', 'FLAG_MISSION')
 * @param {Object} meta - Optional metadata {eventId, note, dfTags}
 * @return {Object} Result {success, player, currentBP, prestige, awarded, overflow, error}
 */
function awardBonusPoints(rawName, amount, source, meta = {}) {
  try {
    // Resolve name
    const preferredName = resolvePlayerName(rawName);

    if (!preferredName) {
      return {
        success: false,
        error: `Unknown player: ${rawName}. Added to UndiscoveredNames.`,
        player: rawName
      };
    }

    // Validate amount
    if (amount <= 0) {
      return {
        success: false,
        error: 'Award amount must be positive',
        player: preferredName
      };
    }

    // Get current balance
    const balance = getPlayerBPBalance(preferredName);
    const newTotal = balance.currentBP + amount;

    // Set new BP (with cap and overflow handling)
    const result = setPlayerBP_(preferredName, newTotal, amount);

    // Log to Integrity_Log
    logIntegrityAction('BP_AWARD', {
      preferredName,
      eventId: meta.eventId || '',
      dfTags: meta.dfTags || [],
      details: `Source: ${source} | Awarded: ${amount} BP | ${balance.currentBP} → ${result.currentBP} (Overflow: ${result.overflow})`,
      status: 'SUCCESS'
    });

    return {
      success: true,
      player: preferredName,
      currentBP: result.currentBP,
      prestige: result.prestige,
      awarded: amount,
      overflow: result.overflow
    };

  } catch (e) {
    console.error('awardBonusPoints error:', e);
    return {
      success: false,
      error: e.message,
      player: rawName
    };
  }
}

/**
 * Redeems BP from a player, adjusting BP_Total and Redeemed_BP
 * @param {string} rawName - Raw player name
 * @param {number} amount - BP to redeem
 * @param {string} sink - Sink (e.g., 'STORE_CREDIT', 'PRIZE_REDEMPTION')
 * @param {Object} meta - Optional metadata {eventId, note, dfTags, itemRedeemed}
 * @return {Object} Result {success, player, newBP, redeemed, error}
 */
function redeemBonusPoints(rawName, amount, sink, meta = {}) {
  try {
    // Resolve name
    const preferredName = resolvePlayerName(rawName);

    if (!preferredName) {
      return {
        success: false,
        error: `Unknown player: ${rawName}. Added to UndiscoveredNames.`,
        player: rawName
      };
    }

    // Validate amount
    if (amount <= 0) {
      return {
        success: false,
        error: 'Redemption amount must be positive',
        player: preferredName
      };
    }

    // Get current balance
    const balance = getPlayerBPBalance(preferredName);

    if (balance.currentBP < amount) {
      return {
        success: false,
        error: `Insufficient BP. Has ${balance.currentBP}, tried to redeem ${amount}`,
        player: preferredName
      };
    }

    // Deduct BP (no change to Historical_BP)
    const newTotal = balance.currentBP - amount;
    const result = setPlayerBP_(preferredName, newTotal, 0);

     // Update Redeemed_BP ledger
    ensureRedeemedBPSchema_();
    const redeemedLifetime = updateRedeemedBPLedger_(
      preferredName,
      amount,
      meta.itemRedeemed || sink || '',
      meta.note || '',
      result.currentBP,
      balance.historicalBP
    );

    // Sync BP_Total: BP_Current and BP_Redeemed
    syncBPTotalRedeemed_(preferredName, result.currentBP, redeemedLifetime);


    // Log to Integrity_Log
    logIntegrityAction('BP_REDEEM', {
      preferredName,
      eventId: meta.eventId || '',
      dfTags: meta.dfTags || [],
      details: `Sink: ${sink} | Redeemed: ${amount} BP | ${balance.currentBP} → ${result.currentBP}`,
      status: 'SUCCESS'
    });

    return {
      success: true,
      player: preferredName,
      newBP: result.currentBP,
      redeemed: amount
    };

  } catch (e) {
    console.error('redeemBonusPoints error:', e);
    return {
      success: false,
      error: e.message,
      player: rawName
    };
  }
}

/**
 * Updates Redeemed_BP ledger (one row per redemption)
 * - Total_Redeemed = lifetime total AFTER this redemption
 * - BP_Current / BP_Historical = snapshots AFTER redemption
 * @param {string} preferredName - Canonical player name
 * @param {number} amount - Amount redeemed this transaction
 * @param {string} itemRedeemed - Description of what they got
 * @param {string} notes - Extra notes from the UI
 * @param {number} bpCurrent - BP_Current after redemption
 * @param {number} bpHistorical - BP_Historical (lifetime earned)
 * @private
 */
function updateRedeemedBPLedger_(preferredName, amount, itemRedeemed, notes, bpCurrent, bpHistorical) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Redeemed_BP');

  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn() || 1;
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];

  const nameCol       = headers.indexOf('PreferredName');
  const totalCol      = headers.indexOf('Total_Redeemed');
  const itemCol       = headers.indexOf('Item_Redeemed');
  const notesCol      = headers.indexOf('Notes');
  const bpCurrentCol  = headers.indexOf('BP_Current');
  const bpHistCol     = headers.indexOf('BP_Historical');
  const updatedCol    = headers.indexOf('LastUpdated');

  // Compute existing lifetime redeemed total for this player
  let currentLifetimeTotal = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const val = coerceNumber(data[i][totalCol], 0);
      if (val > currentLifetimeTotal) {
        currentLifetimeTotal = val;
      }
    }
  }

  const newLifetimeTotal = currentLifetimeTotal + amount;

  // Build row matching header length
  const newRow = new Array(headers.length).fill('');

  if (nameCol      >= 0) newRow[nameCol]      = preferredName;
  if (totalCol     >= 0) newRow[totalCol]     = newLifetimeTotal;
  if (itemCol      >= 0) newRow[itemCol]      = itemRedeemed || '';
  if (notesCol     >= 0) newRow[notesCol]     = notes || '';
  if (bpCurrentCol >= 0) newRow[bpCurrentCol] = bpCurrent;
  if (bpHistCol    >= 0) newRow[bpHistCol]    = bpHistorical;
  if (updatedCol   >= 0) newRow[updatedCol]   = dateISO();

   sheet.appendRow(newRow);
  return newLifetimeTotal;
}

/**
 * Syncs BP_Total row for a player after redemption
 * Updates:
 *  - BP_Current
 *  - BP_Redeemed (lifetime)
 *  - LastUpdated
 * @private
 */
function syncBPTotalRedeemed_(preferredName, bpCurrent, bpRedeemedLifetime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');
  if (!sheet || sheet.getLastRow() === 0) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = headers.indexOf('PreferredName');
  const currentCol =
    headers.indexOf('BP_Current') >= 0
      ? headers.indexOf('BP_Current')
      : headers.indexOf('Current_BP');
  const redeemedCol = headers.indexOf('BP_Redeemed');
  const updatedCol = headers.indexOf('LastUpdated');

  if (nameCol === -1) return;

  // Find player row
  let playerRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      break;
    }
  }

  if (playerRow === -1) {
    // Create new row with minimal fields
    const newRow = new Array(headers.length).fill('');
    if (nameCol      >= 0) newRow[nameCol]      = preferredName;
    if (currentCol   >= 0) newRow[currentCol]   = bpCurrent;
    if (redeemedCol  >= 0) newRow[redeemedCol]  = bpRedeemedLifetime;
    if (updatedCol   >= 0) newRow[updatedCol]   = dateISO();
    sheet.appendRow(newRow);
  } else {
    // Update existing row
    if (currentCol  >= 0) sheet.getRange(playerRow + 1, currentCol  + 1).setValue(bpCurrent);
    if (redeemedCol >= 0) sheet.getRange(playerRow + 1, redeemedCol + 1).setValue(bpRedeemedLifetime);
    if (updatedCol  >= 0) sheet.getRange(playerRow + 1, updatedCol  + 1).setValue(dateISO());
  }
}

// ============================================================================
// UI ENTRYPOINTS FOR ui/redeem_bp
// ============================================================================

/**
 * Opens the Redeem Bonus Points dialog
 */
function openRedeemBPDialog() {
  const html = HtmlService.createTemplateFromFile('ui/redeem_bp')
    .evaluate()
    .setWidth(480)
    .setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, 'Redeem Bonus Points');
}

/**
 * Returns list of PreferredNames for UI datalist
 * @return {string[]} sorted list of names
 */
function getPreferredNamesForUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const names = values
    .map(r => r[0])
    .filter(n => n && n.toString().trim() !== '')
    .map(n => n.toString().trim());

  names.sort();
  return names;
}

/**
 * Resolves name and returns BP balance for UI
 * @param {string} rawName
 */
function getBPInfoForUI(rawName) {
  const preferredName = resolvePlayerName(rawName);
  if (!preferredName) {
    return { success: false, error: 'Player not found in PreferredNames.' };
  }

  const balance = getPlayerBPBalance(preferredName);
  return {
    success: true,
    preferredName,
    currentBP: balance.currentBP,
    historicalBP: balance.historicalBP,
    prestige: balance.prestige
  };
}

/**
 * Handles redemption coming from the HTML UI
 * @param {Object} payload
 *  - rawName
 *  - amount
 *  - itemRedeemed
 *  - notes
 */
function handleRedeemBPFromUI(payload) {
  const rawName = (payload.rawName || '').trim();
  const amount = coerceNumber(payload.amount, 0);
  const itemRedeemed = (payload.itemRedeemed || '').trim();
  const notes = (payload.notes || '').trim();

  if (!rawName) {
    return { success: false, error: 'Player name is required.' };
  }
  if (amount <= 0) {
    return { success: false, error: 'Redemption amount must be positive.' };
  }

  const result = redeemBonusPoints(rawName, amount, 'PRIZE_REDEMPTION', {
    itemRedeemed,
    note: notes,
    dfTags: ['DF-BP-UI-REDEEM']
  });

  return result;
}

// ============================================================================
// EVENT-BASED BP: TOP-4 & BLACK HOLE SURVIVOR
// ============================================================================

/**
 * Recomputes Top-4 and Black Hole Survivor BP from all event sheets
 * Idempotent: safe to run repeatedly
 * @return {Object} Summary {top4Totals, blackHoleTotals, playersAffected}
 */
function recomputeRankAndSurvivorBonuses() {
  const config = getBPConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();

  const top4Counts = {};
  const blackHoleCounts = {};

  // Scan all event tabs
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Match event pattern: MM-DD-YYYY or MM-DD[suffix]-YYYY
    if (!config.eventPattern.test(sheetName)) {
      return; // Skip non-event sheets
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return; // No data

    const headers = data[0];
    const rankCol = headers.indexOf('Rank');
    const nameCol = headers.indexOf('PreferredName');

    if (rankCol === -1 || nameCol === -1) return; // Invalid schema

    // Scan rows for Top-4 and Black Hole Survivor
    let lastPlayerName = null;

    for (let i = 1; i < data.length; i++) {
      const rank = coerceNumber(data[i][rankCol], 0);
      const rawName = data[i][nameCol];

      if (!rawName) continue;

      const preferredName = resolvePlayerName(rawName);
      if (!preferredName) continue;

      lastPlayerName = preferredName;

      // Top-4 logic
      if (rank >= 1 && rank <= 4) {
        top4Counts[preferredName] = (top4Counts[preferredName] || 0) + 1;
      }
    }

    // Black Hole Survivor: last player in column B
    if (lastPlayerName) {
      blackHoleCounts[lastPlayerName] = (blackHoleCounts[lastPlayerName] || 0) + 1;
    }
  });

  // Write to Attendance_Missions
  writeTop4AndSurvivorToAttendanceMissions_(top4Counts, blackHoleCounts);

  return {
    top4Totals: top4Counts,
    blackHoleTotals: blackHoleCounts,
    playersAffected: unique([...Object.keys(top4Counts), ...Object.keys(blackHoleCounts)])
  };
}

/**
 * Writes Top-4 and Black Hole Survivor counts to Attendance_Missions
 * @param {Object} top4Counts - Map of player -> top-4 count
 * @param {Object} blackHoleCounts - Map of player -> black hole count
 * @private
 */
function writeTop4AndSurvivorToAttendanceMissions_(top4Counts, blackHoleCounts) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Missions');

  if (!sheet) {
    // Create sheet if missing
    sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow([
      'PreferredName',
      'Points_From_Dice_Rolls',
      'Bonus_Points_From_Top4',
      'Black_Hole_Survivor_Count',
      'C_Total',
      'Badges',
      'LastUpdated'
    ]);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const top4Col = headers.indexOf('Bonus_Points_From_Top4');
  const bhCol = headers.indexOf('Black_Hole_Survivor_Count');
  const updatedCol = headers.indexOf('LastUpdated');

  if (nameCol === -1) return; // Invalid schema

  // Build update map
  const allPlayers = unique([...Object.keys(top4Counts), ...Object.keys(blackHoleCounts)]);

  allPlayers.forEach(player => {
    const top4 = top4Counts[player] || 0;
    const bh = blackHoleCounts[player] || 0;

    // Find player row
    let playerRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === player) {
        playerRow = i;
        break;
      }
    }

    if (playerRow === -1) {
      // Create new row
      const newRow = [];
      newRow[nameCol] = player;
      if (top4Col >= 0) newRow[top4Col] = top4;
      if (bhCol >= 0) newRow[bhCol] = bh;
      if (updatedCol >= 0) newRow[updatedCol] = dateISO();

      // Pad with empty values
      while (newRow.length < headers.length) {
        if (newRow[newRow.length] === undefined) newRow.push('');
      }

      sheet.appendRow(newRow);
    } else {
      // Update existing
      if (top4Col >= 0) {
        sheet.getRange(playerRow + 1, top4Col + 1).setValue(top4);
      }
      if (bhCol >= 0) {
        sheet.getRange(playerRow + 1, bhCol + 1).setValue(bh);
      }
      if (updatedCol >= 0) {
        sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
      }
    }
  });

  logIntegrityAction('RECOMPUTE_TOP4_BH', {
    details: `Recomputed Top-4 and Black Hole Survivor for ${allPlayers.length} players`,
    status: 'SUCCESS'
  });
}

// ============================================================================
// FLAG MISSIONS
// ============================================================================

/**
 * Recomputes Flag Mission BP for all players
 * Reads Flag_Missions sheet and syncs to BP_Total
 * @return {Object} Summary {playersAffected, totalBPFromFlags}
 */
function recomputeFlagMissionBonuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Flag_Missions');

  if (!sheet) {
    return { playersAffected: [], totalBPFromFlags: 0 };
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { playersAffected: [], totalBPFromFlags: 0 };
  }

  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName') >= 0 ? headers.indexOf('PreferredName') : 0;

  // Try both possible header styles for points
  let pointsCol = headers.indexOf('Flag_Mission_Points');
  if (pointsCol === -1) {
    pointsCol = headers.indexOf('Flag Mission Points');
  }
  if (pointsCol === -1) {
    pointsCol = 1; // fallback
  }

  const playersAffected = [];
  let totalBP = 0;

  for (let i = 1; i < data.length; i++) {
    const rawName = data[i][nameCol];
    const flagPoints = coerceNumber(data[i][pointsCol], 0);

    if (!rawName || flagPoints === 0) continue;

    const preferredName = resolvePlayerName(rawName);
    if (!preferredName) continue;

    // NOTE: This is a recompute summary only right now.
    playersAffected.push(preferredName);
    totalBP += flagPoints;
  }

  logIntegrityAction('RECOMPUTE_FLAG_MISSIONS', {
    details: `Flag missions scanned for ${playersAffected.length} players, ${totalBP} total BP`,
    status: 'SUCCESS'
  });

  return { playersAffected, totalBPFromFlags: totalBP };
}

// ============================================================================
// DICE ROLLS: D20 & HYBRID
// ============================================================================

/**
 * Awards BP from a D20 roll based on resolver legend
 * @param {string} rawName - Raw player name
 * @param {number} roll - D20 roll result (1-20)
 * @param {string} sourceEventId - Event ID for logging
 * @return {Object} Result {success, player, awarded, resolution, error}
 */
function awardFromD20Roll(rawName, roll, sourceEventId) {
  const config = getBPConfig_();
  const resolver = config.d20Resolver; // kept for future flexibility

  // Determine BP award from roll
  let bpAward = 0;
  let resolution = 'None';

  if (roll === 20) {
    resolution = 'Pack';
  } else if (roll >= 17 && roll <= 19) {
    resolution = 'Key Roll';
  } else if (roll >= 14 && roll <= 16) {
    resolution = '+1 BP';
    bpAward = 1;
  } else {
    resolution = 'None';
  }

  if (bpAward === 0) {
    return {
      success: true,
      player: rawName,
      awarded: 0,
      resolution: resolution
    };
  }

  // Award BP
  const result = awardBonusPoints(rawName, bpAward, 'D20_ROLL', {
    eventId: sourceEventId,
    note: `D20 roll: ${roll}`,
    dfTags: ['DF-050']
  });

  return {
    ...result,
    resolution: resolution
  };
}

/**
 * Awards BP from a hybrid roll (judge-controlled amount)
 * @param {string} rawName - Raw player name
 * @param {number} amount - BP to award (subject to hybrid cap)
 * @param {string} sourceEventId - Event ID for logging
 * @return {Object} Result from awardBonusPoints
 */
function awardFromHybridRoll(rawName, amount, sourceEventId) {
  const config = getBPConfig_();
  const hybridCap = config.hybridCap;

  // Enforce hybrid cap
  const cappedAmount = Math.min(amount, hybridCap);

  return awardBonusPoints(rawName, cappedAmount, 'HYBRID_ROLL', {
    eventId: sourceEventId,
    note: `Hybrid roll: ${amount} BP (capped at ${hybridCap})`,
    dfTags: ['DF-050', 'DF-051']
  });
}

// ============================================================================
// FULL RECOMPUTE ORCHESTRATOR
// ============================================================================

/**
 * Recomputes all BP from scratch across all sources
 * WARNING: This is a full recompute and may take time
 * @return {Object} Summary {top4, blackHole, flagMissions, duration}
 */
function recomputeAllBonusPoints() {
  const startTime = new Date();

  SpreadsheetApp.getActiveSpreadsheet().toast('Recomputing all bonus points...', 'BP Service', 5);

  // Step 1: Recompute Top-4 and Black Hole Survivor
  const rankResults = recomputeRankAndSurvivorBonuses();

  // Step 2: Recompute Flag Missions
  const flagResults = recomputeFlagMissionBonuses();

  // Step 3: (Optional) Recompute attendance-based BP
  // Not implemented yet - would scan attendance and award BP per milestone

  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;

  logIntegrityAction('RECOMPUTE_ALL_BP', {
    details: `Full BP recompute completed in ${duration}s. Top-4: ${rankResults.playersAffected.length}, Flags: ${flagResults.playersAffected.length}`,
    status: 'SUCCESS'
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(`BP recompute complete (${duration}s)`, 'BP Service', 3);

  return {
    top4: rankResults,
    blackHole: rankResults.blackHoleTotals,
    flagMissions: flagResults,
    duration: duration
  };
}

// ============================================================================
// SELF-TEST & VALIDATION
// ============================================================================

/**
 * Runs self-test to validate BP service functionality
 * @return {Object} Test report {passed, failed, errors}
 */
function runBonusPointsSelfTest() {
  const report = {
    passed: [],
    failed: [],
    errors: []
  };

  try {
    // Test 1: Schema validation
    ensureBPTotalSchema_();
    ensureBPPrestigeSchema_();
    ensureRedeemedBPSchema_();
    report.passed.push('Schema validation');
  } catch (e) {
    report.failed.push('Schema validation');
    report.errors.push(e.message);
  }

  try {
    // Test 2: Name resolution
    const testName = resolvePlayerName('NonExistentPlayer12345');
    if (testName === null) {
      report.passed.push('Name resolution (negative case)');
    } else {
      report.failed.push('Name resolution (negative case)');
    }
  } catch (e) {
    report.failed.push('Name resolution');
    report.errors.push(e.message);
  }

  try {
    // Test 3: Config loading
    const config = getBPConfig_();
    if (config.globalCap > 0 && config.d20Resolver) {
      report.passed.push('Config loading');
    } else {
      report.failed.push('Config loading');
    }
  } catch (e) {
    report.failed.push('Config loading');
    report.errors.push(e.message);
  }

  try {
    // Test 4: D20 resolver logic
    const result = awardFromD20Roll('TestPlayer', 15, 'TEST');
    if (result.resolution === '+1 BP') {
      report.passed.push('D20 resolver (15 = +1 BP)');
    } else {
      report.failed.push('D20 resolver');
    }
  } catch (e) {
    report.failed.push('D20 resolver');
    report.errors.push(e.message);
  }

  // Log results
  logIntegrityAction('BP_SELF_TEST', {
    details: `Passed: ${report.passed.length}, Failed: ${report.failed.length}`,
    status: report.failed.length === 0 ? 'SUCCESS' : 'FAILURE'
  });

  return report;
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures BP_Total has required schema
 * Current standard headers:
 * PreferredName | BP_Current | Attendance Mission Points | Flag Mission Points |
 * Dice Roll Points | LastUpdated | BP_Historical | BP_Redeemed
 * @private
 */
function ensureBPTotalSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  const requiredHeaders = [
    'PreferredName',
    'BP_Current',
    'Attendance Mission Points',
    'Flag Mission Points',
    'Dice Roll Points',
    'LastUpdated',
    'BP_Historical',
    'BP_Redeemed'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, requiredHeaders.length)
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

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}


/**
 * Ensures BP_Prestige has required schema
 * @private
 */
function ensureBPPrestigeSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Prestige');

  const requiredHeaders = [
    'PreferredName',
    'Prestige_Points',
    'Total_BP_Earned_Lifetime',
    'Total_BP_Redeemed_Lifetime',
    'Prestige_Milestone',
    'Last_Updated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('BP_Prestige');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
  }
}

/**
 * Ensures Redeemed_BP has required schema
 * Columns (per-row ledger):
 * PreferredName | Total_Redeemed | Item_Redeemed | Notes | BP_Current | BP_Historical | LastUpdated
 *
 * Total_Redeemed is the player's lifetime redeemed BP total at the time of this row.
 * BP_Current / BP_Historical are snapshots AFTER this redemption.
 * @private
 */
function ensureRedeemedBPSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Redeemed_BP');

  const requiredHeaders = [
    'PreferredName',
    'Total_Redeemed',
    'Item_Redeemed',
    'Notes',
    'BP_Current',
    'BP_Historical',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Redeemed_BP');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, requiredHeaders.length)
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

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Ensures PreferredNames has required schema
 * @private
 */
function ensurePreferredNamesSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PreferredNames');

  const requiredHeaders = ['PreferredName', 'Aliases', 'FirstSeen', 'LastUpdated'];

  if (!sheet) {
    sheet = ss.insertSheet('PreferredNames');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
  }
}

/**
 * Ensures Attendance_Missions has required schema
 * @private
 */
function ensureAttendanceMissionsSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance_Missions');

  const requiredHeaders = [
    'PreferredName',
    'Points_From_Dice_Rolls',
    'Bonus_Points_From_Top4',
    'Black_Hole_Survivor_Count',
    'C_Total',
    'Badges',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:G1')
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

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

/**
 * Ensures UndiscoveredNames has required schema
 * @private
 */
function ensureUndiscoveredNamesSchema_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('UndiscoveredNames');

  const requiredHeaders = [
    'Potential_Name',
    'First_Seen_Sheet',
    'Reason',
    'Reviewed',
    'Timestamp'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('UndiscoveredNames');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1')
      .setFontWeight('bold')
      .setBackground('#ff9900')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
  }
}
