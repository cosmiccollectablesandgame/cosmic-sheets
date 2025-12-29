/**
 * BP Service - Bonus Points Redemption v7.9.7
 * @fileoverview Manages BP_Total: redemption for catalog or store credit memo
 *
 * NOTE: This service only reads/writes to:
 *   - Current_BP
 *   - BP_Historical
 *   - LastUpdated
 *
 * The mission columns (Attendance Mission Points, Flag Mission Points, Dice Roll Points)
 * are managed EXCLUSIVELY by bpTotalPipeline.gs via updateBPTotalFromSources().
 */

// ============================================================================
// BP OPERATIONS
// ============================================================================

/**
 * Redeems BP for a catalog prize
 * @param {string} preferredName - Player name
 * @param {number} amount - BP to spend
 * @param {string} itemCode - Catalog item code (optional, auto-select if not provided)
 * @return {Object} Result {spent, remaining, item}
 */
function redeemBP_Catalog(preferredName, amount, itemCode = null) {
  ensureBPTotalSchema();

  // Get player BP
  const currentBP = getPlayerBP(preferredName);

  if (currentBP < amount) {
    throwError('Insufficient BP', 'INSUFFICIENT_BP', `Player has ${currentBP} BP, tried to spend ${amount}`);
  }

  // Get catalog item
  const catalog = getCatalog();
  let item = null;

  if (itemCode) {
    item = catalog.find(i => i.Code === itemCode);
  } else {
    // Auto-select: find L1 item with COGS ≈ amount (scaled by BP:COGS ratio, e.g., 1 BP = $0.50)
    const targetCOGS = amount * 0.5; // Assume 1 BP = $0.50
    const l1Items = catalog.filter(i =>
      i.Level === 'L1' &&
      coerceBoolean(i.InStock) &&
      coerceNumber(i.Qty, 0) > 0
    );

    l1Items.sort((a, b) => Math.abs(coerceNumber(a.COGS, 0) - targetCOGS) - Math.abs(coerceNumber(b.COGS, 0) - targetCOGS));
    item = l1Items[0];
  }

  if (!item) {
    throwError('Item not found', 'ITEM_NOT_FOUND', 'No eligible catalog item');
  }

  if (!coerceBoolean(item.InStock) || coerceNumber(item.Qty, 0) <= 0) {
    throwError('Item out of stock', 'OUT_OF_STOCK');
  }

  // Deduct BP
  const newBP = currentBP - amount;
  setPlayerBP(preferredName, newBP);

  // Decrement catalog stock
  decrementCatalogStock_([{
    code: item.Code,
    qty: 1
  }]);

  // Write to Spent_Pool
  const batchId = newBatchId();
  writeSpentPool([{
    eventId: 'BP_REDEEM',
    itemCode: item.Code,
    itemName: item.Name,
    level: item.Level || 'L1',
    qty: 1,
    cogs: coerceNumber(item.COGS, 0),
    eventType: 'BP_REDEMPTION'
  }], batchId);

  logIntegrityAction('BP_REDEEM', {
    preferredName,
    details: `Redeemed ${amount} BP for ${item.Code} (${item.Name}). BP: ${currentBP} → ${newBP}`,
    status: 'SUCCESS'
  });

  return {
    spent: amount,
    remaining: newBP,
    item: {
      code: item.Code,
      name: item.Name,
      level: item.Level
    }
  };
}

/**
 * Redeems BP as store credit memo (no actual balance mutation)
 * @param {string} preferredName - Player name
 * @param {number} amount - BP to convert
 * @return {Object} Result {spent, remaining, creditValue}
 */
function redeemBP_StoreCreditMemo(preferredName, amount) {
  ensureBPTotalSchema();

  const currentBP = getPlayerBP(preferredName);

  if (currentBP < amount) {
    throwError('Insufficient BP', 'INSUFFICIENT_BP');
  }

  // Deduct BP
  const newBP = currentBP - amount;
  setPlayerBP(preferredName, newBP);

  // Compute credit value (1 BP = $0.50)
  const creditValue = amount * 0.5;

  logIntegrityAction('BP_REDEEM', {
    preferredName,
    details: `Redeemed ${amount} BP as store credit memo: ${formatCurrency(creditValue)}. BP: ${currentBP} → ${newBP}. Process in POS.`,
    status: 'SUCCESS'
  });

  return {
    spent: amount,
    remaining: newBP,
    creditValue,
    memo: `Process ${formatCurrency(creditValue)} store credit for ${preferredName} in POS (taxed at redemption)`
  };
}

/**
 * Gets player's BP balance
 * @param {string} preferredName - Player name
 * @return {number} BP balance (0 if not found)
 */
function getPlayerBP(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');

  if (nameCol === -1 || bpCol === -1) return 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      return coerceNumber(data[i][bpCol], 0);
    }
  }

  return 0;
}

/**
 * Sets player's BP balance
 * NOTE: This function only updates Current_BP, BP_Historical, and LastUpdated.
 * It does NOT touch the mission columns (Attendance Mission Points, etc.)
 *
 * @param {string} preferredName - Player name
 * @param {number} amount - New BP amount
 */
function setPlayerBP(preferredName, amount) {
  ensureBPTotalSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BP_Total');

  if (!sheet) {
    throwError('BP_Total not found', 'SHEET_MISSING');
  }

  // Clamp 0-cap (global cap)
  const throttle = getThrottleKV();
  const globalCap = coerceNumber(throttle.BP_Global_Cap, 100);
  const clamped = clamp(amount, 0, globalCap);

  // Handle overflow to Prestige_Overflow
  if (amount > globalCap) {
    const overflow = amount - globalCap;
    addPrestigeOverflow_(preferredName, overflow);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const bpCol = headers.indexOf('Current_BP');
  const historicalCol = headers.indexOf('BP_Historical');
  const updatedCol = headers.indexOf('LastUpdated');

  if (nameCol === -1 || bpCol === -1) {
    throwError('Invalid BP_Total schema', 'SCHEMA_INVALID');
  }

  // Find player row
  let playerRow = -1;
  let currentHistorical = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      if (historicalCol !== -1) {
        currentHistorical = coerceNumber(data[i][historicalCol], 0);
      }
      break;
    }
  }

  const now = dateISO();

  if (playerRow === -1) {
    // Create new row - build array matching header positions
    const newRow = new Array(headers.length).fill('');
    newRow[nameCol] = preferredName;
    newRow[bpCol] = clamped;
    if (historicalCol !== -1) newRow[historicalCol] = amount; // Uncapped value as historical
    if (updatedCol !== -1) newRow[updatedCol] = now;

    sheet.appendRow(newRow);
  } else {
    // Update existing - only touch Current_BP, BP_Historical, LastUpdated
    sheet.getRange(playerRow + 1, bpCol + 1).setValue(clamped);

    // Update BP_Historical as monotonic max (use uncapped amount)
    if (historicalCol !== -1) {
      const newHistorical = Math.max(currentHistorical, amount);
      sheet.getRange(playerRow + 1, historicalCol + 1).setValue(newHistorical);
    }

    if (updatedCol !== -1) {
      sheet.getRange(playerRow + 1, updatedCol + 1).setValue(now);
    }
  }
}

/**
 * Awards BP to player
 * @param {string} preferredName - Player name
 * @param {number} amount - BP to award
 * @return {Object} Result {before, after, awarded}
 */
function awardBP(preferredName, amount) {
  const current = getPlayerBP(preferredName);
  const newAmount = current + amount;

  setPlayerBP(preferredName, newAmount);

  logIntegrityAction('BP_AWARD', {
    preferredName,
    details: `Awarded ${amount} BP. Total: ${current} → ${newAmount}`,
    status: 'SUCCESS'
  });

  return {
    before: current,
    after: newAmount,
    awarded: amount
  };
}

// ============================================================================
// PRESTIGE OVERFLOW
// ============================================================================

/**
 * Adds prestige overflow (BP exceeding global cap)
 * @param {string} preferredName - Player name
 * @param {number} overflow - Overflow amount
 * @private
 */
function addPrestigeOverflow_(preferredName, overflow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prestige_Overflow');

  if (!sheet) {
    sheet = ss.insertSheet('Prestige_Overflow');
    sheet.appendRow(['PreferredName', 'Total_Overflow', 'Last_Updated', 'Prestige_Tier']);
    sheet.setFrozenRows(1);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = 0;
  const overflowCol = 1;
  const updatedCol = 2;

  let playerRow = -1;
  let currentOverflow = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      currentOverflow = coerceNumber(data[i][overflowCol], 0);
      break;
    }
  }

  const newOverflow = currentOverflow + overflow;

  if (playerRow === -1) {
    sheet.appendRow([preferredName, newOverflow, dateISO(), 'Bronze']);
  } else {
    sheet.getRange(playerRow + 1, overflowCol + 1).setValue(newOverflow);
    sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
  }

  logIntegrityAction('PRESTIGE_OVERFLOW', {
    preferredName,
    details: `Overflow: ${currentOverflow} → ${newOverflow} (+${overflow})`,
    status: 'SUCCESS'
  });
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures BP_Total has the minimal required schema for award/redeem operations.
 * This only ensures the columns needed for runtime BP operations:
 *   - PreferredName, Current_BP, BP_Historical, LastUpdated
 *
 * The full schema (including mission columns) is managed by
 * ensureBPTotalSchemaEnhanced_() in bpTotalPipeline.gs
 */
function ensureBPTotalSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');

  // Minimal headers for award/redeem operations
  const minimalHeaders = ['PreferredName', 'Current_BP', 'BP_Historical', 'LastUpdated'];

  if (!sheet) {
    // If sheet doesn't exist, use the full pipeline schema
    if (typeof ensureBPTotalSchemaEnhanced_ === 'function') {
      ensureBPTotalSchemaEnhanced_();
      return;
    }
    // Fallback: create with minimal headers
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(minimalHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    return;
  }

  if (sheet.getLastRow() === 0) {
    // Empty sheet - use full pipeline schema if available
    if (typeof ensureBPTotalSchemaEnhanced_ === 'function') {
      ensureBPTotalSchemaEnhanced_();
      return;
    }
    sheet.appendRow(minimalHeaders);
    sheet.setFrozenRows(1);
    return;
  }

  // Only add minimal headers if missing - don't remove or rename existing columns
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = minimalHeaders.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}