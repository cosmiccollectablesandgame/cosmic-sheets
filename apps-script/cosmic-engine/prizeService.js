/**
 * Prize Service - Preview/Commit Prize Distribution
 * @fileoverview Deterministic prize allocation with Preview→Commit pattern
 */
// ============================================================================
// END PRIZES - PREVIEW
// ============================================================================
/**
 * Previews end prizes for an event
 * @param {string} eventId - Event tab name
 * @param {Object} throttle - Throttle parameters (optional, will fetch if not provided)
 * @param {string} seed - Seed (optional, will use event seed if not provided)
 * @return {Object} Preview object {allocations, spend, hash, rlBand}
 */
function previewEndPrizes(eventId, throttle = null, seed = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);
  if (!sheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND');
  }
  const eventProps = getEventProps(sheet);
  const data = sheet.getDataRange().getValues();
  const players = [];
  // Extract players (skip header)
  for (let i = 1; i < data.length; i++) {
    const preferredName = data[i][1]; // Column B
    if (preferredName) {
      players.push({ preferredName, rank: i });
    }
  }
  if (players.length === 0) {
    throwError('No players in roster', 'NO_PLAYERS');
  }
  // Get throttle params
  if (!throttle) {
    throttle = getThrottleKV();
  }
  // Get or use seed
  const useSeed = seed || eventProps.event_seed || generateSeed();
  // Compute budget
  const budgetInfo = derivedBudgetForEvent(eventProps, players.length);
  const budget = budgetInfo.budget;
  // Get catalog
  const catalog = getCatalog();
  const eligibleItems = catalog.filter(item =>
    coerceBoolean(item.Eligible_End) &&
    coerceBoolean(item.InStock) &&
    coerceNumber(item.Qty, 0) > 0 &&
    players.length >= coerceNumber(item.Player_Threshold, 0)
  );
  if (eligibleItems.length === 0) {
    throwError('No eligible prizes in catalog', 'NO_PRIZES');
  }
  // Allocate prizes deterministically
  const allocations = allocatePrizes_(players, eligibleItems, budget, throttle, useSeed);
  // Compute spend
  const spend = sumBy(allocations, a => coerceNumber(a.cogs, 0) * a.qty);
  // Compute hash
  const previewObj = {
    eventId,
    seed: useSeed,
    allocations: allocations.map(a => ({
      preferredName: a.preferredName,
      code: a.code,
      qty: a.qty
    }))
  };
  const hash = computeHash(previewObj);
  // RL band
  const rlBand = getRLBandInfo(spend, budget);
  return {
    eventId,
    seed: useSeed,
    allocations,
    spend,
    budget,
    hash,
    rlBand: rlBand.band,
    rlPercent: rlBand.percent,
    players: players.length
  };
}
/**
 * Allocates prizes deterministically
 * @param {Array<Object>} players - Player list
 * @param {Array<Object>} eligibleItems - Eligible catalog items
 * @param {number} budget - Budget in COGS
 * @param {Object} throttle - Throttle params
 * @param {string} seed - Seed
 * @return {Array<Object>} Allocations [{preferredName, code, name, level, qty, cogs}]
 * @private
 */
function allocatePrizes_(players, eligibleItems, budget, throttle, seed) {
  const rng = createSeededRandom(seed);
  const allocations = [];
  const efMin = parseFloat(throttle.EF_Clamp_Min || 0.80);
  const efMax = parseFloat(throttle.EF_Clamp_Max || 2.25);
  const consolationRatio = parseFloat(throttle.Consolation_L1_Ratio || 0.20);
  let remainingBudget = budget;
  const itemStock = new Map();
  // Initialize stock tracking
  eligibleItems.forEach(item => {
    itemStock.set(item.Code, coerceNumber(item.Qty, 0));
  });
  // Sort items by level for stratified allocation
  const itemsByLevel = groupBy(eligibleItems, item => item.Level || 'L0');
  // Allocate to each player
  players.forEach(player => {
    // Determine target level based on rank (simplified logic)
    let targetLevel = 'L1';
    if (player.rank <= 4) {
      targetLevel = 'L3'; // Top 4
    } else if (player.rank <= 8) {
      targetLevel = 'L2'; // Top 8
    }
    // Get available items at target level
    let levelItems = itemsByLevel.get(targetLevel) || [];
    levelItems = levelItems.filter(item => itemStock.get(item.Code) > 0);
    if (levelItems.length === 0) {
      // Fall back to consolation (L1 or L0)
      const useL1 = rng() < consolationRatio;
      targetLevel = useL1 ? 'L1' : 'L0';
      levelItems = (itemsByLevel.get(targetLevel) || []).filter(item => itemStock.get(item.Code) > 0);
    }
    if (levelItems.length === 0) return; // No items available
    // Select item using EF clamp (weighted by EV_Cost)
    const item = selectItemWithEF_(levelItems, efMin, efMax, rng);
    if (!item) return;
    const itemCOGS = coerceNumber(item.COGS, 0);
    // Check budget and stock
    if (itemCOGS <= remainingBudget && itemStock.get(item.Code) > 0) {
      allocations.push({
        preferredName: player.preferredName,
        code: item.Code,
        name: item.Name,
        level: item.Level || 'L0',
        qty: 1,
        cogs: itemCOGS
      });
      remainingBudget -= itemCOGS;
      itemStock.set(item.Code, itemStock.get(item.Code) - 1);
    }
  });
  return allocations;
}
/**
 * Selects item using EF clamp (weighted random)
 * @param {Array<Object>} items - Items to choose from
 * @param {number} efMin - EF min
 * @param {number} efMax - EF max
 * @param {Function} rng - RNG function
 * @return {Object|null} Selected item
 * @private
 */
function selectItemWithEF_(items, efMin, efMax, rng) {
  if (items.length === 0) return null;
  // Weight by EV_Cost (clamped)
  const weights = items.map(item => {
    const ev = coerceNumber(item.EV_Cost, 1);
    return clamp(ev, efMin, efMax);
  });
  const totalWeight = weights.reduce((sum, w) => sum + w, 0);
  let roll = rng() * totalWeight;
  for (let i = 0; i < items.length; i++) {
    roll -= weights[i];
    if (roll <= 0) {
      return items[i];
    }
  }
  return items[items.length - 1]; // Fallback
}
// ============================================================================
// END PRIZES - COMMIT
// ============================================================================
/**
 * Commits end prizes (with hash verification)
 * @param {string} eventId - Event tab name
 * @param {string} previewHash - Hash from preview
 * @return {Object} Commit result
 */
function commitEndPrizes(eventId, previewHash) {
  // Get preview artifact
  const artifact = getPreviewArtifact(eventId);
  if (!artifact) {
    throwError('No preview found', 'NO_PREVIEW', 'Generate a preview first');
  }
  if (artifact.previewHash !== previewHash) {
    throwError('Preview hash mismatch', 'HASH_MISMATCH', 'Preview has changed. Regenerate preview.');
  }
  // Regenerate preview to get allocations
  const preview = previewEndPrizes(eventId, null, artifact.seed);
  // Verify hash again
  if (preview.hash !== previewHash) {
    throwError('Preview hash mismatch on regeneration', 'HASH_MISMATCH');
  }
  // Check RL band
  if (preview.rlBand === 'RED') {
    throwError('Budget exceeded', 'BUDGET_RED', 'Reduce allocations or increase budget');
  }
  // Write to event sheet (column F: End_Prizes)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);
  if (!sheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND');
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const endCol = headers.indexOf('End_Prizes');
  if (nameCol === -1 || endCol === -1) {
    throwError('Invalid event schema', 'SCHEMA_INVALID');
  }
  // Group allocations by player
  const allocationMap = new Map();
  preview.allocations.forEach(alloc => {
    if (!allocationMap.has(alloc.preferredName)) {
      allocationMap.set(alloc.preferredName, []);
    }
    for (let i = 0; i < alloc.qty; i++) {
      allocationMap.get(alloc.preferredName).push(alloc.code);
    }
  });
  // Write to sheet
  for (let i = 1; i < data.length; i++) {
    const preferredName = data[i][nameCol];
    const codes = allocationMap.get(preferredName) || [];
    sheet.getRange(i + 1, endCol + 1).setValue(codes.join(', '));
  }
  // Decrement catalog stock
  decrementCatalogStock_(preview.allocations);
  // Write to Spent_Pool
  const batchId = newBatchId();
  const eventProps = getEventProps(sheet);
  const spentEntries = preview.allocations.map(alloc => ({
    eventId,
    itemCode: alloc.code,
    itemName: alloc.name,
    level: alloc.level,
    qty: alloc.qty,
    cogs: alloc.cogs,
    eventType: eventProps.event_type || 'CONSTRUCTED'
  }));
  writeSpentPool(spentEntries, batchId);
  // Log commit
  logCommit(eventId, artifact.seed, previewHash, preview.hash, preview.rlBand, preview.spend);
  // Delete artifact
  deletePreviewArtifact(artifact.artifactId);
  return {
    success: true,
    allocated: preview.allocations.length,
    spend: preview.spend,
    budget: preview.budget,
    batchId
  };
}
/**
 * Decrements catalog stock for allocations
 * @param {Array<Object>} allocations - Allocations
 * @private
 */
function decrementCatalogStock_(allocations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Catalog');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeCol = headers.indexOf('Code');
  const qtyCol = headers.indexOf('Qty');
  const inStockCol = headers.indexOf('InStock');
  if (codeCol === -1 || qtyCol === -1) return;
  // Group by code
  const qtyMap = new Map();
  allocations.forEach(alloc => {
    const current = qtyMap.get(alloc.code) || 0;
    qtyMap.set(alloc.code, current + alloc.qty);
  });
  // Update quantities
  for (let i = 1; i < data.length; i++) {
    const code = data[i][codeCol];
    const deduct = qtyMap.get(code);
    if (deduct) {
      const currentQty = coerceNumber(data[i][qtyCol], 0);
      const newQty = Math.max(0, currentQty - deduct);
      sheet.getRange(i + 1, qtyCol + 1).setValue(newQty);
      // Update InStock
      if (inStockCol !== -1) {
        sheet.getRange(i + 1, inStockCol + 1).setValue(newQty > 0);
      }
    }
  }
}
// ============================================================================
// COMMANDER ROUNDS
// ============================================================================
/**
 * Previews Commander round prizes
 * @param {string} eventId - Event ID
 * @param {number} roundId - Round number (1-3)
 * @param {Object} throttle - Throttle params
 * @param {string} seed - Seed
 * @return {Object} Preview
 */
function previewCommanderRound(eventId, roundId, throttle = null, seed = null) {
  // Fixed seats: R1=1st, R2=1st&4th, R3=1st
  const seats = (roundId === 2) ? [1, 4] : [1];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);
  if (!sheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND');
  }
  const data = sheet.getDataRange().getValues();
  const players = [];
  for (let i = 1; i < data.length; i++) {
    const preferredName = data[i][1];
    if (preferredName) {
      players.push({ preferredName, rank: data[i][0] || i });
    }
  }
  // Get throttle
  if (!throttle) {
    throttle = getThrottleKV();
  }
  const useSeed = seed || generateSeed();
  const catalog = getCatalog();
  const eligibleItems = catalog.filter(item =>
    coerceBoolean(item.Eligible_Rounds) &&
    coerceBoolean(item.InStock) &&
    coerceNumber(item.Qty, 0) > 0
  );
  const allocations = [];
  seats.forEach(seat => {
    const player = players.find(p => p.rank === seat) || players[seat - 1];
    if (!player) return;
    // Allocate simple L1 prize
    const l1Items = eligibleItems.filter(item => item.Level === 'L1' && coerceNumber(item.Qty, 0) > 0);
    if (l1Items.length > 0) {
      const item = l1Items[0]; // Simple: take first
      allocations.push({
        preferredName: player.preferredName,
        code: item.Code,
        name: item.Name,
        level: item.Level,
        qty: 1,
        cogs: coerceNumber(item.COGS, 0),
        round: roundId
      });
    }
  });
  const spend = sumBy(allocations, a => a.cogs * a.qty);
  const hash = computeHash({ eventId, roundId, seed: useSeed, allocations });
  return {
    eventId,
    roundId,
    seed: useSeed,
    allocations,
    spend,
    hash
  };
}
/**
 * Commits Commander round prizes
 * @param {string} eventId - Event ID
 * @param {number} roundId - Round number
 * @param {string} previewHash - Preview hash
 * @return {Object} Commit result
 */
function commitCommanderRound(eventId, roundId, previewHash) {
  // Similar to commitEndPrizes but writes to R1_Prize, R2_Prize, or R3_Prize column
  const colName = `R${roundId}_Prize`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);
  if (!sheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND');
  }
  // Regenerate preview
  const preview = previewCommanderRound(eventId, roundId);
  if (preview.hash !== previewHash) {
    throwError('Hash mismatch', 'HASH_MISMATCH');
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const roundCol = headers.indexOf(colName);
  if (nameCol === -1 || roundCol === -1) {
    throwError('Invalid schema', 'SCHEMA_INVALID');
  }
  // Write allocations
  preview.allocations.forEach(alloc => {
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === alloc.preferredName) {
        sheet.getRange(i + 1, roundCol + 1).setValue(alloc.code);
        break;
      }
    }
  });
  // Decrement stock
  decrementCatalogStock_(preview.allocations);
  // Spent_Pool
  const batchId = newBatchId();
  const eventProps = getEventProps(sheet);
  const spentEntries = preview.allocations.map(alloc => ({
    eventId,
    itemCode: alloc.code,
    itemName: alloc.name,
    level: alloc.level,
    qty: alloc.qty,
    cogs: alloc.cogs,
    eventType: eventProps.event_type || 'CONSTRUCTED'
  }));
  writeSpentPool(spentEntries, batchId);
  logIntegrityAction('ROUND_ALLOCATE', {
    eventId,
    details: `Round ${roundId}: ${preview.allocations.length} prizes`,
    status: 'SUCCESS'
  });
  return {
    success: true,
    allocated: preview.allocations.length,
    spend: preview.spend
  };
}

// ============================================================================
// COMMANDER PRIZE PREVIEW/COMMIT (Full Template-Based)
// ============================================================================

/**
 * Previews Commander prizes from Prize_Throttle template
 * @param {string} eventId - Event ID
 * @return {Object} Preview with all rounds and budget info
 */
function previewCommanderPrizesFromTemplate(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName(eventId);

  if (!eventSheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND', `No tab named "${eventId}"`);
  }

  // Get Prize_Throttle sheet
  const throttleSheet = ss.getSheetByName('Prize_Throttle');
  if (!throttleSheet) {
    throwError('Prize_Throttle not found', 'THROTTLE_MISSING', 'Run Ops → Build/Repair to create it');
  }

  // Read master controls (rows 3-16, columns A-B)
  const controlsData = throttleSheet.getRange('A3:B16').getValues();
  const controls = {};
  controlsData.forEach(([key, value]) => {
    if (key) controls[key] = value;
  });

  const rlPercent = parseFloat(controls.RL_Percentage || '0.95');
  const defaultEntryFee = parseFloat(controls.Default_Entry_Fee || '15');

  // Get event props
  const eventProps = getEventProps(eventSheet);
  const entryFee = eventProps.entry || defaultEntryFee;

  // Get players
  const eventData = eventSheet.getDataRange().getValues();
  const players = [];
  for (let i = 1; i < eventData.length; i++) {
    const rank = eventData[i][0];
    const preferredName = eventData[i][1];
    if (preferredName) {
      players.push({ rank: rank || i, preferredName });
    }
  }

  if (players.length === 0) {
    throwError('No players in roster', 'NO_PLAYERS');
  }

  // Compute RL budget
  const rlBudget = entryFee * players.length * rlPercent;

  // Read template (rows 21-29, columns A-N)
  const templateData = throttleSheet.getRange('A21:N29').getValues();
  const catalog = getCatalog();

  // Parse template and allocate prizes
  const allocations = [];
  let totalCOGS = 0;

  templateData.forEach(row => {
    const [round, seat, targetLevel, ...levelQtys] = row;
    if (!round || !seat) return;

    // Determine which players get prizes based on seat
    let targetPlayers = [];
    if (seat === '1st') {
      targetPlayers = [players[0]];
    } else if (seat === '2nd') {
      targetPlayers = [players[1]];
    } else if (seat === '3rd') {
      targetPlayers = [players[2]];
    } else if (seat === '4th') {
      targetPlayers = [players[3]];
    } else if (seat === '5-8th') {
      targetPlayers = players.slice(4, 8);
    } else if (seat.match(/\d+/)) {
      // Single rank like "1", "2", etc.
      const rankNum = parseInt(seat, 10);
      targetPlayers = [players[rankNum - 1]];
    }

    targetPlayers = targetPlayers.filter(p => p); // Remove undefined

    // For each target player, allocate prizes based on level quantities
    targetPlayers.forEach(player => {
      // Check each level column (L0-L10)
      levelQtys.forEach((qty, idx) => {
        const level = `L${idx}`;
        const qtyNum = parseInt(qty, 10) || 0;
        if (qtyNum === 0) return;

        // Find eligible items at this level
        const eligibleItems = catalog.filter(item =>
          item.Level === level &&
          coerceBoolean(item.InStock) &&
          coerceNumber(item.Qty, 0) > 0
        );

        if (eligibleItems.length === 0) return;

        // Simple selection: pick first available item
        const item = eligibleItems[0];
        const itemCOGS = coerceNumber(item.COGS, 0);

        allocations.push({
          round,
          seat,
          preferredName: player.preferredName,
          code: item.Code,
          name: item.Name,
          level: item.Level,
          qty: qtyNum,
          cogs: itemCOGS
        });

        totalCOGS += itemCOGS * qtyNum;
      });
    });
  });

  // Compute RL band
  const rlBandInfo = getRLBandInfo(totalCOGS, rlBudget);

  // Update validation panel in Prize_Throttle (rows 34-39, column B)
  throttleSheet.getRange('B34').setValue(formatCurrency(totalCOGS));
  throttleSheet.getRange('B35').setValue(formatCurrency(rlBudget));
  throttleSheet.getRange('B36').setValue(rlBandInfo.percentFormatted);
  throttleSheet.getRange('B37').setValue(rlBandInfo.band);
  throttleSheet.getRange('B38').setValue(eventId);
  throttleSheet.getRange('B39').setValue(dateISO());

  // Store preview artifact
  const seed = eventProps.event_seed || generateSeed();
  const previewHash = computeHash({ eventId, allocations, seed });
  storePreviewArtifact(eventId + '_COMMANDER', seed, previewHash, 24);

  logIntegrityAction('COMMANDER_PREVIEW', {
    eventId,
    seed,
    details: `Preview: ${allocations.length} prizes, COGS: ${formatCurrency(totalCOGS)}, RL: ${rlBandInfo.percentFormatted}`,
    rlBand: rlBandInfo.band,
    status: 'SUCCESS'
  });

  return {
    eventId,
    seed,
    allocations,
    expectedCOGS: totalCOGS,
    rlBudget,
    rlBand: rlBandInfo.band,
    rlPercent: rlBandInfo.percent,
    rlPercentFormatted: rlBandInfo.percentFormatted,
    players: players.length,
    hash: previewHash
  };
}

/**
 * Commits Commander prizes from template to event sheet
 * @param {string} eventId - Event ID
 * @param {string} previewHash - Hash from preview
 * @return {Object} Commit result
 */
function commitCommanderPrizesFromTemplate(eventId, previewHash) {
  // Get preview artifact
  const artifact = getPreviewArtifact(eventId + '_COMMANDER');

  if (!artifact) {
    throwError('No preview found', 'NO_PREVIEW', 'Run Preview first');
  }

  if (artifact.previewHash !== previewHash) {
    throwError('Preview hash mismatch', 'HASH_MISMATCH', 'Preview has changed. Regenerate preview.');
  }

  // Regenerate preview to get allocations
  const preview = previewCommanderPrizesFromTemplate(eventId);

  // Verify hash
  if (preview.hash !== previewHash) {
    throwError('Preview hash mismatch on regeneration', 'HASH_MISMATCH');
  }

  // Check RL band
  if (preview.rlBand === 'RED') {
    throwError('Budget exceeded', 'BUDGET_RED', 'Reduce allocations in Prize_Throttle template');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    throwError('Event not found', 'EVENT_NOT_FOUND');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const r1Col = headers.indexOf('R1_Prize');
  const r2Col = headers.indexOf('R2_Prize');
  const r3Col = headers.indexOf('R3_Prize');
  const endCol = headers.indexOf('End_Prizes');

  if (nameCol === -1) {
    throwError('Invalid event schema', 'SCHEMA_INVALID', 'PreferredName column not found');
  }

  // Group allocations by player and round
  const allocationsByPlayer = {};

  preview.allocations.forEach(alloc => {
    const key = alloc.preferredName;
    if (!allocationsByPlayer[key]) {
      allocationsByPlayer[key] = {
        R1: [],
        R2: [],
        R3: [],
        End: []
      };
    }

    const roundKey = alloc.round;
    if (allocationsByPlayer[key][roundKey]) {
      for (let i = 0; i < alloc.qty; i++) {
        allocationsByPlayer[key][roundKey].push(alloc.code);
      }
    }
  });

  // Write to sheet
  for (let i = 1; i < data.length; i++) {
    const preferredName = data[i][nameCol];
    const playerAllocs = allocationsByPlayer[preferredName];

    if (!playerAllocs) continue;

    if (r1Col !== -1 && playerAllocs.R1.length > 0) {
      sheet.getRange(i + 1, r1Col + 1).setValue(playerAllocs.R1.join(', '));
    }
    if (r2Col !== -1 && playerAllocs.R2.length > 0) {
      sheet.getRange(i + 1, r2Col + 1).setValue(playerAllocs.R2.join(', '));
    }
    if (r3Col !== -1 && playerAllocs.R3.length > 0) {
      sheet.getRange(i + 1, r3Col + 1).setValue(playerAllocs.R3.join(', '));
    }
    if (endCol !== -1 && playerAllocs.End.length > 0) {
      sheet.getRange(i + 1, endCol + 1).setValue(playerAllocs.End.join(', '));
    }
  }

  // Decrement catalog stock
  decrementCatalogStock_(preview.allocations);

  // Write to Spent_Pool
  const batchId = newBatchId();
  const eventProps = getEventProps(sheet);
  const spentEntries = preview.allocations.map(alloc => ({
    eventId,
    itemCode: alloc.code,
    itemName: alloc.name,
    level: alloc.level,
    qty: alloc.qty,
    cogs: alloc.cogs,
    eventType: eventProps.event_type || 'CONSTRUCTED'
  }));

  writeSpentPool(spentEntries, batchId);

  // Log commit
  logCommit(eventId + '_COMMANDER', preview.seed, previewHash, preview.hash, preview.rlBand, preview.expectedCOGS);

  // Delete artifact
  deletePreviewArtifact(artifact.artifactId);

  return {
    success: true,
    allocated: preview.allocations.length,
    spend: preview.expectedCOGS,
    budget: preview.rlBudget,
    rlBand: preview.rlBand,
    batchId
  };
}