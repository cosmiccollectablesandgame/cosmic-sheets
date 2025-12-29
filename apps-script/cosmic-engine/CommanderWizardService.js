/**
 * CommanderWizardService.gs
 *
 * Server-side services for the Commander Event Wizard modal.
 * Provides unified management of Commander round and end-of-event prizes.
 *
 * @version 7.9.7
 * @author Cosmic Games Tournament Manager
 */

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * @typedef {Object} CommanderRoundState
 * @property {boolean} r1Awarded
 * @property {boolean} r2Awarded
 * @property {boolean} r3Awarded
 */

/**
 * @typedef {Object} CommanderEndState
 * @property {boolean} endAwarded
 * @property {boolean} locked - true if event is locked/finalized
 */

/**
 * @typedef {Object} CommanderEventState
 * @property {string} eventId
 * @property {string} displayName
 * @property {string} date
 * @property {string} format
 * @property {number|null} playerCount
 * @property {CommanderRoundState} rounds
 * @property {CommanderEndState} end
 */

/**
 * @typedef {Object} PrizeItem
 * @property {string} player
 * @property {string} prizeCode
 * @property {string} prizeName
 * @property {string} rarity - e.g. "L2", "L1", "L0"
 * @property {number} qty
 */

/**
 * @typedef {Object} PrizePreview
 * @property {string} scope - "ROUND" or "END"
 * @property {number|null} round - 1, 2, 3 or null for end
 * @property {PrizeItem[]} items
 * @property {number|null} estimatedCost
 * @property {string|null} budgetLabel - e.g. "Commander Round 1 budget"
 * @property {number|null} budget - total budget available
 * @property {string|null} rlBand - GREEN, AMBER, or RED
 */

/**
 * @typedef {Object} CommanderEventSummary
 * @property {string} sheetName
 * @property {string} displayName
 * @property {string} date
 * @property {number|null} playerCount
 * @property {CommanderRoundState} rounds
 * @property {CommanderEndState} end
 */

// ============================================================================
// Configuration Constants
// ============================================================================

/** Flag names stored in event sheet config row */
const CMD_FLAGS = {
  ROUND1: 'Round1_Awarded',
  ROUND2: 'Round2_Awarded',
  ROUND3: 'Round3_Awarded',
  END: 'End_Awarded',
  LOCKED: 'Locked'
};

/** Commander round seats by round number */
const ROUND_SEATS = {
  1: [1],       // Round 1: 1st place only
  2: [1, 4],    // Round 2: 1st and 4th place
  3: [1]        // Round 3: 1st place only
};

/** L2 is the target level for Commander round prizes */
const ROUND_PRIZE_LEVEL = 'L2';

// ============================================================================
// getCommanderEvents
// ============================================================================

/**
 * Returns a list of Commander event summaries for the wizard.
 * Only returns events detected as Commander format.
 *
 * @return {CommanderEventSummary[]}
 */
function getCommanderEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const eventPattern = /^(\d{2})-(\d{2})x?-(\d{4})/;

  const events = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const match = sheetName.match(eventPattern);

    if (!match) return;

    // Check if this is a Commander event
    const format = detectEventFormat_(sheet);
    if (format !== 'Commander') return;

    // Parse date from sheet name
    const month = match[1];
    const day = match[2];
    const year = match[3];
    const dateStr = `${year}-${month}-${day}`;

    // Count players
    const playerCount = countPlayers_(sheet);

    // Get flags
    const flags = getCommanderFlags_(sheet);

    // Build display name
    const displayName = `${sheetName} Commander`;

    events.push({
      sheetName: sheetName,
      displayName: displayName,
      date: dateStr,
      playerCount: playerCount,
      rounds: {
        r1Awarded: flags.round1,
        r2Awarded: flags.round2,
        r3Awarded: flags.round3
      },
      end: {
        endAwarded: flags.end,
        locked: flags.locked
      }
    });
  });

  // Sort by date, most recent first
  events.sort((a, b) => b.date.localeCompare(a.date));

  return events;
}

// ============================================================================
// getCommanderEventState
// ============================================================================

/**
 * Returns detailed state for a specific Commander event.
 *
 * @param {string} eventId - sheet name
 * @return {CommanderEventState}
 */
function getCommanderEventState(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    throw new Error(`Event sheet "${eventId}" not found.`);
  }

  // Parse date from sheet name
  const match = eventId.match(/^(\d{2})-(\d{2})x?-(\d{4})/);
  const dateStr = match ? `${match[3]}-${match[1]}-${match[2]}` : '';

  // Get format
  const format = detectEventFormat_(sheet);

  // Count players
  const playerCount = countPlayers_(sheet);

  // Get flags
  const flags = getCommanderFlags_(sheet);

  // Build display name
  const displayName = format === 'Commander' ? `${eventId} Commander` : eventId;

  return {
    eventId: eventId,
    displayName: displayName,
    date: dateStr,
    format: format,
    playerCount: playerCount,
    rounds: {
      r1Awarded: flags.round1,
      r2Awarded: flags.round2,
      r3Awarded: flags.round3
    },
    end: {
      endAwarded: flags.end,
      locked: flags.locked
    }
  };
}

// ============================================================================
// previewCommanderRoundPrizes
// ============================================================================

/**
 * Returns a prize preview for a Commander round (R1/R2/R3).
 *
 * @param {string} eventId
 * @param {number} roundNumber - 1, 2, or 3
 * @return {PrizePreview}
 */
function previewCommanderRoundPrizes(eventId, roundNumber) {
  // Validate round number
  if (![1, 2, 3].includes(roundNumber)) {
    return {
      scope: 'ROUND',
      round: roundNumber,
      items: [],
      estimatedCost: null,
      budgetLabel: `Invalid Round ${roundNumber}`
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    return {
      scope: 'ROUND',
      round: roundNumber,
      items: [],
      estimatedCost: null,
      budgetLabel: 'Event not found'
    };
  }

  // Get seats for this round
  const seats = ROUND_SEATS[roundNumber];

  // Get ranked players
  const players = getRankedPlayers_(sheet);

  if (players.length === 0) {
    return {
      scope: 'ROUND',
      round: roundNumber,
      items: [],
      estimatedCost: 0,
      budgetLabel: `Commander Round ${roundNumber}`
    };
  }

  // Get L2 prizes from catalog
  const catalogItems = getCatalogItemsByLevel_(ROUND_PRIZE_LEVEL);

  if (catalogItems.length === 0) {
    return {
      scope: 'ROUND',
      round: roundNumber,
      items: [],
      estimatedCost: null,
      budgetLabel: 'No L2 prizes available'
    };
  }

  // Get event seed for deterministic selection
  const props = typeof getEventProps === 'function' ? getEventProps(sheet) : {};
  const seed = props.event_seed || generateWizardSeed_();
  const rng = createWizardRng_(seed + `R${roundNumber}`);

  // Build prize items
  const items = [];
  let estimatedCost = 0;

  seats.forEach(seatRank => {
    const player = players.find(p => p.rank === seatRank);
    if (!player) return;

    // Select a prize
    const prize = selectPrize_(catalogItems, rng);
    if (!prize) return;

    const cogs = parseFloat(prize.COGS) || 0;

    items.push({
      player: player.name,
      prizeCode: prize.Code || '',
      prizeName: prize.Name,
      rarity: prize.Level || ROUND_PRIZE_LEVEL,
      qty: 1
    });

    estimatedCost += cogs;
  });

  return {
    scope: 'ROUND',
    round: roundNumber,
    items: items,
    estimatedCost: estimatedCost,
    budgetLabel: `Commander Round ${roundNumber}`
  };
}

// ============================================================================
// commitCommanderRoundPrizes
// ============================================================================

/**
 * Commits Commander round prizes for a given event/round.
 *
 * @param {string} eventId
 * @param {number} roundNumber
 * @return {Object} { success: boolean, message: string, preview?: PrizePreview }
 */
function commitCommanderRoundPrizes(eventId, roundNumber) {
  // Validate round number
  if (![1, 2, 3].includes(roundNumber)) {
    return {
      success: false,
      message: `Invalid round number: ${roundNumber}. Must be 1, 2, or 3.`
    };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    return {
      success: false,
      message: `Event sheet "${eventId}" not found.`
    };
  }

  try {
    // Check current state
    const flags = getCommanderFlags_(sheet);
    const flagKey = `round${roundNumber}`;

    if (flags[flagKey]) {
      return {
        success: false,
        message: `Round ${roundNumber} prizes already awarded for this event.`
      };
    }

    // Generate preview
    const preview = previewCommanderRoundPrizes(eventId, roundNumber);

    if (!preview.items || preview.items.length === 0) {
      return {
        success: false,
        message: `No prizes to award for Round ${roundNumber}. Check player rankings and prize catalog.`
      };
    }

    // Get headers and find/create prize column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const prizeColName = `R${roundNumber}_Prize`;
    let prizeColIdx = headers.findIndex(h =>
      String(h).toLowerCase() === prizeColName.toLowerCase()
    );

    if (prizeColIdx < 0) {
      prizeColIdx = headers.length;
      sheet.getRange(1, prizeColIdx + 1).setValue(prizeColName);
    }

    // Get player name column
    const prefNameIdx = findPreferredNameColumn_(headers);

    if (prefNameIdx < 0) {
      return {
        success: false,
        message: 'Cannot find PreferredName column in event sheet.'
      };
    }

    // Build name to row map
    const lastRow = sheet.getLastRow();
    const nameData = sheet.getRange(2, prefNameIdx + 1, lastRow - 1, 1).getValues();
    const nameToRow = new Map();
    nameData.forEach((row, idx) => {
      const name = String(row[0]).trim();
      if (name) nameToRow.set(name, idx + 2);
    });

    // Write prizes to sheet
    let awarded = 0;
    preview.items.forEach(item => {
      const rowNum = nameToRow.get(item.player);
      if (rowNum) {
        sheet.getRange(rowNum, prizeColIdx + 1).setValue(item.prizeName);
        awarded++;
      }
    });

    // Update flags
    setCommanderFlag_(sheet, CMD_FLAGS[`ROUND${roundNumber}`], true);

    // Log to Integrity_Log
    logWizardAction_(`COMMANDER_ROUND_${roundNumber}`, eventId, {
      round: roundNumber,
      prizesAwarded: awarded,
      items: preview.items.map(i => `${i.player}: ${i.prizeName}`).join(', ')
    });

    return {
      success: true,
      message: `Commander Round ${roundNumber} prizes committed: ${awarded} player(s) awarded.`,
      preview: preview
    };
  } catch (e) {
    logWizardAction_(`COMMANDER_ROUND_${roundNumber}_ERROR`, eventId, {
      error: e.message
    });

    return {
      success: false,
      message: `Error committing Round ${roundNumber} prizes: ${e.message}`
    };
  }
}

// ============================================================================
// previewCommanderEndPrizes
// ============================================================================

/**
 * Returns a prize preview for Commander end-of-event prizes.
 *
 * @param {string} eventId
 * @return {PrizePreview}
 */
function previewCommanderEndPrizes(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    return {
      scope: 'END',
      round: null,
      items: [],
      estimatedCost: null,
      budgetLabel: 'Event not found'
    };
  }

  // Get event properties for budget calculation
  const props = typeof getEventProps === 'function' ? getEventProps(sheet) : {};
  const entry = parseFloat(props.entry) || 5.00;
  const kitCost = parseFloat(props.kit_cost_per_player) || 0;

  // Get ranked players
  const players = getRankedPlayers_(sheet);
  const playerCount = players.length;

  if (playerCount === 0) {
    return {
      scope: 'END',
      round: null,
      items: [],
      estimatedCost: 0,
      budgetLabel: 'Commander End Prizes',
      budget: 0,
      rlBand: 'GREEN'
    };
  }

  // Calculate budget (RL95 rule)
  const budget = (entry - kitCost) * playerCount * 0.95;

  // Get catalog items by level
  const catalogL3 = getCatalogItemsByLevel_('L3');
  const catalogL2 = getCatalogItemsByLevel_('L2');
  const catalogL1 = getCatalogItemsByLevel_('L1');
  const catalogL0 = getCatalogItemsByLevel_('L0');

  // Track stock
  const stock = new Map();
  [...catalogL3, ...catalogL2, ...catalogL1, ...catalogL0].forEach(item => {
    stock.set(item.Code || item.Name, item.Qty || 999);
  });

  // Get event seed
  const seed = props.event_seed || generateWizardSeed_();
  const rng = createWizardRng_(seed + 'END');

  // Allocate prizes
  const items = [];
  let estimatedCost = 0;

  players.forEach(player => {
    // Determine target level based on rank
    let targetCatalog;
    if (player.rank <= 4) {
      targetCatalog = catalogL3.length > 0 ? catalogL3 : catalogL2;
    } else if (player.rank <= 8) {
      targetCatalog = catalogL2.length > 0 ? catalogL2 : catalogL1;
    } else {
      targetCatalog = catalogL1.length > 0 ? catalogL1 : catalogL0;
    }

    // Filter by stock
    const available = targetCatalog.filter(item => {
      const key = item.Code || item.Name;
      return (stock.get(key) || 0) > 0;
    });

    if (available.length === 0) return;

    // Select prize
    const prize = selectPrize_(available, rng);
    if (!prize) return;

    const cogs = parseFloat(prize.COGS) || 0;

    // Check budget
    if (estimatedCost + cogs > budget) return;

    // Update stock
    const stockKey = prize.Code || prize.Name;
    stock.set(stockKey, (stock.get(stockKey) || 1) - 1);

    items.push({
      player: player.name,
      prizeCode: prize.Code || '',
      prizeName: prize.Name,
      rarity: prize.Level || 'L1',
      qty: 1
    });

    estimatedCost += cogs;
  });

  // Calculate RL band
  const rlPercent = budget > 0 ? (estimatedCost / budget) * 100 : 0;
  const rlBand = rlPercent <= 90 ? 'GREEN' : (rlPercent <= 95 ? 'AMBER' : 'RED');

  return {
    scope: 'END',
    round: null,
    items: items,
    estimatedCost: estimatedCost,
    budgetLabel: 'Commander End Prizes',
    budget: budget,
    rlBand: rlBand
  };
}

// ============================================================================
// commitCommanderEndPrizes
// ============================================================================

/**
 * Commits Commander end-of-event prizes.
 *
 * @param {string} eventId
 * @return {Object} { success: boolean, message: string, preview?: PrizePreview }
 */
function commitCommanderEndPrizes(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(eventId);

  if (!sheet) {
    return {
      success: false,
      message: `Event sheet "${eventId}" not found.`
    };
  }

  try {
    // Check current state
    const flags = getCommanderFlags_(sheet);

    if (flags.end) {
      return {
        success: false,
        message: 'End-of-event prizes already awarded for this event.'
      };
    }

    // Generate preview
    const preview = previewCommanderEndPrizes(eventId);

    if (!preview.items || preview.items.length === 0) {
      return {
        success: false,
        message: 'No prizes to award. Check player roster and prize catalog.'
      };
    }

    // Check RL band - block RED
    if (preview.rlBand === 'RED') {
      return {
        success: false,
        message: 'Cannot commit: Budget exceeded (RED band). Review prize allocation.'
      };
    }

    // Get headers and find/create End_Prizes column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let endPrizeIdx = headers.findIndex(h =>
      String(h).toLowerCase().includes('end_prize')
    );

    if (endPrizeIdx < 0) {
      endPrizeIdx = headers.length;
      sheet.getRange(1, endPrizeIdx + 1).setValue('End_Prizes');
    }

    // Get player name column
    const prefNameIdx = findPreferredNameColumn_(headers);

    if (prefNameIdx < 0) {
      return {
        success: false,
        message: 'Cannot find PreferredName column in event sheet.'
      };
    }

    // Build name to row map
    const lastRow = sheet.getLastRow();
    const nameData = sheet.getRange(2, prefNameIdx + 1, lastRow - 1, 1).getValues();
    const nameToRow = new Map();
    nameData.forEach((row, idx) => {
      const name = String(row[0]).trim();
      if (name) nameToRow.set(name, idx + 2);
    });

    // Write prizes to sheet
    let awarded = 0;
    const spentPoolEntries = [];

    preview.items.forEach(item => {
      const rowNum = nameToRow.get(item.player);
      if (rowNum) {
        sheet.getRange(rowNum, endPrizeIdx + 1).setValue(item.prizeName);
        awarded++;

        spentPoolEntries.push({
          eventId: eventId,
          itemCode: item.prizeCode,
          itemName: item.prizeName,
          level: item.rarity,
          qty: item.qty,
          cogs: 0, // Will be looked up if needed
          eventType: 'COMMANDER'
        });
      }
    });

    // Write to Spent_Pool if available
    if (typeof writeSpentPool === 'function' && spentPoolEntries.length > 0) {
      const batchId = typeof newBatchId === 'function' ? newBatchId() : `CMD-END-${Date.now()}`;
      writeSpentPool(spentPoolEntries, batchId);
    }

    // Update flags
    setCommanderFlag_(sheet, CMD_FLAGS.END, true);
    setCommanderFlag_(sheet, CMD_FLAGS.LOCKED, true);

    // Log to Integrity_Log
    logWizardAction_('COMMANDER_END_PRIZES', eventId, {
      prizesAwarded: awarded,
      estimatedCost: preview.estimatedCost,
      budget: preview.budget,
      rlBand: preview.rlBand
    });

    return {
      success: true,
      message: `Commander end prizes committed: ${awarded} player(s) awarded. Spent $${preview.estimatedCost?.toFixed(2) || '0.00'} of $${preview.budget?.toFixed(2) || '0.00'} budget.`,
      preview: preview
    };
  } catch (e) {
    logWizardAction_('COMMANDER_END_PRIZES_ERROR', eventId, {
      error: e.message
    });

    return {
      success: false,
      message: `Error committing end prizes: ${e.message}`
    };
  }
}

// ============================================================================
// Private Helpers
// ============================================================================

/**
 * Detects event format from sheet metadata or content.
 * @param {Sheet} sheet
 * @return {string} 'Commander' or 'Unknown'
 * @private
 */
function detectEventFormat_(sheet) {
  try {
    // Check developer metadata first
    if (typeof getEventProps === 'function') {
      const props = getEventProps(sheet);
      if (props.event_type === 'COMMANDER' || props.format === 'Commander') {
        return 'Commander';
      }
    }

    // Check A1 note for config JSON
    const a1Note = sheet.getRange('A1').getNote();
    if (a1Note) {
      try {
        const config = JSON.parse(a1Note);
        if (config.format === 'Commander' || config.event_type === 'COMMANDER') {
          return 'Commander';
        }
      } catch (e) {
        // Not JSON, check for text
        if (a1Note.toLowerCase().includes('commander')) {
          return 'Commander';
        }
      }
    }

    // Check sheet name
    const name = sheet.getName().toLowerCase();
    if (name.includes('commander') || name.includes('cmd')) {
      return 'Commander';
    }

    // Check for Commander-specific columns (R1_Prize, R2_Prize, R3_Prize)
    const headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
    const hasRoundPrizes = headers.some(h =>
      /r[123]_prize/i.test(String(h))
    );
    if (hasRoundPrizes) {
      return 'Commander';
    }

    return 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

/**
 * Counts players in an event sheet.
 * @param {Sheet} sheet
 * @return {number}
 * @private
 */
function countPlayers_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const prefNameIdx = findPreferredNameColumn_(headers);

    if (prefNameIdx >= 0) {
      const names = sheet.getRange(2, prefNameIdx + 1, lastRow - 1, 1).getValues();
      return names.filter(row => row[0] && String(row[0]).trim()).length;
    }

    return Math.max(0, lastRow - 1);
  } catch (e) {
    return 0;
  }
}

/**
 * Gets ranked players from sheet.
 * @param {Sheet} sheet
 * @return {Array<{rank: number, name: string, row: number}>}
 * @private
 */
function getRankedPlayers_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

    const rankIdx = headers.findIndex(h => String(h).toLowerCase() === 'rank');
    const prefNameIdx = findPreferredNameColumn_(headers);

    const players = data
      .map((row, idx) => ({
        rank: rankIdx >= 0 ? (parseInt(row[rankIdx]) || 999) : (idx + 1),
        name: prefNameIdx >= 0 ? String(row[prefNameIdx]).trim() : '',
        row: idx + 2
      }))
      .filter(p => p.name)
      .sort((a, b) => a.rank - b.rank);

    return players;
  } catch (e) {
    return [];
  }
}

/**
 * Finds the PreferredName column index.
 * @param {Array} headers
 * @return {number} Index or -1
 * @private
 */
function findPreferredNameColumn_(headers) {
  return headers.findIndex(h => {
    const lower = String(h).toLowerCase();
    return lower.includes('preferredname') ||
           lower.includes('preferred_name') ||
           lower === 'name' ||
           lower === 'player';
  });
}

/**
 * Gets Commander flags from event sheet.
 * Stores flags in developer metadata or a config row.
 * @param {Sheet} sheet
 * @return {{round1: boolean, round2: boolean, round3: boolean, end: boolean, locked: boolean}}
 * @private
 */
function getCommanderFlags_(sheet) {
  const defaults = { round1: false, round2: false, round3: false, end: false, locked: false };

  try {
    // Try developer metadata first
    const metadata = sheet.getDeveloperMetadata();
    const flags = { ...defaults };

    metadata.forEach(m => {
      const key = m.getKey();
      const value = m.getValue();
      if (key === CMD_FLAGS.ROUND1) flags.round1 = value === 'true';
      if (key === CMD_FLAGS.ROUND2) flags.round2 = value === 'true';
      if (key === CMD_FLAGS.ROUND3) flags.round3 = value === 'true';
      if (key === CMD_FLAGS.END) flags.end = value === 'true';
      if (key === CMD_FLAGS.LOCKED) flags.locked = value === 'true';
    });

    // Also check Integrity_Log for historical data
    const logFlags = checkIntegrityLogForFlags_(sheet.getName());
    flags.round1 = flags.round1 || logFlags.round1;
    flags.round2 = flags.round2 || logFlags.round2;
    flags.round3 = flags.round3 || logFlags.round3;
    flags.end = flags.end || logFlags.end;
    flags.locked = flags.locked || logFlags.end; // End implies locked

    return flags;
  } catch (e) {
    return defaults;
  }
}

/**
 * Sets a Commander flag on the event sheet.
 * @param {Sheet} sheet
 * @param {string} flagName
 * @param {boolean} value
 * @private
 */
function setCommanderFlag_(sheet, flagName, value) {
  try {
    // Remove existing metadata with this key
    const metadata = sheet.getDeveloperMetadata();
    metadata.forEach(m => {
      if (m.getKey() === flagName) {
        m.remove();
      }
    });

    // Add new metadata
    sheet.addDeveloperMetadata(flagName, String(value));
  } catch (e) {
    console.error('Failed to set flag:', flagName, e);
  }
}

/**
 * Checks Integrity_Log for previously awarded prizes.
 * @param {string} eventId
 * @return {{round1: boolean, round2: boolean, round3: boolean, end: boolean}}
 * @private
 */
function checkIntegrityLogForFlags_(eventId) {
  const flags = { round1: false, round2: false, round3: false, end: false };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('Integrity_Log');

    if (!logSheet || logSheet.getLastRow() <= 1) return flags;

    const data = logSheet.getDataRange().getValues();
    const headers = data[0];

    const eventIdx = headers.findIndex(h => String(h).toLowerCase().includes('event'));
    const actionIdx = headers.findIndex(h => String(h).toLowerCase().includes('action'));
    const statusIdx = headers.findIndex(h => String(h).toLowerCase().includes('status'));

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const event = eventIdx >= 0 ? row[eventIdx] : '';
      const action = actionIdx >= 0 ? String(row[actionIdx]).toUpperCase() : '';
      const status = statusIdx >= 0 ? row[statusIdx] : '';

      if (event !== eventId) continue;
      if (status !== 'SUCCESS') continue;

      if (action.includes('ROUND_1') || action === 'CMD_R1' || action === 'COMMANDER_ROUND_1') {
        flags.round1 = true;
      }
      if (action.includes('ROUND_2') || action === 'CMD_R2' || action === 'COMMANDER_ROUND_2') {
        flags.round2 = true;
      }
      if (action.includes('ROUND_3') || action === 'CMD_R3' || action === 'COMMANDER_ROUND_3') {
        flags.round3 = true;
      }
      if (action.includes('END') || action === 'CMD_END' || action === 'COMMANDER_END') {
        flags.end = true;
      }
    }

    return flags;
  } catch (e) {
    return flags;
  }
}

/**
 * Gets catalog items filtered by level.
 * @param {string} level - L0, L1, L2, L3, etc.
 * @return {Array<{Code: string, Name: string, Level: string, COGS: number, Qty: number}>}
 * @private
 */
function getCatalogItemsByLevel_(level) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const catalog = ss.getSheetByName('Prize_Catalog');

    if (!catalog || catalog.getLastRow() <= 1) {
      // Return defaults
      return getDefaultCatalogItems_(level);
    }

    const data = catalog.getDataRange().getValues();
    const headers = data[0];

    const codeIdx = headers.findIndex(h => String(h).toLowerCase() === 'code');
    const nameIdx = headers.findIndex(h => String(h).toLowerCase() === 'name');
    const levelIdx = headers.findIndex(h => String(h).toLowerCase() === 'level');
    const cogsIdx = headers.findIndex(h => String(h).toLowerCase() === 'cogs');
    const qtyIdx = headers.findIndex(h => String(h).toLowerCase() === 'qty');

    const items = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const itemLevel = levelIdx >= 0 ? String(row[levelIdx]).toUpperCase() : '';
      const qty = qtyIdx >= 0 ? parseInt(row[qtyIdx]) || 0 : 999;

      if (itemLevel === level.toUpperCase() && qty > 0) {
        items.push({
          Code: codeIdx >= 0 ? row[codeIdx] : `${level}-${i}`,
          Name: nameIdx >= 0 ? row[nameIdx] : `Prize ${level}`,
          Level: level,
          COGS: cogsIdx >= 0 ? parseFloat(row[cogsIdx]) || 0 : 2.00,
          Qty: qty
        });
      }
    }

    return items.length > 0 ? items : getDefaultCatalogItems_(level);
  } catch (e) {
    return getDefaultCatalogItems_(level);
  }
}

/**
 * Gets default catalog items for a level.
 * @param {string} level
 * @return {Array}
 * @private
 */
function getDefaultCatalogItems_(level) {
  const defaults = {
    'L0': [{ Code: 'L0-DEF', Name: 'Participation Prize', Level: 'L0', COGS: 1.00, Qty: 100 }],
    'L1': [{ Code: 'L1-DEF', Name: 'Booster Pack', Level: 'L1', COGS: 4.00, Qty: 50 }],
    'L2': [{ Code: 'L2-DEF', Name: 'Commander Prize Pack', Level: 'L2', COGS: 8.00, Qty: 20 }],
    'L3': [{ Code: 'L3-DEF', Name: 'Premium Pack', Level: 'L3', COGS: 15.00, Qty: 10 }]
  };

  return defaults[level] || defaults['L1'];
}

/**
 * Selects a prize from available items using RNG.
 * @param {Array} items
 * @param {Function} rng
 * @return {Object|null}
 * @private
 */
function selectPrize_(items, rng) {
  if (!items || items.length === 0) return null;
  const index = Math.floor(rng() * items.length);
  return items[index];
}

/**
 * Creates a seeded RNG function.
 * @param {string} seed
 * @return {Function}
 * @private
 */
function createWizardRng_(seed) {
  if (typeof createSeededRandom === 'function') {
    return createSeededRandom(seed);
  }

  // Simple LCG fallback
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = ((hash << 5) - hash) + seed.charCodeAt(i);
    hash = hash & hash;
  }
  let state = Math.abs(hash) || 1;

  return function() {
    state = (state * 1664525 + 1013904223) % 4294967296;
    return state / 4294967296;
  };
}

/**
 * Generates a random seed.
 * @return {string}
 * @private
 */
function generateWizardSeed_() {
  if (typeof generateSeed === 'function') {
    return generateSeed();
  }

  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let seed = '';
  for (let i = 0; i < 10; i++) {
    seed += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return seed;
}

/**
 * Logs a wizard action to Integrity_Log.
 * @param {string} action
 * @param {string} eventId
 * @param {Object} details
 * @private
 */
function logWizardAction_(action, eventId, details) {
  try {
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction(action, {
        eventId: eventId,
        details: JSON.stringify(details),
        status: details.error ? 'FAILURE' : 'SUCCESS'
      });
      return;
    }

    // Fallback: write directly to Integrity_Log
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Integrity_Log');

    if (!logSheet) {
      logSheet = ss.insertSheet('Integrity_Log');
      logSheet.appendRow([
        'Timestamp', 'Event_ID', 'Action', 'Operator', 'Details', 'Status'
      ]);
    }

    const timestamp = new Date().toISOString();
    const operator = Session.getEffectiveUser().getEmail() || 'system';
    const status = details.error ? 'FAILURE' : 'SUCCESS';

    logSheet.appendRow([
      timestamp,
      eventId,
      action,
      operator,
      JSON.stringify(details),
      status
    ]);
  } catch (e) {
    console.error('Failed to log wizard action:', e);
  }
}