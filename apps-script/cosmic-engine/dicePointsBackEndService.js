/**
 * Award Dice Points - Routes through Dice_Points → BP_Total
 * Called by Award BP UI when source is D20 or HYBRID
 * 
 * @param {string} playerName - Player's preferred name
 * @param {number} amount - Points to award
 * @param {string} source - D20 or HYBRID
 * @param {Object} metadata - { note, dfTags, routeViaDicePoints }
 * @returns {Object} { success, awarded, player, currentBP, prestige, overflow, error }
 */
function awardDicePoints(playerName, amount, source, metadata) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  
  try {
    // Acquire lock to prevent race conditions
    lock.waitLock(10000);
    
    // Validate inputs
    if (!playerName || typeof playerName !== 'string') {
      return { success: false, error: 'Invalid player name' };
    }
    if (!amount || amount <= 0 || !Number.isInteger(amount)) {
      return { success: false, error: 'Amount must be a positive integer' };
    }
    if (!['D20', 'HYBRID'].includes(source)) {
      return { success: false, error: 'Invalid source for dice points. Use D20 or HYBRID.' };
    }
    
    // Resolve canonical player name
    const canonicalName = resolveCanonicalName_(playerName);
    if (!canonicalName) {
      return { success: false, error: `Player "${playerName}" not found in PreferredNames` };
    }
    
    // Step 1: Write to Dice_Points sheet
    const diceResult = writeToDicePoints_(ss, canonicalName, amount, source, metadata);
    if (!diceResult.success) {
      return diceResult;
    }
    
    // Step 2: Aggregate Dice_Points → BP_Total
    const aggregateResult = aggregateDiceToBPTotal_(ss, canonicalName);
    if (!aggregateResult.success) {
      return aggregateResult;
    }
    
    // Step 3: Apply governor layer (cap at 100, overflow to prestige)
    const governorResult = applyBPGovernor_(ss, canonicalName);
    
    // Step 4: Log to Integrity_Log
    logToIntegrityLog_(ss, {
      timestamp: new Date(),
      player: canonicalName,
      action: 'DICE_POINTS_AWARD',
      amount: amount,
      source: source,
      note: metadata?.note || '',
      dfTags: metadata?.dfTags?.join(', ') || 'DF-080, DF-081, DF-110, DF-115',
      route: 'Dice_Points → BP_Total → Prestige',
      resultBP: governorResult.currentBP,
      resultPrestige: governorResult.prestige,
      overflow: governorResult.overflow
    });
    
    return {
      success: true,
      awarded: amount,
      player: canonicalName,
      currentBP: governorResult.currentBP,
      prestige: governorResult.prestige,
      overflow: governorResult.overflow,
      route: 'Dice_Points → BP_Total'
    };
    
  } catch (e) {
    console.error('awardDicePoints error:', e);
    return { success: false, error: e.message || 'Unknown error' };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Write points to Dice_Points sheet
 * @private
 */
function writeToDicePoints_(ss, playerName, amount, source, metadata) {
  try {
    const sheet = ss.getSheetByName('Dice_Points');
    if (!sheet) {
      return { success: false, error: 'Dice_Points sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const playerCol = headers.indexOf('preferred_name_id');
    const pointsCol = headers.indexOf('Points');
    const sourceCol = headers.indexOf('Source'); // Optional column for tracking
    const lastUpdatedCol = headers.indexOf('LastUpdated');
    
    if (playerCol === -1 || pointsCol === -1) {
      return { success: false, error: 'Dice_Points sheet missing required columns (preferred_name_id, Points)' };
    }
    
    // Find player row
    let playerRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][playerCol] === playerName) {
        playerRow = i + 1; // 1-indexed for sheet
        break;
      }
    }
    
    if (playerRow === -1) {
      // Player doesn't exist in Dice_Points - add new row
      const newRow = new Array(headers.length).fill('');
      newRow[playerCol] = playerName;
      newRow[pointsCol] = amount;
      if (sourceCol !== -1) newRow[sourceCol] = source;
      if (lastUpdatedCol !== -1) newRow[lastUpdatedCol] = new Date();
      sheet.appendRow(newRow);
    } else {
      // Update existing player
      const currentPoints = Number(data[playerRow - 1][pointsCol]) || 0;
      const newPoints = currentPoints + amount;
      
      sheet.getRange(playerRow, pointsCol + 1).setValue(newPoints);
      if (sourceCol !== -1) {
        sheet.getRange(playerRow, sourceCol + 1).setValue(source);
      }
      if (lastUpdatedCol !== -1) {
        sheet.getRange(playerRow, lastUpdatedCol + 1).setValue(new Date());
      }
    }
    
    return { success: true };
    
  } catch (e) {
    console.error('writeToDicePoints_ error:', e);
    return { success: false, error: 'Failed to write to Dice_Points: ' + e.message };
  }
}


/**
 * Aggregate Dice_Points into BP_Total for a specific player
 * @private
 */
function aggregateDiceToBPTotal_(ss, playerName) {
  try {
    // Read current Dice_Points for player
    const diceSheet = ss.getSheetByName('Dice_Points');
    const bpTotalSheet = ss.getSheetByName('BP_Total');
    
    if (!diceSheet || !bpTotalSheet) {
      return { success: false, error: 'Required sheets not found (Dice_Points or BP_Total)' };
    }
    
    // Get dice points for player
    const diceData = diceSheet.getDataRange().getValues();
    const diceHeaders = diceData[0];
    const dicePlayerCol = diceHeaders.indexOf('preferred_name_id');
    const dicePointsCol = diceHeaders.indexOf('Points');
    
    let dicePoints = 0;
    for (let i = 1; i < diceData.length; i++) {
      if (diceData[i][dicePlayerCol] === playerName) {
        dicePoints = Number(diceData[i][dicePointsCol]) || 0;
        break;
      }
    }
    
    // Update BP_Total
    const bpData = bpTotalSheet.getDataRange().getValues();
    const bpHeaders = bpData[0];
    const bpPlayerCol = bpHeaders.indexOf('Player');
    const bpDiceCol = bpHeaders.indexOf('Dice Roll Points');
    
    if (bpPlayerCol === -1) {
      return { success: false, error: 'BP_Total missing Player column' };
    }
    
    // Find or create player row in BP_Total
    let bpPlayerRow = -1;
    for (let i = 1; i < bpData.length; i++) {
      if (bpData[i][bpPlayerCol] === playerName) {
        bpPlayerRow = i + 1;
        break;
      }
    }
    
    if (bpPlayerRow === -1) {
      // Create new row
      const newRow = new Array(bpHeaders.length).fill('');
      newRow[bpPlayerCol] = playerName;
      if (bpDiceCol !== -1) newRow[bpDiceCol] = dicePoints;
      bpTotalSheet.appendRow(newRow);
    } else if (bpDiceCol !== -1) {
      // Update existing row
      bpTotalSheet.getRange(bpPlayerRow, bpDiceCol + 1).setValue(dicePoints);
    }
    
    return { success: true, dicePoints: dicePoints };
    
  } catch (e) {
    console.error('aggregateDiceToBPTotal_ error:', e);
    return { success: false, error: 'Failed to aggregate to BP_Total: ' + e.message };
  }
}


/**
 * Apply BP Governor Layer - enforce 0-100 cap, overflow to Prestige
 * Per DF-080 (cap) and DF-081 (prestige overflow)
 * @private
 */
function applyBPGovernor_(ss, playerName) {
  const BP_CAP = 100;
  
  try {
    const bpTotalSheet = ss.getSheetByName('BP_Total');
    const prestigeSheet = ss.getSheetByName('BP_Prestige');
    
    if (!bpTotalSheet) {
      return { success: false, error: 'BP_Total sheet not found' };
    }
    
    const bpData = bpTotalSheet.getDataRange().getValues();
    const bpHeaders = bpData[0];
    
    // Find columns
    const playerCol = bpHeaders.indexOf('Player');
    const attendanceCol = bpHeaders.indexOf('Attendance Missions');
    const flagCol = bpHeaders.indexOf('Flag Missions');
    const diceCol = bpHeaders.indexOf('Dice Roll Points');
    const totalCol = bpHeaders.indexOf('Total_Points');
    const currentBPCol = bpHeaders.indexOf('Current_BP');
    const historicalCol = bpHeaders.indexOf('Historical_BP');
    
    // Find player row
    let playerRow = -1;
    let rowData = null;
    for (let i = 1; i < bpData.length; i++) {
      if (bpData[i][playerCol] === playerName) {
        playerRow = i + 1;
        rowData = bpData[i];
        break;
      }
    }
    
    if (playerRow === -1 || !rowData) {
      return { success: false, error: 'Player not found in BP_Total' };
    }
    
    // Calculate total from rivers
    const attendance = Number(rowData[attendanceCol]) || 0;
    const flags = Number(rowData[flagCol]) || 0;
    const dice = Number(rowData[diceCol]) || 0;
    const rawTotal = attendance + flags + dice;
    
    // Get current historical and calculate current BP
    let historical = Number(rowData[historicalCol]) || 0;
    historical = Math.max(historical, rawTotal); // Historical never decreases
    
    // Get redeemed BP
    const redeemedSheet = ss.getSheetByName('Redeemed_BP');
    let redeemed = 0;
    if (redeemedSheet) {
      const redeemedData = redeemedSheet.getDataRange().getValues();
      const redeemedHeaders = redeemedData[0];
      const redPlayerCol = redeemedHeaders.indexOf('Player');
      const redTotalCol = redeemedHeaders.indexOf('Total_Redeemed');
      
      for (let i = 1; i < redeemedData.length; i++) {
        if (redeemedData[i][redPlayerCol] === playerName) {
          redeemed = Number(redeemedData[i][redTotalCol]) || 0;
          break;
        }
      }
    }
    
    // Calculate current BP (before cap)
    let currentBP = historical - redeemed;
    let overflow = 0;
    let prestige = 0;
    
    // Apply DF-080: Cap at 100
    if (currentBP > BP_CAP) {
      overflow = currentBP - BP_CAP;
      currentBP = BP_CAP;
      
      // Apply DF-081: Overflow → Prestige
      if (prestigeSheet) {
        prestige = updatePrestige_(prestigeSheet, playerName, overflow);
      }
    }
    
    // Update BP_Total with calculated values
    if (totalCol !== -1) {
      bpTotalSheet.getRange(playerRow, totalCol + 1).setValue(rawTotal);
    }
    if (currentBPCol !== -1) {
      bpTotalSheet.getRange(playerRow, currentBPCol + 1).setValue(currentBP);
    }
    if (historicalCol !== -1) {
      bpTotalSheet.getRange(playerRow, historicalCol + 1).setValue(historical);
    }
    
    return {
      success: true,
      currentBP: currentBP,
      prestige: prestige,
      overflow: overflow,
      historical: historical
    };
    
  } catch (e) {
    console.error('applyBPGovernor_ error:', e);
    return { success: false, error: 'Governor layer error: ' + e.message, currentBP: 0, prestige: 0, overflow: 0 };
  }
}


/**
 * Update prestige for overflow points
 * @private
 */
function updatePrestige_(prestigeSheet, playerName, overflowAmount) {
  try {
    const data = prestigeSheet.getDataRange().getValues();
    const headers = data[0];
    const playerCol = headers.indexOf('Player');
    const prestigeCol = headers.indexOf('Prestige');
    const overflowCol = headers.indexOf('Total_Overflow');
    
    if (playerCol === -1 || prestigeCol === -1) {
      return 0;
    }
    
    // Find player
    for (let i = 1; i < data.length; i++) {
      if (data[i][playerCol] === playerName) {
        const currentPrestige = Number(data[i][prestigeCol]) || 0;
        const totalOverflow = (Number(data[i][overflowCol]) || 0) + overflowAmount;
        
        // 1 Prestige per 100 overflow (or your desired ratio)
        const newPrestige = Math.floor(totalOverflow / 100);
        
        prestigeSheet.getRange(i + 1, prestigeCol + 1).setValue(newPrestige);
        if (overflowCol !== -1) {
          prestigeSheet.getRange(i + 1, overflowCol + 1).setValue(totalOverflow);
        }
        
        return newPrestige;
      }
    }
    
    // Player not in prestige sheet - add them
    const newRow = new Array(headers.length).fill('');
    newRow[playerCol] = playerName;
    newRow[prestigeCol] = 0;
    if (overflowCol !== -1) newRow[overflowCol] = overflowAmount;
    prestigeSheet.appendRow(newRow);
    
    return 0;
    
  } catch (e) {
    console.error('updatePrestige_ error:', e);
    return 0;
  }
}


/**
 * Log action to Integrity_Log
 * @private
 */
function logToIntegrityLog_(ss, logData) {
  try {
    const sheet = ss.getSheetByName('Integrity_Log');
    if (!sheet) {
      console.warn('Integrity_Log sheet not found');
      return;
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(h => {
      const key = h.toLowerCase().replace(/[^a-z0-9]/g, '');
      // Map header to logData keys
      if (key === 'timestamp') return logData.timestamp;
      if (key === 'player') return logData.player;
      if (key === 'action') return logData.action;
      if (key === 'amount') return logData.amount;
      if (key === 'source') return logData.source;
      if (key === 'note') return logData.note;
      if (key === 'dftags' || key === 'decisionflags') return logData.dfTags;
      if (key === 'route') return logData.route;
      if (key === 'resultbp' || key === 'currentbp') return logData.resultBP;
      if (key === 'prestige') return logData.resultPrestige;
      if (key === 'overflow') return logData.overflow;
      return '';
    });
    
    sheet.appendRow(newRow);
    
  } catch (e) {
    console.error('logToIntegrityLog_ error:', e);
  }
}


/**
 * Resolve player name to canonical PreferredNames entry
 * @private
 */
function resolveCanonicalName_(inputName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PreferredNames');
    if (!sheet) return inputName; // Fallback
    
    const data = sheet.getDataRange().getValues();
    const normalizedInput = inputName.toLowerCase().trim();
    
    // Exact match first
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase().trim() === normalizedInput) {
        return data[i][0];
      }
    }
    
    // Fuzzy match (contains)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toLowerCase().includes(normalizedInput)) {
        return data[i][0];
      }
    }
    
    return null;
    
  } catch (e) {
    console.error('resolveCanonicalName_ error:', e);
    return inputName;
  }
}


/**
 * Get canonical names from PreferredNames for UI dropdown
 * (You may already have this - included for completeness)
 */
function getCanonicalNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('PreferredNames');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const names = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && String(data[i][0]).trim()) {
        names.push(String(data[i][0]).trim());
      }
    }
    
    return names.sort((a, b) => a.localeCompare(b));
    
  } catch (e) {
    console.error('getCanonicalNames error:', e);
    return [];
  }
}