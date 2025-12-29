/**
 * Award BP Backend Functions
 * @fileoverview Missing backend functions for ui/award_bp.html
 * 
 * ADD THIS FILE TO YOUR APPS SCRIPT PROJECT
 * 
 * These functions bridge the Award BP UI to your existing BP services.
 */

// ============================================================================
// REQUIRED: getCanonicalNames - Called by Award BP UI for player dropdown
// ============================================================================

/**
 * Gets list of canonical player names for the Award BP search dropdown
 * @return {string[]} Array of player names
 */
function getCanonicalNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PreferredNames');
  
  if (!sheet) {
    console.warn('PreferredNames sheet not found');
    return [];
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  return data
    .map(row => row[0])
    .filter(name => name && String(name).trim())
    .map(name => String(name).trim());
}

// ============================================================================
// REQUIRED: addDicePoints - Called by Award BP UI for D20/HYBRID sources
// ============================================================================

/**
 * Adds dice points to a player's Dice Roll Points sheet
 * Called from Award BP UI when source is D20 or HYBRID
 * 
 * Flow: Dice Roll Points → (sync) → BP_Total → (overflow) → BP_Prestige
 * 
 * @param {string} preferredName - Player name (already canonical from UI)
 * @param {number} amount - Points to add
 * @param {Object} metadata - {note, dfTags, source, timestamp}
 * @return {Object} {success, player, awarded, overflowToPrestige, error}
 */
function addDicePoints(preferredName, amount, metadata) {
  try {
    // Validate inputs
    if (!preferredName || typeof preferredName !== 'string') {
      return { success: false, error: 'Invalid player name' };
    }
    
    if (!amount || amount <= 0) {
      return { success: false, error: 'Amount must be positive' };
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Find the dice points sheet (support both naming conventions)
    let sheet = ss.getSheetByName('Dice Roll Points');
    if (!sheet) {
      sheet = ss.getSheetByName('Dice_Points');
    }
    
    // Create sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet('Dice Roll Points');
      sheet.appendRow(['PreferredName', 'Dice Roll Points', 'LastUpdated']);
      sheet.setFrozenRows(1);
      sheet.getRange('A1:C1')
        .setFontWeight('bold')
        .setBackground('#4285f4')
        .setFontColor('#ffffff');
      sheet.autoResizeColumns(1, 3);
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find columns with flexible matching
    const nameCol = findColumnIndex_(headers, ['PreferredName', 'Name', 'Player']);
    const pointsCol = findColumnIndex_(headers, ['Dice Roll Points', 'Dice_Points', 'DicePoints', 'Points']);
    const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated', 'Updated']);
    
    if (nameCol === -1) {
      return { success: false, error: 'PreferredName column not found in Dice Roll Points sheet' };
    }
    
    if (pointsCol === -1) {
      return { success: false, error: 'Points column not found in Dice Roll Points sheet' };
    }
    
    // Find existing player row
    let playerRow = -1;
    let currentPoints = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][nameCol] || '').trim();
      if (rowName.toLowerCase() === preferredName.toLowerCase()) {
        playerRow = i;
        currentPoints = Number(data[i][pointsCol]) || 0;
        break;
      }
    }
    
    const newPoints = currentPoints + amount;
    const now = new Date();
    
    if (playerRow === -1) {
      // Create new row for this player
      const newRow = new Array(headers.length).fill('');
      newRow[nameCol] = preferredName;
      newRow[pointsCol] = newPoints;
      if (updatedCol !== -1) newRow[updatedCol] = now;
      
      sheet.appendRow(newRow);
      playerRow = sheet.getLastRow() - 1; // 0-indexed for data array
    } else {
      // Update existing row
      const sheetRow = playerRow + 1; // 1-indexed for sheet
      sheet.getRange(sheetRow, pointsCol + 1).setValue(newPoints);
      if (updatedCol !== -1) {
        sheet.getRange(sheetRow, updatedCol + 1).setValue(now);
      }
    }
    
    // Trigger BP_Total sync if available
    let overflowToPrestige = 0;
    try {
      if (typeof updateBPTotalFromSources === 'function') {
        updateBPTotalFromSources();
      } else if (typeof refreshBP_Total === 'function') {
        refreshBP_Total();
      }
      // Check for overflow (would need to read BP_Total to get actual overflow)
    } catch (syncErr) {
      console.warn('BP sync after dice points:', syncErr);
    }
    
    // Log to Integrity_Log
    try {
      if (typeof logIntegrityAction === 'function') {
        logIntegrityAction('DICE_POINTS_AWARD', {
          preferredName: preferredName,
          dfTags: metadata?.dfTags || ['DF-080', 'DF-081'],
          details: `Source: ${metadata?.source || 'D20'} | +${amount} dice points | ${currentPoints} → ${newPoints}`,
          status: 'SUCCESS'
        });
      }
    } catch (logErr) {
      console.warn('Failed to log dice points award:', logErr);
    }
    
    return {
      success: true,
      player: preferredName,
      awarded: amount,
      previousPoints: currentPoints,
      newPoints: newPoints,
      overflowToPrestige: overflowToPrestige
    };
    
  } catch (e) {
    console.error('addDicePoints error:', e);
    return {
      success: false,
      error: e.message || String(e),
      player: preferredName
    };
  }
}

// ============================================================================
// REQUIRED: awardBP - Called by Award BP UI for non-dice sources
// ============================================================================

/**
 * Awards BP directly to BP_Total (for non-dice sources)
 * Called from Award BP UI when source is MANUAL, ATTENDANCE, TOP4, etc.
 * 
 * This is a UI-facing wrapper that ensures consistent return shape.
 * 
 * @param {string} preferredName - Player name
 * @param {number} amount - BP to award
 * @param {string} source - Source type (MANUAL, ATTENDANCE, TOP4, FLAG_MISSION, etc.)
 * @param {Object} metadata - {note, dfTags, timestamp}
 * @return {Object} {success, player, awarded, currentBP, overflowToPrestige, error}
 */
function awardBP(preferredName, amount, source, metadata) {
  try {
    // Validate inputs
    if (!preferredName || typeof preferredName !== 'string') {
      return { success: false, error: 'Invalid player name' };
    }
    
    if (!amount || amount <= 0) {
      return { success: false, error: 'Amount must be positive' };
    }
    
    // Try to use existing awardBonusPoints if available
    if (typeof awardBonusPoints === 'function') {
      const result = awardBonusPoints(preferredName, amount, source || 'MANUAL', metadata || {});
      
      // Normalize return shape for UI
      return {
        success: result.success,
        player: result.player || preferredName,
        awarded: result.awarded || amount,
        currentBP: result.currentBP || 0,
        overflowToPrestige: result.overflow || 0,
        error: result.error
      };
    }
    
    // Fallback: Direct BP_Total update
    return awardBPDirect_(preferredName, amount, source, metadata);
    
  } catch (e) {
    console.error('awardBP error:', e);
    return {
      success: false,
      error: e.message || String(e),
      player: preferredName
    };
  }
}

/**
 * Direct BP_Total update fallback
 * Used when awardBonusPoints is not available
 * @private
 */
function awardBPDirect_(preferredName, amount, source, metadata) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Total');
  
  // Create BP_Total if missing
  if (!sheet) {
    sheet = ss.insertSheet('BP_Total');
    sheet.appendRow(['PreferredName', 'Current_BP', 'Historical_BP', 'LastUpdated']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nameCol = findColumnIndex_(headers, ['PreferredName', 'Name']);
  const currentBPCol = findColumnIndex_(headers, ['Current_BP', 'BP_Current', 'CurrentBP']);
  const historicalCol = findColumnIndex_(headers, ['Historical_BP', 'HistoricalBP']);
  const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated']);
  
  if (nameCol === -1 || currentBPCol === -1) {
    return { success: false, error: 'Invalid BP_Total schema' };
  }
  
  // Get global cap (default 100)
  let globalCap = 100;
  try {
    if (typeof getThrottleNumber === 'function') {
      globalCap = getThrottleNumber('BP_Global_Cap', 100);
    }
  } catch (e) {
    // Use default
  }
  
  // Find player row
  let playerRow = -1;
  let currentBP = 0;
  let historicalBP = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim().toLowerCase() === preferredName.toLowerCase()) {
      playerRow = i;
      currentBP = Number(data[i][currentBPCol]) || 0;
      if (historicalCol !== -1) {
        historicalBP = Number(data[i][historicalCol]) || 0;
      }
      break;
    }
  }
  
  const newTotal = currentBP + amount;
  const clampedBP = Math.min(newTotal, globalCap);
  const overflow = Math.max(0, newTotal - globalCap);
  const now = new Date();
  
  if (playerRow === -1) {
    // Create new row
    const newRow = new Array(headers.length).fill('');
    newRow[nameCol] = preferredName;
    newRow[currentBPCol] = clampedBP;
    if (historicalCol !== -1) newRow[historicalCol] = amount;
    if (updatedCol !== -1) newRow[updatedCol] = now;
    
    sheet.appendRow(newRow);
  } else {
    // Update existing
    const sheetRow = playerRow + 1;
    sheet.getRange(sheetRow, currentBPCol + 1).setValue(clampedBP);
    if (historicalCol !== -1) {
      sheet.getRange(sheetRow, historicalCol + 1).setValue(historicalBP + amount);
    }
    if (updatedCol !== -1) {
      sheet.getRange(sheetRow, updatedCol + 1).setValue(now);
    }
  }
  
  // Handle overflow to prestige
  if (overflow > 0) {
    try {
      if (typeof addPrestigeOverflow_ === 'function') {
        addPrestigeOverflow_(preferredName, overflow);
      } else {
        addPrestigeOverflowFallback_(preferredName, overflow);
      }
    } catch (e) {
      console.warn('Failed to add prestige overflow:', e);
    }
  }
  
  // Log to Integrity_Log
  try {
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('BP_AWARD', {
        preferredName: preferredName,
        dfTags: metadata?.dfTags || ['DF-080', 'DF-081'],
        details: `Source: ${source || 'MANUAL'} | +${amount} BP | ${currentBP} → ${clampedBP} (overflow: ${overflow})`,
        status: 'SUCCESS'
      });
    }
  } catch (e) {
    console.warn('Failed to log BP award:', e);
  }
  
  return {
    success: true,
    player: preferredName,
    awarded: amount,
    currentBP: clampedBP,
    overflowToPrestige: overflow
  };
}

/**
 * Fallback prestige overflow handler
 * @private
 */
function addPrestigeOverflowFallback_(preferredName, overflow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('BP_Prestige');
  
  if (!sheet) {
    sheet = ss.insertSheet('BP_Prestige');
    sheet.appendRow(['PreferredName', 'Prestige_Points', 'Last_Updated']);
    sheet.setFrozenRows(1);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nameCol = 0;
  const prestigeCol = findColumnIndex_(headers, ['Prestige_Points', 'PrestigePoints', 'Prestige']);
  const updatedCol = findColumnIndex_(headers, ['Last_Updated', 'LastUpdated']);
  
  if (prestigeCol === -1) return;
  
  let playerRow = -1;
  let currentPrestige = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][nameCol]).trim().toLowerCase() === preferredName.toLowerCase()) {
      playerRow = i;
      currentPrestige = Number(data[i][prestigeCol]) || 0;
      break;
    }
  }
  
  const newPrestige = currentPrestige + overflow;
  const now = new Date();
  
  if (playerRow === -1) {
    const newRow = [preferredName, newPrestige];
    if (updatedCol !== -1) newRow.push(now);
    sheet.appendRow(newRow);
  } else {
    const sheetRow = playerRow + 1;
    sheet.getRange(sheetRow, prestigeCol + 1).setValue(newPrestige);
    if (updatedCol !== -1) {
      sheet.getRange(sheetRow, updatedCol + 1).setValue(now);
    }
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Finds column index with flexible header matching
 * @param {Array} headers - Header row
 * @param {Array<string>} possibleNames - Possible column names (case-insensitive)
 * @return {number} Column index or -1 if not found
 * @private
 */
function findColumnIndex_(headers, possibleNames) {
  const lowerHeaders = headers.map(h => String(h).toLowerCase().trim());
  
  for (const name of possibleNames) {
    const idx = lowerHeaders.indexOf(name.toLowerCase());
    if (idx !== -1) return idx;
  }
  
  // Try partial match
  for (const name of possibleNames) {
    const lowerName = name.toLowerCase();
    for (let i = 0; i < lowerHeaders.length; i++) {
      if (lowerHeaders[i].includes(lowerName) || lowerName.includes(lowerHeaders[i])) {
        return i;
      }
    }
  }
  
  return -1;
}