/**
 * Key Service - Player Key Management
 * @fileoverview Manages Key_Tracker: add keys, rainbow conversion
 * 
 * UNLOCK RULES:
 * - Player needs to cover all 5 colors (Red, Blue, Green, Yellow, Purple)
 * - Each color can be covered by: having ≥1 key of that color, OR using a Rainbow (wild)
 * - Rainbow keys act as wildcards for ANY missing color
 * - Player can trade 3 keys of the SAME color → 1 Rainbow key (no cap on conversions)
 * 
 * EXAMPLES:
 * (1,1,1,1,1) + 0 Rainbow = WIN (all 5 colors)
 * (1,1,1,4,0) = WIN (4 colors + convert 3 from the 4-stack → 1 Rainbow for missing color)
 * (1,1,7,0,0) = WIN (3 colors + convert 6 from 7-stack → 2 Rainbows for 2 missing)
 * (1,10,0,0,0) = WIN (2 colors + convert 9 from 10-stack → 3 Rainbows for 3 missing)
 * (13,0,0,0,0) = WIN (1 color + convert 12 → 4 Rainbows for 4 missing)
 */

// ============================================================================
// KEY OPERATIONS
// ============================================================================

/**
 * Adds key(s) to a player
 * @param {string} preferredName - Player name
 * @param {string} color - Key color: Red, Blue, Green, Yellow, Purple
 * @param {number} qty - Quantity (default: 1)
 * @return {Object} Result {before, after, added}
 */
function addKey(preferredName, color, qty = 1) {
  ensureKeyTrackerSchema();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');
  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }
  
  // Validate color
  const validColors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  if (!validColors.includes(color)) {
    throwError('Invalid key color', 'INVALID_COLOR', `Must be one of: ${validColors.join(', ')}`);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const colorCol = headers.indexOf(color);
  
  if (nameCol === -1 || colorCol === -1) {
    throwError('Invalid Key_Tracker schema', 'SCHEMA_INVALID');
  }
  
  let playerRow = -1;
  let currentQty = 0;
  
  // Find player row
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      currentQty = coerceNumber(data[i][colorCol], 0);
      break;
    }
  }
  
  // Create row if player not found
  if (playerRow === -1) {
    const newRow = new Array(headers.length).fill(0);
    newRow[nameCol] = preferredName;
    newRow[colorCol] = qty;
    newRow[headers.indexOf('LastUpdated')] = dateISO();
    sheet.appendRow(newRow);
    
    logIntegrityAction('KEY_AWARD', {
      preferredName,
      details: `Added ${qty} ${color} (new player)`,
      status: 'SUCCESS'
    });
    
    // Update eligibility for new player
    const throttle = getThrottleKV();
    const ratioStr = throttle.Rainbow_Rate || '3:1';
    const ratio = parseRatio(ratioStr);
    if (ratio && ratio.den === 1) {
      // Re-fetch data to get new row
      const newData = sheet.getDataRange().getValues();
      updatePlayerEligibility_(preferredName, sheet, newData[0], ratio.num);
    }
    
    return {
      before: 0,
      after: qty,
      added: qty
    };
  }
  
  // Update existing row
  const newQty = currentQty + qty;
  sheet.getRange(playerRow + 1, colorCol + 1).setValue(newQty);
  sheet.getRange(playerRow + 1, headers.indexOf('LastUpdated') + 1).setValue(dateISO());

  // Update eligibility and check if player just unlocked
  const throttle = getThrottleKV();
  const ratioStr = throttle.Rainbow_Rate || '3:1';
  const ratio = parseRatio(ratioStr);
  if (ratio && ratio.den === 1) {
    const justUnlocked = updatePlayerEligibility_(preferredName, sheet, headers, ratio.num);
    if (justUnlocked) {
      showUnlockEligiblePopup_(preferredName, sheet, headers);
    }
  }

  logIntegrityAction('KEY_AWARD', {
    preferredName,
    details: `${color}: ${currentQty} → ${newQty} (+${qty})`,
    status: 'SUCCESS'
  });
  
  return {
    before: currentQty,
    after: newQty,
    added: qty
  };
}

/**
 * Converts keys to Rainbow using ratio (e.g., "3:1")
 * Converts from a SINGLE color (3 of same color → 1 Rainbow)
 * @param {string} preferredName - Player name
 * @param {string} sourceColor - Color to convert FROM (Red, Blue, Green, Yellow, Purple)
 * @param {number} setsToConvert - Number of sets to convert (default: 1)
 * @param {string} ratioStr - Ratio string (e.g., "3:1")
 * @return {Object} Result {converted, rainbowCount, remainingColor}
 */
function convertRainbow(preferredName, sourceColor, setsToConvert = 1, ratioStr = '3:1') {
  ensureKeyTrackerSchema();
  
  const ratio = parseRatio(ratioStr);
  if (!ratio || ratio.den !== 1) {
    throwError('Invalid ratio', 'INVALID_RATIO', 'Must be "X:1" format');
  }
  
  const validColors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  if (!validColors.includes(sourceColor)) {
    throwError('Invalid source color', 'INVALID_COLOR', `Must be one of: ${validColors.join(', ')}`);
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');
  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const rainbowCol = headers.indexOf('RainbowEligible');
  const sourceCol = headers.indexOf(sourceColor);
  
  if (nameCol === -1 || rainbowCol === -1 || sourceCol === -1) {
    throwError('Invalid Key_Tracker schema', 'SCHEMA_INVALID');
  }
  
  // Find player
  let playerRow = -1;
  let currentSourceQty = 0;
  let currentRainbow = 0;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      playerRow = i;
      currentSourceQty = coerceNumber(data[i][sourceCol], 0);
      currentRainbow = coerceNumber(data[i][rainbowCol], 0);
      break;
    }
  }
  
  if (playerRow === -1) {
    throwError('Player not found', 'PLAYER_NOT_FOUND', `No entry for "${preferredName}"`);
  }
  
  // Check if player has enough keys
  const keysNeeded = setsToConvert * ratio.num;
  if (currentSourceQty < keysNeeded) {
    throwError('Insufficient keys', 'INSUFFICIENT_KEYS', 
      `Need ${keysNeeded} ${sourceColor} keys, have ${currentSourceQty}`);
  }
  
  // Deduct keys from source color
  const newSourceQty = currentSourceQty - keysNeeded;
  sheet.getRange(playerRow + 1, sourceCol + 1).setValue(newSourceQty);
  
  // Add to rainbow count (NO CAP)
  const newRainbow = currentRainbow + setsToConvert;
  sheet.getRange(playerRow + 1, rainbowCol + 1).setValue(newRainbow);
  sheet.getRange(playerRow + 1, headers.indexOf('LastUpdated') + 1).setValue(dateISO());

  // Update eligibility and check if player just unlocked
  const justUnlocked = updatePlayerEligibility_(preferredName, sheet, headers, ratio.num);
  if (justUnlocked) {
    showUnlockEligiblePopup_(preferredName, sheet, headers);
  }

  logIntegrityAction('RAINBOW_CONVERT', {
    preferredName,
    details: `Converted ${setsToConvert} set(s) from ${sourceColor} (${currentSourceQty}→${newSourceQty}) → ${newRainbow} Rainbow total`,
    status: 'SUCCESS'
  });
  
  return {
    converted: setsToConvert,
    rainbowCount: newRainbow,
    sourceColor: sourceColor,
    remainingSourceKeys: newSourceQty
  };
}

/**
 * Gets player's key counts
 * @param {string} preferredName - Player name
 * @return {Object|null} Key counts or null if not found
 */
function getPlayerKeys(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  if (nameCol === -1) return null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = data[i][idx];
      });
      return row;
    }
  }
  
  return null;
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures Key_Tracker has required schema
 */
function ensureKeyTrackerSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Key_Tracker');
  
  const requiredHeaders = [
    'PreferredName',
    'Red',
    'Blue',
    'Green',
    'Yellow',
    'Purple',
    'RainbowEligible',
    'Able to Unlock?',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Key_Tracker');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:I1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    return;
  }
  
  // Check headers
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));
  
  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}

// ============================================================================
// UNLOCK ELIGIBILITY TRACKING
// ============================================================================

/**
 * Calculates unlock eligibility with UNLIMITED rainbow conversions
 * 
 * RULE: Player wins if they can cover all 5 colors using:
 *   - Keys they have (≥1 of a color = that color is covered)
 *   - Rainbow keys (each covers 1 missing color)
 *   - Potential conversions: 3 keys of same color → 1 Rainbow (no cap)
 * 
 * @param {Array<number>} colorQtys - [Red, Blue, Green, Yellow, Purple]
 * @param {number} rainbowQty - Current Rainbow key count
 * @param {number} conversionRatio - Keys needed per Rainbow (e.g., 3 for 3:1)
 * @returns {Object} - {eligible: 0|1, canUnlockNow: boolean, details: string}
 * @private
 */
function calculateUnlockEligibility_(colorQtys, rainbowQty, conversionRatio) {
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  
  // Count colors the player already has (≥1 key)
  const colorsHeld = colorQtys.filter(qty => qty >= 1).length;
  const colorsMissing = 5 - colorsHeld;
  
  // Current rainbows cover missing colors
  let rainbowsAvailable = rainbowQty;
  
  // Calculate potential additional rainbows from conversions
  // For each color, count how many EXCESS keys can be converted
  // (keeping at least 1 to maintain the color)
  let potentialRainbows = 0;
  
  for (let i = 0; i < colorQtys.length; i++) {
    const qty = colorQtys[i];
    if (qty >= 1) {
      // Keep 1 to maintain color coverage, convert excess
      const excessKeys = qty - 1;
      potentialRainbows += Math.floor(excessKeys / conversionRatio);
    } else {
      // Color not held - all keys here can be converted
      // (but wait, if qty is 0 there are no keys to convert)
      // This case: qty === 0, so no contribution
    }
  }
  
  // Total wildcards available (current + potential)
  const totalWildcards = rainbowsAvailable + potentialRainbows;
  
  // Can unlock if wildcards cover all missing colors
  const canUnlock = totalWildcards >= colorsMissing;
  
  // Build details for reporting
  let details = '';
  if (canUnlock) {
    if (colorsMissing === 0) {
      details = 'All 5 colors held!';
    } else if (rainbowsAvailable >= colorsMissing) {
      details = `${colorsHeld} colors + ${colorsMissing} Rainbow(s) cover missing`;
    } else {
      const conversionsNeeded = colorsMissing - rainbowsAvailable;
      details = `${colorsHeld} colors + ${rainbowsAvailable} Rainbow + convert ${conversionsNeeded} more`;
    }
  } else {
    const shortfall = colorsMissing - totalWildcards;
    details = `Need ${shortfall} more Rainbow (or keys to convert)`;
  }
  
  return {
    eligible: canUnlock ? 1 : 0,
    canUnlockNow: rainbowsAvailable >= colorsMissing || colorsMissing === 0,
    colorsHeld,
    colorsMissing,
    rainbowsAvailable,
    potentialRainbows,
    totalWildcards,
    details
  };
}

/**
 * Checks if a player can unlock RIGHT NOW (without needing conversions)
 * @param {Array<number>} colorQtys - [Red, Blue, Green, Yellow, Purple]
 * @param {number} rainbowQty - Current Rainbow key count
 * @returns {boolean}
 */
function canUnlockNow_(colorQtys, rainbowQty) {
  const colorsHeld = colorQtys.filter(qty => qty >= 1).length;
  const colorsMissing = 5 - colorsHeld;
  return rainbowQty >= colorsMissing;
}

/**
 * Updates eligibility for a single player (called after key operations)
 * @param {string} preferredName - Player name
 * @param {Sheet} sheet - Key_Tracker sheet
 * @param {Array} headers - Header row
 * @param {number} conversionRatio - e.g., 3 for 3:1 ratio
 * @returns {boolean} - true if player just became eligible (0→1 transition)
 * @private
 */
function updatePlayerEligibility_(preferredName, sheet, headers, conversionRatio) {
  const nameCol = headers.indexOf('PreferredName');
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');
  const unlockCol = headers.indexOf('Able to Unlock?');

  if (unlockCol === -1) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const colorQtys = colorCols.map(col => coerceNumber(data[i][col], 0));
      const rainbowQty = rainbowCol !== -1 ? coerceNumber(data[i][rainbowCol], 0) : 0;
      const previousEligibility = coerceNumber(data[i][unlockCol], 0);

      const result = calculateUnlockEligibility_(colorQtys, rainbowQty, conversionRatio);

      sheet.getRange(i + 1, unlockCol + 1).setValue(result.eligible);

      return previousEligibility === 0 && result.eligible === 1;
    }
  }

  return false;
}

/**
 * Shows popup notification when player becomes eligible to unlock
 * @param {string} preferredName - Player name
 * @param {Sheet} sheet - Key_Tracker sheet
 * @param {Array} headers - Header row
 * @private
 */
function showUnlockEligiblePopup_(preferredName, sheet, headers) {
  const throttle = getThrottleKV();
  const ratioStr = throttle.Rainbow_Rate || '3:1';
  const ratio = parseRatio(ratioStr);
  const conversionRatio = ratio ? ratio.num : 3;
  
  const nameCol = headers.indexOf('PreferredName');
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const colorQtys = colorCols.map(col => coerceNumber(data[i][col], 0));
      const rainbowQty = rainbowCol !== -1 ? coerceNumber(data[i][rainbowCol], 0) : 0;

      const result = calculateUnlockEligibility_(colorQtys, rainbowQty, conversionRatio);

      // Build key summary
      const keySummary = colors.map((name, idx) => `${name}: ${colorQtys[idx]}`).join(', ');
      
      // Build unlock path explanation
      let howToUnlock = '';
      if (result.colorsMissing === 0) {
        howToUnlock = '🎯 All 5 colors collected! Ready to unlock!';
      } else if (result.canUnlockNow) {
        howToUnlock = `🎯 Use ${result.colorsMissing} Rainbow key(s) as wild cards!`;
      } else {
        // Need conversions
        const conversionsNeeded = result.colorsMissing - result.rainbowsAvailable;
        howToUnlock = `🎯 Convert ${conversionsNeeded * conversionRatio} excess keys → ${conversionsNeeded} Rainbow(s)`;
        
        // Show which colors have excess
        const excessColors = colors.filter((c, idx) => colorQtys[idx] > 1);
        if (excessColors.length > 0) {
          howToUnlock += `\n   (Can convert from: ${excessColors.join(', ')})`;
        }
      }

      const message = `🎉 ${preferredName} is ELIGIBLE TO UNLOCK! 🎉\n\n` +
                     `Current Keys: ${keySummary}\n` +
                     `Rainbow Keys: ${rainbowQty}\n\n` +
                     `${howToUnlock}\n\n` +
                     `Colors held: ${result.colorsHeld}/5 | ` +
                     `Wildcards available: ${result.totalWildcards} (${result.rainbowsAvailable} now + ${result.potentialRainbows} convertible)`;

      SpreadsheetApp.getUi().alert(message);
      return;
    }
  }
}

/**
 * Updates eligibility for all players in Key_Tracker
 * Call this after adding the "Able to Unlock?" column or to refresh all values
 */
function updateAllPlayersEligibility() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');

  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }

  const throttle = getThrottleKV();
  const ratioStr = throttle.Rainbow_Rate || '3:1';
  const ratio = parseRatio(ratioStr);

  if (!ratio || ratio.den !== 1) {
    throwError('Invalid ratio', 'INVALID_RATIO');
  }

  const conversionRatio = ratio.num;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');
  const unlockCol = headers.indexOf('Able to Unlock?');

  if (unlockCol === -1) {
    throwError('Able to Unlock? column not found', 'SCHEMA_INVALID');
  }

  let eligibleCount = 0;

  for (let i = 1; i < data.length; i++) {
    const colorQtys = colorCols.map(col => coerceNumber(data[i][col], 0));
    const rainbowQty = rainbowCol !== -1 ? coerceNumber(data[i][rainbowCol], 0) : 0;

    const result = calculateUnlockEligibility_(colorQtys, rainbowQty, conversionRatio);
    sheet.getRange(i + 1, unlockCol + 1).setValue(result.eligible);
    
    if (result.eligible) eligibleCount++;
  }

  logIntegrityAction('ELIGIBILITY_UPDATE', {
    details: `Updated eligibility for ${data.length - 1} players (${eligibleCount} eligible)`,
    status: 'SUCCESS'
  });
}

// ============================================================================
// KEY REPORTS AND UTILITIES
// ============================================================================

/**
 * Generates "Who's Closest?" report showing players ranked by proximity to unlocking
 * Shows what each player needs to win
 * @returns {string} Formatted report
 */
function whosClosestReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');

  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }

  const throttle = getThrottleKV();
  const ratioStr = throttle.Rainbow_Rate || '3:1';
  const ratio = parseRatio(ratioStr);

  if (!ratio || ratio.den !== 1) {
    throwError('Invalid ratio', 'INVALID_RATIO');
  }

  const conversionRatio = ratio.num;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');

  const playerAnalysis = [];

  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    if (!name) continue; // Skip empty rows
    
    const colorQtys = colorCols.map(col => coerceNumber(data[i][col], 0));
    const rainbowQty = rainbowCol !== -1 ? coerceNumber(data[i][rainbowCol], 0) : 0;

    const result = calculateUnlockEligibility_(colorQtys, rainbowQty, conversionRatio);

    // Calculate "distance" for sorting (lower = closer to winning)
    let distance;
    let whatTheyNeed;

    if (result.eligible) {
      if (result.canUnlockNow) {
        distance = 0;
        whatTheyNeed = '✅ CAN UNLOCK NOW!';
      } else {
        distance = 0.5;
        const conversionsNeeded = result.colorsMissing - result.rainbowsAvailable;
        whatTheyNeed = `✅ ELIGIBLE - Convert ${conversionsNeeded}x to unlock`;
      }
    } else {
      // Not eligible - calculate how many more keys needed
      const shortfall = result.colorsMissing - result.totalWildcards;
      distance = shortfall;
      
      // Figure out fastest path
      const missingColors = colors.filter((_, idx) => colorQtys[idx] === 0);
      if (missingColors.length <= shortfall) {
        whatTheyNeed = `Need: ${missingColors.join(', ')}`;
      } else {
        whatTheyNeed = `Need ${shortfall} more key(s) (any color or ${shortfall * conversionRatio} same-color to convert)`;
      }
    }

    playerAnalysis.push({
      name,
      colorQtys,
      rainbowQty,
      result,
      distance,
      whatTheyNeed
    });
  }

  // Sort by distance (closest first)
  playerAnalysis.sort((a, b) => a.distance - b.distance);

  // Build report
  let report = '=== WHO\'S CLOSEST TO UNLOCK? ===\n\n';
  report += `Unlock Rule: Cover all 5 colors (keys + Rainbow wildcards)\n`;
  report += `Conversion: ${conversionRatio} same-color keys → 1 Rainbow (no cap)\n\n`;

  if (playerAnalysis.length === 0) {
    report += 'No players in tracker yet.\n';
  } else {
    playerAnalysis.forEach((player, idx) => {
      const rank = idx + 1;
      const keySummary = colors.map((c, i) => `${c[0]}:${player.colorQtys[i]}`).join(' ');
      const rainbowStr = player.rainbowQty > 0 ? ` 🌈:${player.rainbowQty}` : '';

      report += `${rank}. ${player.name}\n`;
      report += `   [${keySummary}${rainbowStr}]\n`;
      report += `   ${player.whatTheyNeed}\n\n`;
    });
  }

  return report;
}

/**
 * Shows "Who's Closest?" report as a popup
 */
function showWhosClosest() {
  const report = whosClosestReport();
  SpreadsheetApp.getUi().alert('Who\'s Closest to Unlock?', report, SpreadsheetApp.getUi().ButtonSet.OK);

  logIntegrityAction('REPORT_GENERATED', {
    details: 'Who\'s Closest report viewed',
    status: 'SUCCESS'
  });
}

/**
 * Clears all keys for a SPECIFIC player (after unlock redemption)
 * @param {string} preferredName - Player name
 */
function clearPlayerKeys(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');

  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');
  const unlockCol = headers.indexOf('Able to Unlock?');
  const lastUpdatedCol = headers.indexOf('LastUpdated');

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      // Clear all color keys
      colorCols.forEach(col => {
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(0);
      });

      // Clear rainbow
      if (rainbowCol !== -1) sheet.getRange(i + 1, rainbowCol + 1).setValue(0);

      // Clear unlock eligibility
      if (unlockCol !== -1) sheet.getRange(i + 1, unlockCol + 1).setValue(0);

      // Update timestamp
      if (lastUpdatedCol !== -1) sheet.getRange(i + 1, lastUpdatedCol + 1).setValue(dateISO());

      logIntegrityAction('PLAYER_KEYS_CLEARED', {
        preferredName,
        details: 'All keys reset after unlock redemption',
        status: 'SUCCESS'
      });

      return true;
    }
  }

  throwError('Player not found', 'PLAYER_NOT_FOUND', `No entry for "${preferredName}"`);
}

/**
 * Clears all keys for all players (resets to zero)
 * Requires confirmation before executing
 */
function clearAllKeys() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Clear All Keys?',
    'This will reset ALL players\' keys to ZERO.\n\nThis action cannot be undone!\n\nAre you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Operation cancelled');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Key_Tracker');

  if (!sheet) {
    throwError('Key_Tracker not found', 'SHEET_MISSING');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  const colorCols = colors.map(c => headers.indexOf(c));
  const rainbowCol = headers.indexOf('RainbowEligible');
  const unlockCol = headers.indexOf('Able to Unlock?');
  const lastUpdatedCol = headers.indexOf('LastUpdated');

  for (let i = 1; i < data.length; i++) {
    colorCols.forEach(col => {
      if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(0);
    });

    if (rainbowCol !== -1) sheet.getRange(i + 1, rainbowCol + 1).setValue(0);
    if (unlockCol !== -1) sheet.getRange(i + 1, unlockCol + 1).setValue(0);
    if (lastUpdatedCol !== -1) sheet.getRange(i + 1, lastUpdatedCol + 1).setValue(dateISO());
  }

  logIntegrityAction('KEYS_CLEARED', {
    details: `All keys cleared for ${data.length - 1} players`,
    status: 'SUCCESS'
  });

  ui.alert('Keys Cleared', `All keys have been reset to zero for ${data.length - 1} players.`, ui.ButtonSet.OK);
}

// ============================================================================
// TEST CASES (for verification)
// ============================================================================

/**
 * Test function to verify unlock eligibility logic
 * Run this to confirm the rules work correctly
 */
function testUnlockEligibility() {
  const conversionRatio = 3;
  
  const testCases = [
    { colors: [1, 1, 1, 1, 1], rainbow: 0, expected: 1, desc: '(1,1,1,1,1) - all 5 colors' },
    { colors: [1, 1, 1, 4, 0], rainbow: 0, expected: 1, desc: '(1,1,1,4,0) - 4 colors + convert 3 excess' },
    { colors: [1, 1, 7, 0, 0], rainbow: 0, expected: 1, desc: '(1,1,7,0,0) - 3 colors + convert 6 excess → 2 rainbow' },
    { colors: [1, 10, 0, 0, 0], rainbow: 0, expected: 1, desc: '(1,10,0,0,0) - 2 colors + convert 9 excess → 3 rainbow' },
    { colors: [13, 0, 0, 0, 0], rainbow: 0, expected: 1, desc: '(13,0,0,0,0) - 1 color + convert 12 excess → 4 rainbow' },
    { colors: [0, 0, 0, 0, 5], rainbow: 0, expected: 0, desc: '(0,0,0,0,5) - only 1 color, need 4 rainbow but only 1 convertible' },
    { colors: [2, 2, 2, 2, 0], rainbow: 0, expected: 0, desc: '(2,2,2,2,0) - 4 colors, only 4 excess (can make 1 rainbow, need 1)' },
    { colors: [2, 2, 2, 2, 0], rainbow: 0, expected: 0, desc: '(2,2,2,2,0) - 4 colors, 4 excess keys, floor(4/3)=1 rainbow' },
    { colors: [3, 2, 2, 2, 0], rainbow: 0, expected: 1, desc: '(3,2,2,2,0) - 4 colors + 5 excess → 1 rainbow covers 1 missing' },
    { colors: [1, 1, 1, 1, 0], rainbow: 1, expected: 1, desc: '(1,1,1,1,0) + 1 rainbow - rainbow covers missing' },
    { colors: [4, 4, 0, 0, 0], rainbow: 0, expected: 1, desc: '(4,4,0,0,0) - 2 colors + 6 excess → 2 rainbow (need 3 total)' },
    { colors: [4, 4, 0, 0, 0], rainbow: 1, expected: 1, desc: '(4,4,0,0,0) + 1 rainbow - 2 colors + 2 convertible + 1 existing = 3 wildcards' },
  ];
  
  let passed = 0;
  let failed = 0;
  let report = '=== UNLOCK ELIGIBILITY TESTS ===\n\n';
  
  testCases.forEach((test, idx) => {
    const result = calculateUnlockEligibility_(test.colors, test.rainbow, conversionRatio);
    const status = result.eligible === test.expected ? '✅ PASS' : '❌ FAIL';
    
    if (result.eligible === test.expected) {
      passed++;
    } else {
      failed++;
    }
    
    report += `${idx + 1}. ${status}: ${test.desc}\n`;
    report += `   Result: eligible=${result.eligible}, colorsHeld=${result.colorsHeld}, `;
    report += `wildcards=${result.totalWildcards} (${result.rainbowsAvailable}+${result.potentialRainbows})\n`;
    report += `   ${result.details}\n\n`;
  });
  
  report += `\n=== SUMMARY: ${passed}/${testCases.length} passed ===\n`;
  
  Logger.log(report);
  SpreadsheetApp.getUi().alert('Test Results', report, SpreadsheetApp.getUi().ButtonSet.OK);
  
  return { passed, failed, total: testCases.length };
}