/**
 * Player Lookup Service - Main service for Player Lookup Sidebar
 * @fileoverview Provides searchCustomers and getPlayerLookupProfile functions
 */

// ============================================================================
// SEARCH CUSTOMERS
// ============================================================================

/**
 * Searches for customers matching a query string
 * @param {string} query - Search query (name, email, partial)
 * @return {CustomerSearchResult[]} Array of matching customers
 */
function searchCustomers(query) {
  if (!query || query.trim().length < 2) {
    return [];
  }

  const normalizedQuery = query.trim().toLowerCase();
  const results = [];
  const seenNames = new Set();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Search Key_Tracker
  try {
    const keySheet = ss.getSheetByName('Key_Tracker');
    if (keySheet && keySheet.getLastRow() > 1) {
      const keyData = keySheet.getDataRange().getValues();
      const headers = keyData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < keyData.length; i++) {
          const name = String(keyData[i][nameCol] || '').trim();
          if (!name || seenNames.has(name.toLowerCase())) continue;

          const score = computeMatchScore_(name, normalizedQuery);
          if (score > 0) {
            seenNames.add(name.toLowerCase());
            const keyCount = sumKeyCount_(keyData[i], headers);
            results.push({
              name: name,
              preferredName: name,
              email: '',
              source: 'Key_Tracker',
              hasPreorders: false,
              bp: 0,
              keyCount: keyCount,
              score: score
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching Key_Tracker: ' + e.message);
  }

  // Search BP_Total
  try {
    const bpSheet = ss.getSheetByName('BP_Total');
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const bpData = bpSheet.getDataRange().getValues();
      const headers = bpData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const bpCol = findColumnIndex_(headers, ['BP_Current', 'BP', 'Bonus_Points']);

      if (nameCol !== -1) {
        for (let i = 1; i < bpData.length; i++) {
          const name = String(bpData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const score = computeMatchScore_(name, normalizedQuery);

          if (score > 0) {
            const bp = bpCol !== -1 ? coerceNumber(bpData[i][bpCol], 0) : 0;

            if (seenNames.has(nameLower)) {
              // Update existing result with BP info
              const existing = results.find(r => r.name.toLowerCase() === nameLower);
              if (existing) {
                existing.bp = bp;
                if (existing.source !== 'BP_Total') {
                  existing.source += ', BP_Total';
                }
              }
            } else {
              seenNames.add(nameLower);
              results.push({
                name: name,
                preferredName: name,
                email: '',
                source: 'BP_Total',
                hasPreorders: false,
                bp: bp,
                keyCount: 0,
                score: score
              });
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching BP_Total: ' + e.message);
  }

  // Search Attendance_Missions
  try {
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const attendData = attendSheet.getDataRange().getValues();
      const headers = attendData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < attendData.length; i++) {
          const name = String(attendData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const score = computeMatchScore_(name, normalizedQuery);

          if (score > 0 && !seenNames.has(nameLower)) {
            seenNames.add(nameLower);
            results.push({
              name: name,
              preferredName: name,
              email: '',
              source: 'Attendance',
              hasPreorders: false,
              bp: 0,
              keyCount: 0,
              score: score
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching Attendance_Missions: ' + e.message);
  }

  // Search PreferredNames sheet if it exists
  try {
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (prefSheet && prefSheet.getLastRow() > 1) {
      const prefData = prefSheet.getDataRange().getValues();
      const headers = prefData[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const emailCol = findColumnIndex_(headers, ['Email', 'email']);
      const aliasCol = findColumnIndex_(headers, ['Alt_Names', 'Aliases', 'AltNames']);

      if (nameCol !== -1) {
        for (let i = 1; i < prefData.length; i++) {
          const name = String(prefData[i][nameCol] || '').trim();
          if (!name) continue;

          const nameLower = name.toLowerCase();
          const email = emailCol !== -1 ? String(prefData[i][emailCol] || '') : '';
          const aliases = aliasCol !== -1 ? String(prefData[i][aliasCol] || '') : '';

          // Score based on name, email, or aliases
          let score = computeMatchScore_(name, normalizedQuery);
          if (email && email.toLowerCase().includes(normalizedQuery)) {
            score = Math.max(score, 80);
          }
          if (aliases) {
            const aliasList = aliases.split(',').map(a => a.trim());
            for (const alias of aliasList) {
              const aliasScore = computeMatchScore_(alias, normalizedQuery);
              score = Math.max(score, aliasScore * 0.9); // Slight penalty for alias match
            }
          }

          if (score > 0) {
            if (seenNames.has(nameLower)) {
              // Update existing with PreferredNames source (higher priority)
              const existing = results.find(r => r.name.toLowerCase() === nameLower);
              if (existing) {
                existing.source = 'PreferredNames';
                existing.email = email || existing.email;
                existing.score = Math.max(existing.score, score + 10); // Bonus for PreferredNames
              }
            } else {
              seenNames.add(nameLower);
              results.push({
                name: name,
                preferredName: name,
                email: email,
                source: 'PreferredNames',
                hasPreorders: false,
                bp: 0,
                keyCount: 0,
                score: score + 10 // Bonus for PreferredNames
              });
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error searching PreferredNames: ' + e.message);
  }

  // Sort by score descending, then by name
  results.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return a.name.localeCompare(b.name);
  });

  // Limit to top 20 results
  return results.slice(0, 20);
}

/**
 * Computes a match score between a name and query
 * @param {string} name - Name to match
 * @param {string} query - Normalized query (lowercase)
 * @return {number} Score (0 = no match, 100 = exact match)
 * @private
 */
function computeMatchScore_(name, query) {
  const nameLower = name.toLowerCase();

  // Exact match
  if (nameLower === query) return 100;

  // Starts with query
  if (nameLower.startsWith(query)) return 90;

  // Contains query as a word
  const words = nameLower.split(/\s+/);
  for (const word of words) {
    if (word === query) return 85;
    if (word.startsWith(query)) return 75;
  }

  // Contains query anywhere
  if (nameLower.includes(query)) return 60;

  // Query contains name (partial name typed)
  if (query.includes(nameLower)) return 40;

  return 0;
}

/**
 * Finds column index from possible header names
 * @param {Array} headers - Header row
 * @param {string[]} possibleNames - Possible column names
 * @return {number} Column index or -1
 * @private
 */
function findColumnIndex_(headers, possibleNames) {
  for (const name of possibleNames) {
    const idx = headers.indexOf(name);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * Sums key count from a row
 * @param {Array} row - Data row
 * @param {Array} headers - Header row
 * @return {number} Total key count
 * @private
 */
function sumKeyCount_(row, headers) {
  const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
  let total = 0;

  for (const color of colors) {
    const col = headers.indexOf(color);
    if (col !== -1) {
      total += coerceNumber(row[col], 0);
    }
  }

  return total;
}

// ============================================================================
// GET PLAYER LOOKUP PROFILE
// ============================================================================

/**
 * Main entrypoint for Player Lookup sidebar.
 * @param {string} query - What staff typed (name, email, partial).
 * @return {PlayerLookupProfile}
 */
function getPlayerLookupProfile(query) {
  try {
    // Validate input
    if (!query || query.trim().length < 2) {
      return createErrorResponse_(query || '', 'Please type at least 2 characters.');
    }

    const trimmedQuery = query.trim();

    // Search for matching players
    const matches = searchCustomers(trimmedQuery);

    if (matches.length === 0) {
      return {
        status: 'NOT_FOUND',
        query: trimmedQuery,
        message: 'No player found matching "' + trimmedQuery + '".'
      };
    }

    // Resolve canonical name
    const canonicalName = resolveCanonicalName_(matches, trimmedQuery);

    // Build full profile
    return buildPlayerProfile_(canonicalName, trimmedQuery);

  } catch (e) {
    Logger.log('Error in getPlayerLookupProfile: ' + e.message);
    return createErrorResponse_(query || '', 'An error occurred: ' + e.message);
  }
}

/**
 * Creates an error response
 * @param {string} query - Original query
 * @param {string} message - Error message
 * @return {PlayerLookupProfile}
 * @private
 */
function createErrorResponse_(query, message) {
  return {
    status: 'ERROR',
    query: query,
    message: message,
    meta: {
      lastUpdated: dateISO(),
      source: 'v7.9.7 PlayerLookup',
      errors: [message]
    }
  };
}

/**
 * Resolves the best canonical name from matches
 * @param {CustomerSearchResult[]} matches - Search results
 * @param {string} query - Original query
 * @return {string} Canonical name
 * @private
 */
function resolveCanonicalName_(matches, query) {
  const normalizedQuery = query.toLowerCase();

  // Exact match (case-insensitive)
  for (const match of matches) {
    if (match.name.toLowerCase() === normalizedQuery) {
      return match.preferredName || match.name;
    }
  }

  // Prefer PreferredNames source
  const prefNameMatch = matches.find(m => m.source.includes('PreferredNames'));
  if (prefNameMatch) {
    return prefNameMatch.preferredName || prefNameMatch.name;
  }

  // Starts with query
  const startsWithMatch = matches.find(m => m.name.toLowerCase().startsWith(normalizedQuery));
  if (startsWithMatch) {
    return startsWithMatch.preferredName || startsWithMatch.name;
  }

  // Return highest scoring match
  return matches[0].preferredName || matches[0].name;
}

/**
 * Builds a complete player profile
 * @param {string} canonicalName - Resolved player name
 * @param {string} query - Original query
 * @return {PlayerLookupProfile}
 * @private
 */
function buildPlayerProfile_(canonicalName, query) {
  const errors = [];

  // Get identity info
  const identity = getIdentityInfo_(canonicalName, errors);

  // Get BP info
  const bonusPoints = getBonusPointsInfo_(canonicalName, errors);

  // Get keys info
  const keys = getKeysInfo_(canonicalName, errors);

  // Get store credit info
  const storeCredit = getStoreCreditInfo_(canonicalName, errors);

  // Get preorders info
  const preorders = getPreordersInfo_(canonicalName, errors);

  // Get attendance info
  const attendance = getAttendanceInfo_(canonicalName, errors);

  // Get missions info
  const missions = getMissionsInfo_(canonicalName, errors);

  // Get flags info
  const flags = getFlagsInfo_(canonicalName, errors);

  // Build summary line
  const summary = buildSummaryLine_(identity, bonusPoints, keys, storeCredit, attendance);

  return {
    status: 'OK',
    query: query,
    playerId: canonicalName,
    preferredName: canonicalName,
    displayName: identity.displayName || canonicalName,
    aliases: identity.aliases || [],

    contact: identity.contact,
    membership: identity.membership,
    summary: summary,
    bonusPoints: bonusPoints,
    keys: keys,
    storeCredit: storeCredit,
    preorders: preorders,
    attendance: attendance,
    missions: missions,
    leagues: [],
    flags: flags,

    meta: {
      lastUpdated: dateISO(),
      source: 'v7.9.7 PlayerLookup',
      errors: errors.length > 0 ? errors : undefined
    }
  };
}

/**
 * Gets identity and contact info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array to append to
 * @return {Object} Identity info
 * @private
 */
function getIdentityInfo_(name, errors) {
  const result = {
    displayName: name,
    aliases: [],
    contact: {
      email: '',
      phone: '',
      discord: ''
    },
    membership: {
      isMember: false,
      tier: null,
      joinedAt: null,
      expiresAt: null
    }
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try PreferredNames sheet
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (prefSheet && prefSheet.getLastRow() > 1) {
      const data = prefSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const emailCol = findColumnIndex_(headers, ['Email', 'email']);
      const phoneCol = findColumnIndex_(headers, ['Phone', 'phone']);
      const discordCol = findColumnIndex_(headers, ['Discord', 'discord']);
      const aliasCol = findColumnIndex_(headers, ['Alt_Names', 'Aliases', 'AltNames']);
      const displayCol = findColumnIndex_(headers, ['DisplayName', 'Display_Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (emailCol !== -1) result.contact.email = String(data[i][emailCol] || '');
            if (phoneCol !== -1) result.contact.phone = String(data[i][phoneCol] || '');
            if (discordCol !== -1) result.contact.discord = String(data[i][discordCol] || '');
            if (displayCol !== -1 && data[i][displayCol]) {
              result.displayName = String(data[i][displayCol]);
            }
            if (aliasCol !== -1 && data[i][aliasCol]) {
              result.aliases = String(data[i][aliasCol]).split(',').map(a => a.trim()).filter(a => a);
            }
            break;
          }
        }
      }
    }

    // Try Members sheet for membership info
    const membersSheet = ss.getSheetByName('Members');
    if (membersSheet && membersSheet.getLastRow() > 1) {
      const data = membersSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const tierCol = findColumnIndex_(headers, ['Tier', 'tier', 'Level']);
      const joinedCol = findColumnIndex_(headers, ['JoinedAt', 'Joined', 'Start_Date']);
      const expiresCol = findColumnIndex_(headers, ['ExpiresAt', 'Expires', 'End_Date']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.membership.isMember = true;
            if (tierCol !== -1) result.membership.tier = String(data[i][tierCol] || '');
            if (joinedCol !== -1 && data[i][joinedCol]) {
              result.membership.joinedAt = formatDateSafe_(data[i][joinedCol]);
            }
            if (expiresCol !== -1 && data[i][expiresCol]) {
              result.membership.expiresAt = formatDateSafe_(data[i][expiresCol]);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Identity lookup error: ' + e.message);
    Logger.log('Identity lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets bonus points info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {BonusPointsInfo}
 * @private
 */
function getBonusPointsInfo_(name, errors) {
  const result = {
    current: 0,
    lifetimeEarned: 0,
    lifetimeSpent: 0,
    pending: 0,
    lastEarnedFrom: 'Unknown',
    lastUpdated: ''
  };

  try {
    // Use existing getPlayerBP function
    result.current = getPlayerBP(name);

    // Try to get additional BP info from BP_Total
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const bpSheet = ss.getSheetByName('BP_Total');
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const data = bpSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const histCol = findColumnIndex_(headers, ['Historical_BP', 'Lifetime_Earned', 'Total_Earned']);
      const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (histCol !== -1) {
              result.lifetimeEarned = coerceNumber(data[i][histCol], result.current);
            } else {
              result.lifetimeEarned = result.current;
            }
            if (updatedCol !== -1 && data[i][updatedCol]) {
              result.lastUpdated = formatDateSafe_(data[i][updatedCol]);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('BP lookup error: ' + e.message);
    Logger.log('BP lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets keys info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {KeyInfo}
 * @private
 */
function getKeysInfo_(name, errors) {
  const result = {
    Red: 0,
    Blue: 0,
    Green: 0,
    Yellow: 0,
    Purple: 0,
    RainbowEligible: 0,
    lifetimeUnlocked: 0,
    lastUnlockDate: ''
  };

  try {
    // Use existing getPlayerKeys function
    const keys = getPlayerKeys(name);
    if (keys) {
      result.Red = coerceNumber(keys.Red, 0);
      result.Blue = coerceNumber(keys.Blue, 0);
      result.Green = coerceNumber(keys.Green, 0);
      result.Yellow = coerceNumber(keys.Yellow, 0);
      result.Purple = coerceNumber(keys.Purple, 0);
      result.RainbowEligible = coerceNumber(keys.RainbowEligible, 0);

      if (keys.LastUpdated) {
        result.lastUnlockDate = formatDateSafe_(keys.LastUpdated);
      }
    }
  } catch (e) {
    errors.push('Keys lookup error: ' + e.message);
    Logger.log('Keys lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets store credit info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {StoreCreditInfo}
 * @private
 */
function getStoreCreditInfo_(name, errors) {
  const result = {
    balance: 0,
    currency: 'USD',
    lastUpdated: '',
    lastTxSummary: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try Store_Credit sheet
    const scSheet = ss.getSheetByName('Store_Credit');
    if (scSheet && scSheet.getLastRow() > 1) {
      const data = scSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const balanceCol = findColumnIndex_(headers, ['Balance', 'Credit', 'Amount']);
      const updatedCol = findColumnIndex_(headers, ['LastUpdated', 'Last_Updated']);

      if (nameCol !== -1 && balanceCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.balance = coerceNumber(data[i][balanceCol], 0);
            if (updatedCol !== -1 && data[i][updatedCol]) {
              result.lastUpdated = formatDateSafe_(data[i][updatedCol]);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Store credit lookup error: ' + e.message);
    Logger.log('Store credit lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets preorders info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {PreordersInfo}
 * @private
 */
function getPreordersInfo_(name, errors) {
  const result = {
    active: [],
    historyCount: 0
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try Preorders sheet
    const preorderSheet = ss.getSheetByName('Preorders');
    if (preorderSheet && preorderSheet.getLastRow() > 1) {
      const data = preorderSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name', 'Customer']);
      const itemCol = findColumnIndex_(headers, ['Item', 'Item_Name', 'Product']);
      const qtyCol = findColumnIndex_(headers, ['Qty', 'Quantity']);
      const statusCol = findColumnIndex_(headers, ['Status', 'status']);
      const setCol = findColumnIndex_(headers, ['Set', 'SetName', 'Set_Name']);
      const priceCol = findColumnIndex_(headers, ['Price', 'Unit_Price']);
      const paidCol = findColumnIndex_(headers, ['Paid', 'Paid_Amount']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.historyCount++;

            const status = statusCol !== -1 ? String(data[i][statusCol] || 'Active') : 'Active';
            const isActive = !['Completed', 'Cancelled', 'Picked Up'].includes(status);

            if (isActive) {
              const unitPrice = priceCol !== -1 ? coerceNumber(data[i][priceCol], 0) : 0;
              const qty = qtyCol !== -1 ? coerceNumber(data[i][qtyCol], 1) : 1;
              const paid = paidCol !== -1 ? coerceNumber(data[i][paidCol], 0) : 0;

              result.active.push({
                rowIndex: i + 1,
                setName: setCol !== -1 ? String(data[i][setCol] || '') : '',
                itemName: itemCol !== -1 ? String(data[i][itemCol] || '') : '',
                qty: qty,
                unitPrice: unitPrice,
                paidAmount: paid,
                balanceDue: (unitPrice * qty) - paid,
                status: status
              });
            }
          }
        }
      }
    }
  } catch (e) {
    errors.push('Preorders lookup error: ' + e.message);
    Logger.log('Preorders lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets attendance info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {AttendanceInfo}
 * @private
 */
function getAttendanceInfo_(name, errors) {
  const result = {
    totalEvents: 0,
    eventsLast90Days: 0,
    lastEventDate: '',
    recentEvents: [],
    primaryFormat: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Search event tabs for player attendance
    const eventTabs = listEventTabs();
    const now = new Date();
    const ninetyDaysAgo = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
    const recentEvents = [];

    for (const tabName of eventTabs.slice(-50)) { // Check last 50 events max
      try {
        const sheet = ss.getSheetByName(tabName);
        if (!sheet || sheet.getLastRow() <= 1) continue;

        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

        if (nameCol === -1) continue;

        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            result.totalEvents++;

            // Parse event date from tab name (MM-DD-YYYY)
            const dateMatch = tabName.match(/^(\d{2})-(\d{2})-(\d{4})/);
            if (dateMatch) {
              const eventDate = new Date(dateMatch[3], dateMatch[1] - 1, dateMatch[2]);

              if (eventDate >= ninetyDaysAgo) {
                result.eventsLast90Days++;
              }

              if (!result.lastEventDate || eventDate > new Date(result.lastEventDate)) {
                result.lastEventDate = formatDateSafe_(eventDate);
              }

              // Get event type from metadata
              const props = getEventProps(sheet);
              const eventType = props.event_type || 'Event';

              recentEvents.push({
                date: tabName,
                type: eventType
              });
            }
            break;
          }
        }
      } catch (e) {
        // Skip problematic tabs
      }
    }

    // Get last 5 recent events
    result.recentEvents = recentEvents
      .slice(-5)
      .reverse()
      .map(e => e.date + ' (' + e.type + ')');

    // Try Attendance_Missions for additional info
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const data = attendSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            // If we have total from missions sheet, use it if higher
            const cTotalCol = headers.indexOf('C_Total');
            if (cTotalCol !== -1) {
              const missionTotal = coerceNumber(data[i][cTotalCol], 0);
              result.totalEvents = Math.max(result.totalEvents, missionTotal);
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Attendance lookup error: ' + e.message);
    Logger.log('Attendance lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets missions info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {MissionsInfo}
 * @private
 */
function getMissionsInfo_(name, errors) {
  const result = {
    completed: 0,
    inProgress: 0,
    badges: [],
    recentAwards: []
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try Attendance_Missions for badges
    const attendSheet = ss.getSheetByName('Attendance_Missions');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const data = attendSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const badgesCol = findColumnIndex_(headers, ['Badges', 'badges']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (badgesCol !== -1 && data[i][badgesCol]) {
              result.badges = String(data[i][badgesCol]).split(',').map(b => b.trim()).filter(b => b);
              result.completed = result.badges.length;
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Missions lookup error: ' + e.message);
    Logger.log('Missions lookup error: ' + e.message);
  }

  return result;
}

/**
 * Gets flags and notes info
 * @param {string} name - Player name
 * @param {string[]} errors - Error array
 * @return {FlagsInfo}
 * @private
 */
function getFlagsInfo_(name, errors) {
  const result = {
    isBanned: false,
    warningsCount: 0,
    tags: [],
    internalNote: ''
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Try Player_Notes sheet
    const notesSheet = ss.getSheetByName('Player_Notes');
    if (notesSheet && notesSheet.getLastRow() > 1) {
      const data = notesSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = findColumnIndex_(headers, ['PreferredName', 'Preferred_Name', 'Name']);
      const bannedCol = findColumnIndex_(headers, ['IsBanned', 'Banned', 'is_banned']);
      const warningsCol = findColumnIndex_(headers, ['Warnings', 'WarningCount', 'warnings_count']);
      const tagsCol = findColumnIndex_(headers, ['Tags', 'tags']);
      const noteCol = findColumnIndex_(headers, ['Note', 'Notes', 'InternalNote']);

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][nameCol]).toLowerCase() === name.toLowerCase()) {
            if (bannedCol !== -1) result.isBanned = coerceBoolean(data[i][bannedCol]);
            if (warningsCol !== -1) result.warningsCount = coerceNumber(data[i][warningsCol], 0);
            if (tagsCol !== -1 && data[i][tagsCol]) {
              result.tags = String(data[i][tagsCol]).split(',').map(t => t.trim()).filter(t => t);
            }
            if (noteCol !== -1) result.internalNote = String(data[i][noteCol] || '');
            break;
          }
        }
      }
    }
  } catch (e) {
    errors.push('Flags lookup error: ' + e.message);
    Logger.log('Flags lookup error: ' + e.message);
  }

  return result;
}

/**
 * Builds a summary line from profile data
 * @param {Object} identity - Identity info
 * @param {BonusPointsInfo} bonusPoints - BP info
 * @param {KeyInfo} keys - Keys info
 * @param {StoreCreditInfo} storeCredit - Store credit info
 * @param {AttendanceInfo} attendance - Attendance info
 * @return {string} Summary line
 * @private
 */
function buildSummaryLine_(identity, bonusPoints, keys, storeCredit, attendance) {
  const parts = [];

  // Membership
  if (identity.membership.isMember) {
    const tier = identity.membership.tier ? identity.membership.tier + ' ' : '';
    parts.push(tier + 'Member');
  }

  // Events
  if (attendance.totalEvents > 0) {
    parts.push(attendance.totalEvents + ' events');
  }

  // Keys
  const totalKeys = keys.Red + keys.Blue + keys.Green + keys.Yellow + keys.Purple;
  if (totalKeys > 0 || keys.RainbowEligible > 0) {
    const keyText = keys.RainbowEligible > 0
      ? keys.RainbowEligible + ' rainbow'
      : totalKeys + ' keys';
    parts.push(keyText);
  }

  // Store credit
  if (storeCredit.balance > 0) {
    parts.push(formatCurrency(storeCredit.balance) + ' credit');
  }

  // BP
  if (bonusPoints.current > 0) {
    parts.push(bonusPoints.current + ' BP');
  }

  if (parts.length === 0) {
    return 'New player';
  }

  return parts.join(' • ');
}

/**
 * Safely formats a date value
 * @param {*} value - Date value
 * @return {string} ISO string or empty string
 * @private
 */
function formatDateSafe_(value) {
  if (!value) return '';

  try {
    if (value instanceof Date) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    if (typeof value === 'string') {
      // Already a string, try to parse and reformat
      const date = new Date(value);
      if (!isNaN(date.getTime())) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
      return value;
    }
    return String(value);
  } catch (e) {
    return String(value || '');
  }
}