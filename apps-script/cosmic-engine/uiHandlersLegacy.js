/**
 * Legacy Player Profile Functions
 * @fileoverview Deprecated player profile functions kept for backwards compatibility
 *
 * Version: 7.9.7
 *
 * WARNING: These functions are DEPRECATED.
 * Use getPlayerLookupProfile() from ui.handler.playerLookup.gs instead.
 *
 * These functions are preserved only for backwards compatibility with any
 * existing code that may depend on the old getPlayerProfile_() interface.
 */

// ============================================================================
// LEGACY PLAYER PROFILE - FULL IMPLEMENTATION (DEPRECATED)
// ============================================================================

/**
 * Returns a complete PlayerProfile object for a given player search value.
 * This powers the Player Lookup sidebar and future Player Dashboard sheet.
 *
 * @deprecated Use getPlayerLookupProfile() instead
 * @param {string} preferredNameOrId - what staff typed in the lookup box
 * @return {Object} PlayerProfile
 */
function getPlayerProfile_(preferredNameOrId) {
  try {
    // Validate input
    if (!preferredNameOrId || typeof preferredNameOrId !== 'string') {
      return {
        found: false,
        status: 'INVALID_INPUT',
        message: 'Please enter a player name to search.'
      };
    }

    const searchValue = preferredNameOrId.trim();
    if (!searchValue) {
      return {
        found: false,
        status: 'INVALID_INPUT',
        message: 'Please enter a player name to search.'
      };
    }

    // Find player in PreferredNames (or fallback to Key_Tracker/BP_Total)
    const playerResult = findPlayerInPreferredNames_(searchValue);

    if (!playerResult) {
      return {
        found: false,
        status: 'NOT_FOUND',
        message: `No player found for "${searchValue}".`,
        PreferredName: null,
        Alt_Names: null,
        Email: null,
        Created: null,
        Updated: null,
        Roll_Profile: null,
        Notes: null,
        Store_Credit: 0,
        Store_Credit_LastSync: null,
        Keys: { total: 0, display: '0 keys' },
        Current_BP: 0,
        Historical_BP: 0,
        Preorders_Open: 0,
        ThisNeeds_Open: 0,
        Missions_Completed: 0,
        Awards_Count: 0,
        Last_Visit: null,
        Lifetime_Visits: 0,
        Primary_Format: null
      };
    }

    // Extract base identity fields from the found row
    const preferredName = playerResult.PreferredName;

    // Load all profile components
    const bpProfile = loadBPProfileForPlayer_(preferredName);
    const keysProfile = loadKeysProfileForPlayer_(preferredName);
    const storeCreditProfile = loadStoreCreditProfileForPlayer_(preferredName);
    const missionsAwards = loadMissionsAndAwardsForPlayer_(preferredName);
    const attendance = loadAttendanceProfileForPlayer_(preferredName);
    const queues = loadQueuesProfileForPlayer_(preferredName);

    // Construct the full PlayerProfile
    return {
      found: true,
      status: 'OK',
      message: null,

      // 1) Identity & Profile
      PreferredName: preferredName,
      Alt_Names: playerResult.Alt_Names || null,
      Email: playerResult.Email || null,
      Created: playerResult.Created || null,
      Updated: playerResult.Updated || null,
      Roll_Profile: playerResult.Roll_Profile || playerResult.Notes || null,
      Notes: playerResult.Notes || null,

      // 2) Economy / Rewards
      Store_Credit: storeCreditProfile.balance,
      Store_Credit_LastSync: storeCreditProfile.lastSync,
      Keys: keysProfile,
      Current_BP: bpProfile.currentBP,
      Historical_BP: bpProfile.historicalBP,

      // 3) Open Queues
      Preorders_Open: queues.preordersOpen,
      ThisNeeds_Open: queues.thisNeedsOpen,

      // 4) Missions & Awards
      Missions_Completed: missionsAwards.missionsCompleted,
      Awards_Count: missionsAwards.awardsCount,

      // 5) Attendance
      Last_Visit: attendance.lastVisit,
      Lifetime_Visits: attendance.lifetimeVisits,
      Primary_Format: attendance.primaryFormat
    };

  } catch (e) {
    Logger.log('getPlayerProfile_ error: ' + e.message);
    return {
      found: false,
      status: 'ERROR',
      message: 'An error occurred: ' + e.message,
      PreferredName: null,
      Alt_Names: null,
      Email: null,
      Created: null,
      Updated: null,
      Roll_Profile: null,
      Notes: null,
      Store_Credit: 0,
      Store_Credit_LastSync: null,
      Keys: { total: 0, display: '0 keys' },
      Current_BP: 0,
      Historical_BP: 0,
      Preorders_Open: 0,
      ThisNeeds_Open: 0,
      Missions_Completed: 0,
      Awards_Count: 0,
      Last_Visit: null,
      Lifetime_Visits: 0,
      Primary_Format: null
    };
  }
}

/**
 * Finds a player row in PreferredNames by preferred name or alt name.
 * Falls back to Key_Tracker and BP_Total if PreferredNames doesn't exist.
 *
 * Matching rules:
 *   - First try exact match on PreferredName (case-insensitive).
 *   - Then search Alt_Names (split on comma/pipe, case-insensitive).
 *
 * @param {string} searchValue - Search term
 * @return {Object|null} Player data object or null if not found
 * @private
 */
function findPlayerInPreferredNames_(searchValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const normalizedSearch = searchValue.toLowerCase().trim();

  // Try PreferredNames sheet first (canonical source)
  const prefNamesSheet = ss.getSheetByName('PreferredNames');
  if (prefNamesSheet && prefNamesSheet.getLastRow() > 1) {
    const data = prefNamesSheet.getDataRange().getValues();
    const headers = data[0];

    // Find column indices
    const prefNameCol = headers.indexOf('PreferredName');
    const altNamesCol = headers.indexOf('Alt_Names');
    const emailCol = Math.max(headers.indexOf('Email'), headers.indexOf('email'));
    const createdCol = headers.indexOf('Created');
    const updatedCol = headers.indexOf('Updated');
    const notesCol = headers.indexOf('Notes');
    const rollProfileCol = headers.indexOf('Roll_Profile');

    if (prefNameCol !== -1) {
      // First pass: exact match on PreferredName
      for (let i = 1; i < data.length; i++) {
        const rowPrefName = String(data[i][prefNameCol] || '').trim();
        if (rowPrefName.toLowerCase() === normalizedSearch) {
          return extractPlayerRow_(data[i], headers, {
            prefNameCol, altNamesCol, emailCol, createdCol, updatedCol, notesCol, rollProfileCol
          });
        }
      }

      // Second pass: search Alt_Names
      if (altNamesCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          const altNames = String(data[i][altNamesCol] || '');
          const aliases = altNames.split(/[,|]/).map(a => a.trim().toLowerCase());
          if (aliases.includes(normalizedSearch)) {
            return extractPlayerRow_(data[i], headers, {
              prefNameCol, altNamesCol, emailCol, createdCol, updatedCol, notesCol, rollProfileCol
            });
          }
        }
      }

      // Third pass: partial match on PreferredName
      for (let i = 1; i < data.length; i++) {
        const rowPrefName = String(data[i][prefNameCol] || '').trim().toLowerCase();
        if (rowPrefName.includes(normalizedSearch) || normalizedSearch.includes(rowPrefName)) {
          return extractPlayerRow_(data[i], headers, {
            prefNameCol, altNamesCol, emailCol, createdCol, updatedCol, notesCol, rollProfileCol
          });
        }
      }
    }
  }

  // Fallback: search in Key_Tracker and BP_Total (current canonical sources)
  const canonicalNames = getCanonicalNames();

  // Try exact match first
  for (const name of canonicalNames) {
    if (name.toLowerCase() === normalizedSearch) {
      return {
        PreferredName: name,
        Alt_Names: null,
        Email: null,
        Created: null,
        Updated: null,
        Notes: null,
        Roll_Profile: null
      };
    }
  }

  // Try partial match
  for (const name of canonicalNames) {
    if (name.toLowerCase().includes(normalizedSearch) || normalizedSearch.includes(name.toLowerCase())) {
      return {
        PreferredName: name,
        Alt_Names: null,
        Email: null,
        Created: null,
        Updated: null,
        Notes: null,
        Roll_Profile: null
      };
    }
  }

  return null;
}

/**
 * Extracts player data from a row in PreferredNames sheet
 * @param {Array} row - Row data
 * @param {Array} headers - Header row
 * @param {Object} cols - Column indices
 * @return {Object} Player data object
 * @private
 */
function extractPlayerRow_(row, headers, cols) {
  const formatDate = (val) => {
    if (!val) return null;
    if (val instanceof Date) {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    return String(val);
  };

  return {
    PreferredName: cols.prefNameCol !== -1 ? String(row[cols.prefNameCol] || '') : null,
    Alt_Names: cols.altNamesCol !== -1 ? String(row[cols.altNamesCol] || '') : null,
    Email: cols.emailCol !== -1 ? String(row[cols.emailCol] || '') : null,
    Created: cols.createdCol !== -1 ? formatDate(row[cols.createdCol]) : null,
    Updated: cols.updatedCol !== -1 ? formatDate(row[cols.updatedCol]) : null,
    Notes: cols.notesCol !== -1 ? String(row[cols.notesCol] || '') : null,
    Roll_Profile: cols.rollProfileCol !== -1 ? String(row[cols.rollProfileCol] || '') : null
  };
}

/**
 * Loads BP profile for a player
 * @param {string} preferredName - Player name
 * @return {Object} { currentBP, historicalBP }
 * @private
 */
function loadBPProfileForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let currentBP = 0;
  let historicalBP = 0;

  try {
    // Get current BP from BP_Total
    const bpSheet = ss.getSheetByName('BP_Total');
    if (bpSheet && bpSheet.getLastRow() > 1) {
      const data = bpSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf('PreferredName');
      const currentCol = headers.indexOf('BP_Current');
      const historicalCol = headers.indexOf('Historical_BP');

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][nameCol] === preferredName) {
            currentBP = currentCol !== -1 ? coerceNumber(data[i][currentCol], 0) : 0;
            historicalBP = historicalCol !== -1 ? coerceNumber(data[i][historicalCol], 0) : currentBP;
            break;
          }
        }
      }
    }

    // Check Prestige_Overflow for additional historical BP
    const prestigeSheet = ss.getSheetByName('Prestige_Overflow');
    if (prestigeSheet && prestigeSheet.getLastRow() > 1) {
      const data = prestigeSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf('PreferredName');
      const overflowCol = headers.indexOf('Total_Overflow');

      if (nameCol !== -1 && overflowCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][nameCol] === preferredName) {
            historicalBP += coerceNumber(data[i][overflowCol], 0);
            break;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('loadBPProfileForPlayer_ error: ' + e.message);
  }

  return { currentBP, historicalBP };
}

/**
 * Loads keys profile for a player
 * @param {string} preferredName - Player name
 * @return {Object} { total, display }
 * @private
 */
function loadKeysProfileForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let total = 0;
  let display = '0 keys';

  try {
    const keySheet = ss.getSheetByName('Key_Tracker');
    if (keySheet && keySheet.getLastRow() > 1) {
      const data = keySheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf('PreferredName');

      if (nameCol !== -1) {
        const colors = ['Red', 'Blue', 'Green', 'Yellow', 'Purple'];
        const colorCols = colors.map(c => headers.indexOf(c));
        const rainbowCol = headers.indexOf('RainbowEligible');

        for (let i = 1; i < data.length; i++) {
          if (data[i][nameCol] === preferredName) {
            // Count total keys and colors earned
            let colorCount = 0;
            let keyCount = 0;

            colorCols.forEach((col, idx) => {
              if (col !== -1) {
                const qty = coerceNumber(data[i][col], 0);
                keyCount += qty;
                if (qty > 0) colorCount++;
              }
            });

            // Add rainbow keys
            const rainbowKeys = rainbowCol !== -1 ? coerceNumber(data[i][rainbowCol], 0) : 0;

            total = keyCount + rainbowKeys;

            // Build display string
            if (colorCount === 5 && rainbowKeys > 0) {
              display = `${total} keys (Lockbox Complete, ${rainbowKeys} Rainbow)`;
            } else if (colorCount === 5) {
              display = `${total} keys (5/5 colors - Rainbow Eligible!)`;
            } else if (total > 0) {
              display = `${total} keys (${colorCount}/5 colors)`;
            } else {
              display = '0 keys';
            }

            break;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('loadKeysProfileForPlayer_ error: ' + e.message);
  }

  return { total, display };
}

/**
 * Loads store credit profile for a player
 * @param {string} preferredName - Player name
 * @return {Object} { balance, lastSync }
 * @private
 */
function loadStoreCreditProfileForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let balance = 0;
  let lastSync = null;

  try {
    // Try Store_Credit_Viewer first
    const viewerSheet = ss.getSheetByName('Store_Credit_Viewer');
    if (viewerSheet && viewerSheet.getLastRow() > 1) {
      const data = viewerSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = Math.max(headers.indexOf('PreferredName'), headers.indexOf('Name'), headers.indexOf('Customer'));
      const balanceCol = headers.indexOf('Balance');
      const syncCol = Math.max(headers.indexOf('Source_Timestamp'), headers.indexOf('LastSync'), headers.indexOf('Last_Updated'));

      if (nameCol !== -1 && balanceCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          const rowName = String(data[i][nameCol] || '').toLowerCase();
          if (rowName === preferredName.toLowerCase()) {
            balance = coerceNumber(data[i][balanceCol], 0);
            if (syncCol !== -1 && data[i][syncCol]) {
              const syncVal = data[i][syncCol];
              if (syncVal instanceof Date) {
                lastSync = Utilities.formatDate(syncVal, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
              } else {
                lastSync = String(syncVal);
              }
            }
            break;
          }
        }
      }
    }

    // Fallback: try Store_Credit_Ledger for most recent balance
    if (balance === 0) {
      const ledgerSheet = ss.getSheetByName('Store_Credit_Ledger');
      if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
        const data = ledgerSheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = Math.max(headers.indexOf('PreferredName'), headers.indexOf('Name'), headers.indexOf('Customer'));
        const balanceAfterCol = headers.indexOf('Balance_After');
        const timestampCol = headers.indexOf('Timestamp');

        if (nameCol !== -1 && balanceAfterCol !== -1) {
          // Find most recent transaction for this player
          for (let i = data.length - 1; i >= 1; i--) {
            const rowName = String(data[i][nameCol] || '').toLowerCase();
            if (rowName === preferredName.toLowerCase()) {
              balance = coerceNumber(data[i][balanceAfterCol], 0);
              if (timestampCol !== -1 && data[i][timestampCol]) {
                const ts = data[i][timestampCol];
                if (ts instanceof Date) {
                  lastSync = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
                } else {
                  lastSync = String(ts);
                }
              }
              break;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log('loadStoreCreditProfileForPlayer_ error: ' + e.message);
  }

  return { balance, lastSync };
}

/**
 * Loads missions and awards counts for a player
 * @param {string} preferredName - Player name
 * @return {Object} { missionsCompleted, awardsCount }
 * @private
 */
function loadMissionsAndAwardsForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let missionsCompleted = 0;
  let awardsCount = 0;

  try {
    // Count completed missions from Mission_Log_1 and Mission_Log_2
    const missionSheets = ['Mission_Log_1', 'Mission_Log_2', 'Mission_Log'];
    missionSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = Math.max(headers.indexOf('PreferredName'), headers.indexOf('Name'), headers.indexOf('Player'));
        const statusCol = headers.indexOf('Status');

        if (nameCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const rowName = String(data[i][nameCol] || '').toLowerCase();
            if (rowName === preferredName.toLowerCase()) {
              if (statusCol === -1) {
                // No status column - count all rows as missions
                missionsCompleted++;
              } else {
                const status = String(data[i][statusCol] || '').toLowerCase();
                if (status === 'complete' || status === 'completed' || status === 'claimed') {
                  missionsCompleted++;
                }
              }
            }
          }
        }
      }
    });

    // Check Attendance_Missions for badge/mission counts
    const attendMissionsSheet = ss.getSheetByName('Attendance_Missions');
    if (attendMissionsSheet && attendMissionsSheet.getLastRow() > 1) {
      const data = attendMissionsSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf('PreferredName');
      const badgesCol = headers.indexOf('Badges');

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][nameCol] === preferredName) {
            if (badgesCol !== -1) {
              const badges = String(data[i][badgesCol] || '');
              // Count badges (comma-separated)
              if (badges) {
                missionsCompleted += badges.split(',').filter(b => b.trim()).length;
              }
            }
            break;
          }
        }
      }
    }

    // Count awards from players_awards_lists
    const awardsSheet = ss.getSheetByName('players_awards_lists');
    if (awardsSheet && awardsSheet.getLastRow() > 1) {
      const data = awardsSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = Math.max(headers.indexOf('PreferredName'), headers.indexOf('Name'), headers.indexOf('Player'));

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          const rowName = String(data[i][nameCol] || '').toLowerCase();
          if (rowName === preferredName.toLowerCase()) {
            awardsCount++;
          }
        }
      }
    }
  } catch (e) {
    Logger.log('loadMissionsAndAwardsForPlayer_ error: ' + e.message);
  }

  return { missionsCompleted, awardsCount };
}

/**
 * Loads attendance profile for a player
 * @param {string} preferredName - Player name
 * @return {Object} { lastVisit, lifetimeVisits, primaryFormat }
 * @private
 */
function loadAttendanceProfileForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lastVisit = null;
  let lifetimeVisits = 0;
  let primaryFormat = null;
  const formatCounts = {};

  try {
    // Check Attendance_Calendar sheet
    const attendSheet = ss.getSheetByName('Attendance_Calendar');
    if (attendSheet && attendSheet.getLastRow() > 1) {
      const data = attendSheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = Math.max(headers.indexOf('PreferredName'), headers.indexOf('preferred_name_id'), headers.indexOf('Name'));
      const dateCol = Math.max(headers.indexOf('Event_Date'), headers.indexOf('Date'), headers.indexOf('Timestamp'));
      const typeCol = Math.max(headers.indexOf('Event_Type'), headers.indexOf('Format'), headers.indexOf('Type'));

      if (nameCol !== -1) {
        let maxDate = null;

        for (let i = 1; i < data.length; i++) {
          const rowName = String(data[i][nameCol] || '').toLowerCase();
          if (rowName === preferredName.toLowerCase()) {
            lifetimeVisits++;

            // Track date for last visit
            if (dateCol !== -1 && data[i][dateCol]) {
              const eventDate = data[i][dateCol];
              if (eventDate instanceof Date) {
                if (!maxDate || eventDate > maxDate) {
                  maxDate = eventDate;
                }
              }
            }

            // Track format frequency
            if (typeCol !== -1 && data[i][typeCol]) {
              const format = String(data[i][typeCol]).trim();
              if (format) {
                formatCounts[format] = (formatCounts[format] || 0) + 1;
              }
            }
          }
        }

        if (maxDate) {
          lastVisit = Utilities.formatDate(maxDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      }
    }

    // Fallback: scan event tabs (MM-DD-YYYY format) for player attendance
    if (lifetimeVisits === 0) {
      const eventTabs = listEventTabs();
      let maxDate = null;

      eventTabs.forEach(tabName => {
        const sheet = ss.getSheetByName(tabName);
        if (sheet && sheet.getLastRow() > 1) {
          const data = sheet.getDataRange().getValues();
          const headers = data[0];
          const nameCol = Math.max(headers.indexOf('PreferredName'), 1); // Default to column B

          for (let i = 1; i < data.length; i++) {
            const rowName = String(data[i][nameCol] || '').toLowerCase();
            if (rowName === preferredName.toLowerCase()) {
              lifetimeVisits++;

              // Parse date from tab name (MM-DD-YYYY)
              const match = tabName.match(/^(\d{2})-(\d{2})-(\d{4})/);
              if (match) {
                const eventDate = new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));
                if (!maxDate || eventDate > maxDate) {
                  maxDate = eventDate;
                }

                // Try to get event type from metadata
                const props = getEventProps(sheet);
                if (props.event_type) {
                  formatCounts[props.event_type] = (formatCounts[props.event_type] || 0) + 1;
                }
              }
              break; // Player found in this tab
            }
          }
        }
      });

      if (maxDate) {
        lastVisit = Utilities.formatDate(maxDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
    }

    // Determine primary format (most common)
    let maxCount = 0;
    Object.keys(formatCounts).forEach(format => {
      if (formatCounts[format] > maxCount) {
        maxCount = formatCounts[format];
        primaryFormat = format;
      }
    });

    // If multiple formats with same count, say "Mixed"
    const topFormats = Object.keys(formatCounts).filter(f => formatCounts[f] === maxCount);
    if (topFormats.length > 1) {
      primaryFormat = 'Mixed';
    }

  } catch (e) {
    Logger.log('loadAttendanceProfileForPlayer_ error: ' + e.message);
  }

  return { lastVisit, lifetimeVisits, primaryFormat };
}

/**
 * Loads queue counts (preorders, this needs) for a player
 * @param {string} preferredName - Player name
 * @return {Object} { preordersOpen, thisNeedsOpen }
 * @private
 */
function loadQueuesProfileForPlayer_(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let preordersOpen = 0;
  let thisNeedsOpen = 0;

  try {
    // Count open preorders
    const preorderSheets = ['Preorders', 'Preorder_Sales', 'PreorderSales'];
    preorderSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = Math.max(
          headers.indexOf('PreferredName'),
          headers.indexOf('Name'),
          headers.indexOf('Customer'),
          headers.indexOf('Contact')
        );
        const statusCol = headers.indexOf('Status');

        if (nameCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const rowName = String(data[i][nameCol] || '').toLowerCase();
            if (rowName === preferredName.toLowerCase()) {
              if (statusCol === -1) {
                // No status column - count all as open
                preordersOpen++;
              } else {
                const status = String(data[i][statusCol] || '').toUpperCase();
                // Count if not picked up or cancelled
                if (status !== 'PICKED_UP' && status !== 'CANCELLED' && status !== 'COMPLETE' && status !== 'COMPLETED') {
                  preordersOpen++;
                }
              }
            }
          }
        }
      }
    });

    // Count open "This Needs" assignments
    const assignmentSheets = ['This_Needs', 'ThisNeeds', 'Employee_Log', 'Assignments', 'Tasks'];
    assignmentSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet && sheet.getLastRow() > 1) {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const nameCol = Math.max(
          headers.indexOf('customer_name'),
          headers.indexOf('PreferredName'),
          headers.indexOf('Customer'),
          headers.indexOf('Name'),
          headers.indexOf('Player')
        );
        const statusCol = Math.max(headers.indexOf('status'), headers.indexOf('Status'));
        const completedCol = Math.max(headers.indexOf('completed'), headers.indexOf('Completed'));

        if (nameCol !== -1) {
          for (let i = 1; i < data.length; i++) {
            const rowName = String(data[i][nameCol] || '').toLowerCase();
            if (rowName === preferredName.toLowerCase()) {
              let isOpen = true;

              // Check status column
              if (statusCol !== -1) {
                const status = String(data[i][statusCol] || '').toLowerCase();
                if (status === 'complete' || status === 'completed' || status === 'done' || status === 'closed') {
                  isOpen = false;
                }
              }

              // Check completed column
              if (isOpen && completedCol !== -1) {
                const completed = data[i][completedCol];
                if (coerceBoolean(completed)) {
                  isOpen = false;
                }
              }

              if (isOpen) {
                thisNeedsOpen++;
              }
            }
          }
        }
      }
    });

  } catch (e) {
    Logger.log('loadQueuesProfileForPlayer_ error: ' + e.message);
  }

  return { preordersOpen, thisNeedsOpen };
}