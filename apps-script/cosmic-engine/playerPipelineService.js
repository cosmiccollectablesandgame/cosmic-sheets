/**
 * Player Pipeline Services - Service Stubs
 * @fileoverview Service stubs for player pipeline data sources.
 *               These services handle: Preorders, This Needs, Missions,
 *               Attendance, and Store Credit.
 *
 * Version: 7.9.7
 *
 * INTEGRATION NOTES:
 * - Each service reads from its dedicated sheet
 * - Services return typed objects for the UI handler to consume
 * - Replace stub implementations with actual business logic as sheets are populated
 *
 * Architecture:
 * - UI Handler (ui.handler.playerLookup.gs) calls these services
 * - Services own raw sheet logic for their domain
 * - Services should NOT call UI handler functions (one-way dependency)
 */

// ============================================================================
// PREORDERS SERVICE
// ============================================================================

/**
 * @typedef {Object} PreorderItem
 * @property {string} product - Product name/SKU
 * @property {string} status - Order status (Pending, Ready, Fulfilled)
 * @property {string} [orderedDate] - When order was placed
 * @property {string} [expectedDate] - Expected arrival
 * @property {number} [quantity] - Quantity ordered
 * @property {number} [deposit] - Deposit amount paid
 */

/**
 * @typedef {Object} PlayerPreordersResult
 * @property {number} count - Number of active preorders
 * @property {string} summary - Human-readable summary
 * @property {PreorderItem[]} items - Individual preorder items
 * @property {number} totalDeposits - Total deposits on file
 */

/**
 * Ensures Preorders sheet exists with canonical headers.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The Preorders sheet
 */
function ensurePreordersSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preorders');

  const headers = [
    'PreferredName',
    'Product',
    'SKU',
    'Quantity',
    'Deposit',
    'Status',
    'Ordered_Date',
    'Expected_Date',
    'Notes',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Preorders');
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    return sheet;
  }

  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Gets all preorders for a player.
 * @param {string} preferredName - Player name
 * @return {PlayerPreordersResult} Preorders result
 */
function getPlayerPreorders(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders');

  const result = {
    count: 0,
    summary: 'No active preorders',
    items: [],
    totalDeposits: 0
  };

  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const productCol = headers.indexOf('Product');
  const skuCol = headers.indexOf('SKU');
  const qtyCol = headers.indexOf('Quantity');
  const depositCol = headers.indexOf('Deposit');
  const statusCol = headers.indexOf('Status');
  const orderedCol = headers.indexOf('Ordered_Date');
  const expectedCol = headers.indexOf('Expected_Date');

  if (nameCol === -1) return result;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const status = statusCol !== -1 ? String(data[i][statusCol] || 'Pending') : 'Pending';

      // Only include non-fulfilled orders
      if (status.toLowerCase() !== 'fulfilled' && status.toLowerCase() !== 'cancelled') {
        const item = {
          product: productCol !== -1 ? data[i][productCol] : (skuCol !== -1 ? data[i][skuCol] : 'Unknown'),
          status: status,
          orderedDate: orderedCol !== -1 ? formatDateSafe_(data[i][orderedCol]) : '',
          expectedDate: expectedCol !== -1 ? formatDateSafe_(data[i][expectedCol]) : '',
          quantity: qtyCol !== -1 ? coerceNumber(data[i][qtyCol], 1) : 1,
          deposit: depositCol !== -1 ? coerceNumber(data[i][depositCol], 0) : 0
        };

        result.items.push(item);
        result.totalDeposits += item.deposit;
      }
    }
  }

  result.count = result.items.length;
  if (result.count > 0) {
    const productList = result.items.map(i => i.product).slice(0, 3).join(', ');
    result.summary = `${result.count} preorder(s): ${productList}${result.count > 3 ? '...' : ''}`;
  }

  return result;
}

/**
 * Adds a preorder for a player.
 * @param {string} preferredName - Player name
 * @param {Object} preorder - Preorder data {product, sku, quantity, deposit}
 * @return {Object} Result {success, message}
 */
function addPlayerPreorder(preferredName, preorder) {
  try {
    ensurePreordersSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Preorders');

    const row = [
      preferredName,
      preorder.product || '',
      preorder.sku || '',
      preorder.quantity || 1,
      preorder.deposit || 0,
      'Pending',
      dateISO(),
      preorder.expectedDate || '',
      preorder.notes || '',
      dateISO()
    ];

    sheet.appendRow(row);

    logIntegrityAction('PREORDER_ADD', {
      preferredName,
      details: `Added preorder: ${preorder.product || preorder.sku}`,
      status: 'SUCCESS'
    });

    return { success: true, message: 'Preorder added' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Updates preorder status.
 * @param {string} preferredName - Player name
 * @param {string} product - Product identifier
 * @param {string} newStatus - New status (Ready, Fulfilled, Cancelled)
 * @return {Object} Result {success, message}
 */
function updatePreorderStatus(preferredName, product, newStatus) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Preorders');

    if (!sheet) return { success: false, message: 'Preorders sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');
    const productCol = headers.indexOf('Product');
    const statusCol = headers.indexOf('Status');
    const updatedCol = headers.indexOf('LastUpdated');

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === preferredName &&
          (data[i][productCol] === product || data[i][headers.indexOf('SKU')] === product)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        if (updatedCol !== -1) {
          sheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());
        }

        logIntegrityAction('PREORDER_STATUS', {
          preferredName,
          details: `${product}: ${data[i][statusCol]} -> ${newStatus}`,
          status: 'SUCCESS'
        });

        return { success: true, message: 'Status updated' };
      }
    }

    return { success: false, message: 'Preorder not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// THIS NEEDS SERVICE (Service Tasks)
// ============================================================================

/**
 * @typedef {Object} ServiceTask
 * @property {string} task - Task description
 * @property {string} status - Task status (Pending, In Progress, Done)
 * @property {string} [priority] - Priority level (Low, Normal, High, Urgent)
 * @property {string} [createdDate] - When task was created
 * @property {string} [assignedTo] - Staff member assigned
 */

/**
 * @typedef {Object} PlayerServiceNeedsResult
 * @property {number} count - Number of pending tasks
 * @property {string} summary - Human-readable summary
 * @property {ServiceTask[]} items - Individual tasks
 * @property {number} urgentCount - Number of urgent tasks
 */

/**
 * Ensures This_Needs sheet exists with canonical headers.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The This_Needs sheet
 */
function ensureThisNeedsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('This_Needs');

  const headers = [
    'PreferredName',
    'Task',
    'Priority',
    'Status',
    'Assigned_To',
    'Created_Date',
    'Completed_Date',
    'Notes',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('This_Needs');
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#ea4335')
      .setFontColor('#ffffff');
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Gets all pending service tasks for a player.
 * @param {string} preferredName - Player name
 * @return {PlayerServiceNeedsResult} Service needs result
 */
function getPlayerServiceNeeds(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('This_Needs');

  const result = {
    count: 0,
    summary: 'No pending tasks',
    items: [],
    urgentCount: 0
  };

  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const taskCol = headers.indexOf('Task');
  const priorityCol = headers.indexOf('Priority');
  const statusCol = headers.indexOf('Status');
  const assignedCol = headers.indexOf('Assigned_To');
  const createdCol = headers.indexOf('Created_Date');

  if (nameCol === -1) return result;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const status = statusCol !== -1 ? String(data[i][statusCol] || 'Pending').toLowerCase() : 'pending';

      if (status !== 'done' && status !== 'completed') {
        const priority = priorityCol !== -1 ? String(data[i][priorityCol] || 'Normal') : 'Normal';
        const task = {
          task: taskCol !== -1 ? data[i][taskCol] : 'Task',
          status: status,
          priority: priority,
          createdDate: createdCol !== -1 ? formatDateSafe_(data[i][createdCol]) : '',
          assignedTo: assignedCol !== -1 ? data[i][assignedCol] : ''
        };

        result.items.push(task);
        if (priority.toLowerCase() === 'urgent' || priority.toLowerCase() === 'high') {
          result.urgentCount++;
        }
      }
    }
  }

  result.count = result.items.length;
  if (result.count > 0) {
    const taskList = result.items.map(t => t.task).slice(0, 2).join(', ');
    const urgentTag = result.urgentCount > 0 ? ` (${result.urgentCount} urgent!)` : '';
    result.summary = `${result.count} task(s): ${taskList}${result.count > 2 ? '...' : ''}${urgentTag}`;
  }

  return result;
}

/**
 * Adds a service task for a player.
 * @param {string} preferredName - Player name
 * @param {Object} task - Task data {task, priority, assignedTo, notes}
 * @return {Object} Result {success, message}
 */
function addPlayerServiceTask(preferredName, task) {
  try {
    ensureThisNeedsSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('This_Needs');

    const row = [
      preferredName,
      task.task || 'New Task',
      task.priority || 'Normal',
      'Pending',
      task.assignedTo || '',
      dateISO(),
      '',
      task.notes || '',
      dateISO()
    ];

    sheet.appendRow(row);

    logIntegrityAction('SERVICE_TASK_ADD', {
      preferredName,
      details: `Added task: ${task.task}`,
      status: 'SUCCESS'
    });

    return { success: true, message: 'Task added' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Marks a service task as complete.
 * @param {string} preferredName - Player name
 * @param {string} taskDescription - Task to complete
 * @return {Object} Result {success, message}
 */
function completeServiceTask(preferredName, taskDescription) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('This_Needs');

    if (!sheet) return { success: false, message: 'This_Needs sheet not found' };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');
    const taskCol = headers.indexOf('Task');
    const statusCol = headers.indexOf('Status');
    const completedCol = headers.indexOf('Completed_Date');
    const updatedCol = headers.indexOf('LastUpdated');

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === preferredName && data[i][taskCol] === taskDescription) {
        sheet.getRange(i + 1, statusCol + 1).setValue('Done');
        if (completedCol !== -1) {
          sheet.getRange(i + 1, completedCol + 1).setValue(dateISO());
        }
        if (updatedCol !== -1) {
          sheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());
        }

        logIntegrityAction('SERVICE_TASK_COMPLETE', {
          preferredName,
          details: `Completed: ${taskDescription}`,
          status: 'SUCCESS'
        });

        return { success: true, message: 'Task completed' };
      }
    }

    return { success: false, message: 'Task not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// MISSIONS SERVICE
// ============================================================================

/**
 * @typedef {Object} MissionProgress
 * @property {string} missionId - Mission identifier
 * @property {string} name - Mission name
 * @property {string} status - Status (Active, Completed, Expired)
 * @property {number} progress - Current progress value
 * @property {number} target - Target value to complete
 * @property {string} [reward] - Reward description
 */

/**
 * @typedef {Object} PlayerMissionsResult
 * @property {number} completedCount - Missions completed
 * @property {number} activeCount - Missions in progress
 * @property {string} summary - Human-readable summary
 * @property {MissionProgress[]} active - Active missions
 * @property {string[]} completedRecently - Recently completed mission names
 */

/**
 * Gets missions data for a player from Attendance_Missions.
 * @param {string} preferredName - Player name
 * @return {PlayerMissionsResult} Missions result
 */
function getPlayerMissions(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');

  const result = {
    completedCount: 0,
    activeCount: 0,
    summary: 'No mission data',
    active: [],
    completedRecently: []
  };

  if (!sheet || sheet.getLastRow() <= 1) return result;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('PreferredName');
  const diceCol = headers.indexOf('Points_From_Dice_Rolls');
  const topCol = headers.indexOf('Bonus_Points_From_Top4');
  const totalCol = headers.indexOf('C_Total');
  const badgesCol = headers.indexOf('Badges');

  if (nameCol === -1) return result;

  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] === preferredName) {
      const dicePoints = diceCol !== -1 ? coerceNumber(data[i][diceCol], 0) : 0;
      const topPoints = topCol !== -1 ? coerceNumber(data[i][topCol], 0) : 0;
      const total = totalCol !== -1 ? coerceNumber(data[i][totalCol], 0) : 0;
      const badges = badgesCol !== -1 ? String(data[i][badgesCol] || '') : '';

      result.completedCount = total;
      result.summary = `${total} missions completed`;

      if (badges) {
        result.completedRecently = badges.split(',').map(b => b.trim()).filter(b => b);
      }

      // Build active missions based on point thresholds
      // This is a simplified example - real implementation would track individual missions
      if (dicePoints > 0 && dicePoints < 10) {
        result.active.push({
          missionId: 'dice_master',
          name: 'Dice Master',
          status: 'Active',
          progress: dicePoints,
          target: 10,
          reward: '1 Random Key'
        });
        result.activeCount++;
      }

      break;
    }
  }

  return result;
}

/**
 * Awards a mission completion to a player.
 * @param {string} preferredName - Player name
 * @param {string} missionId - Mission identifier
 * @return {Object} Result {success, message, reward}
 */
function completeMission(preferredName, missionId) {
  // STUB: Implement mission completion logic
  // This would update Attendance_Missions and potentially award BP/Keys

  logIntegrityAction('MISSION_COMPLETE', {
    preferredName,
    details: `Completed mission: ${missionId}`,
    status: 'SUCCESS'
  });

  return {
    success: true,
    message: 'Mission completed!',
    reward: 'Reward pending implementation'
  };
}

// ============================================================================
// ATTENDANCE SERVICE
// ============================================================================

/**
 * @typedef {Object} AttendanceRecord
 * @property {string} eventId - Event identifier (date)
 * @property {string} eventType - Event type (Commander, Modern, etc.)
 * @property {string} date - Event date
 * @property {number} [placement] - Final placement if available
 */

/**
 * @typedef {Object} PlayerAttendanceResult
 * @property {string} lastVisit - Most recent visit date
 * @property {number} lifetimeVisits - Total events attended
 * @property {string} primaryFormat - Most attended format
 * @property {AttendanceRecord[]} recentEvents - Recent event attendance
 * @property {Object} formatBreakdown - Events per format
 */

/**
 * Gets attendance data for a player.
 * Scans event tabs and Players_Prize-Wall-Points.
 *
 * @param {string} preferredName - Player name
 * @return {PlayerAttendanceResult} Attendance result
 */
function getPlayerAttendance(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const result = {
    lastVisit: 'Never',
    lifetimeVisits: 0,
    primaryFormat: 'Not specified',
    recentEvents: [],
    formatBreakdown: {}
  };

  // Pattern for event tabs (MM-DD-YYYY)
  const eventPattern = /^(\d{2})-(\d{2})-(\d{4})$/;
  const events = [];

  for (const sheet of sheets) {
    const name = sheet.getName();
    const match = name.match(eventPattern);

    if (match && sheet.getLastRow() > 1) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const nameCol = headers.indexOf('PreferredName');
      const rankCol = headers.indexOf('Rank');

      if (nameCol !== -1) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][nameCol] === preferredName) {
            const eventDate = new Date(parseInt(match[3]), parseInt(match[1]) - 1, parseInt(match[2]));

            // Try to get event type from metadata
            let eventType = 'Unknown';
            try {
              const props = getEventProps(sheet);
              eventType = props.event_type || 'Unknown';
            } catch (e) {
              // Metadata not available
            }

            events.push({
              eventId: name,
              eventType: eventType,
              date: name,
              dateObj: eventDate,
              placement: rankCol !== -1 ? coerceNumber(data[i][rankCol], 0) : null
            });

            // Track format breakdown
            if (!result.formatBreakdown[eventType]) {
              result.formatBreakdown[eventType] = 0;
            }
            result.formatBreakdown[eventType]++;

            break;
          }
        }
      }
    }
  }

  // Sort events by date (most recent first)
  events.sort((a, b) => b.dateObj - a.dateObj);

  result.lifetimeVisits = events.length;
  result.recentEvents = events.slice(0, 5).map(e => ({
    eventId: e.eventId,
    eventType: e.eventType,
    date: e.date,
    placement: e.placement
  }));

  if (events.length > 0) {
    result.lastVisit = events[0].date;
  }

  // Determine primary format
  let maxFormat = '';
  let maxCount = 0;
  for (const [format, count] of Object.entries(result.formatBreakdown)) {
    if (count > maxCount && format !== 'Unknown') {
      maxCount = count;
      maxFormat = format;
    }
  }
  if (maxFormat) {
    result.primaryFormat = maxFormat;
  }

  return result;
}

// ============================================================================
// STORE CREDIT SERVICE
// ============================================================================

/**
 * @typedef {Object} StoreCreditTransaction
 * @property {string} date - Transaction date
 * @property {string} type - Type (Add, Redeem, Adjust)
 * @property {number} amount - Transaction amount
 * @property {string} [reference] - Reference/reason
 */

/**
 * @typedef {Object} PlayerStoreCreditResult
 * @property {number} balance - Current balance
 * @property {string} summary - Human-readable summary
 * @property {StoreCreditTransaction[]} recentTransactions - Recent transactions
 * @property {string} lastUpdated - Last update timestamp
 */

/**
 * Ensures Store_Credit sheet exists with canonical headers.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The Store_Credit sheet
 */
function ensureStoreCreditSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Store_Credit');

  const headers = [
    'PreferredName',
    'Balance',
    'LastUpdated'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Store_Credit');
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#fbbc04')
      .setFontColor('#000000');
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Ensures Store_Credit_Ledger sheet exists for transaction history.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The ledger sheet
 */
function ensureStoreCreditLedgerSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Store_Credit_Ledger');

  const headers = [
    'PreferredName',
    'Date',
    'Type',
    'Amount',
    'Balance_After',
    'Reference',
    'Operator'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Store_Credit_Ledger');
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    return sheet;
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Gets store credit data for a player.
 * @param {string} preferredName - Player name
 * @return {PlayerStoreCreditResult} Store credit result
 */
function getPlayerStoreCredit(preferredName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Store_Credit');

  const result = {
    balance: 0,
    summary: '$0.00',
    recentTransactions: [],
    lastUpdated: null
  };

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');
    const balanceCol = headers.indexOf('Balance');
    const updatedCol = headers.indexOf('LastUpdated');

    if (nameCol !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][nameCol] === preferredName) {
          result.balance = balanceCol !== -1 ? coerceNumber(data[i][balanceCol], 0) : 0;
          result.lastUpdated = updatedCol !== -1 ? formatDateSafe_(data[i][updatedCol]) : null;
          result.summary = formatCurrency(result.balance);
          break;
        }
      }
    }
  }

  // Get recent transactions from ledger
  const ledger = ss.getSheetByName('Store_Credit_Ledger');
  if (ledger && ledger.getLastRow() > 1) {
    const ledgerData = ledger.getDataRange().getValues();
    const ledgerHeaders = ledgerData[0];
    const lNameCol = ledgerHeaders.indexOf('PreferredName');
    const lDateCol = ledgerHeaders.indexOf('Date');
    const lTypeCol = ledgerHeaders.indexOf('Type');
    const lAmountCol = ledgerHeaders.indexOf('Amount');
    const lRefCol = ledgerHeaders.indexOf('Reference');

    if (lNameCol !== -1) {
      const transactions = [];
      for (let i = 1; i < ledgerData.length; i++) {
        if (ledgerData[i][lNameCol] === preferredName) {
          transactions.push({
            date: lDateCol !== -1 ? formatDateSafe_(ledgerData[i][lDateCol]) : '',
            type: lTypeCol !== -1 ? ledgerData[i][lTypeCol] : '',
            amount: lAmountCol !== -1 ? coerceNumber(ledgerData[i][lAmountCol], 0) : 0,
            reference: lRefCol !== -1 ? ledgerData[i][lRefCol] : ''
          });
        }
      }
      // Return most recent 5
      result.recentTransactions = transactions.slice(-5).reverse();
    }
  }

  return result;
}

/**
 * Adds store credit to a player.
 * @param {string} preferredName - Player name
 * @param {number} amount - Amount to add
 * @param {string} reference - Reference/reason
 * @return {Object} Result {success, newBalance, message}
 */
function addStoreCredit(preferredName, amount, reference = '') {
  try {
    ensureStoreCreditSchema();
    ensureStoreCreditLedgerSchema();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Store_Credit');
    const ledger = ss.getSheetByName('Store_Credit_Ledger');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');
    const balanceCol = headers.indexOf('Balance');
    const updatedCol = headers.indexOf('LastUpdated');

    let playerRow = -1;
    let currentBalance = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === preferredName) {
        playerRow = i;
        currentBalance = coerceNumber(data[i][balanceCol], 0);
        break;
      }
    }

    const newBalance = currentBalance + amount;

    if (playerRow === -1) {
      // Add new player
      sheet.appendRow([preferredName, newBalance, dateISO()]);
    } else {
      // Update existing
      sheet.getRange(playerRow + 1, balanceCol + 1).setValue(newBalance);
      if (updatedCol !== -1) {
        sheet.getRange(playerRow + 1, updatedCol + 1).setValue(dateISO());
      }
    }

    // Log to ledger
    ledger.appendRow([
      preferredName,
      dateISO(),
      'Add',
      amount,
      newBalance,
      reference,
      currentUser()
    ]);

    logIntegrityAction('STORE_CREDIT_ADD', {
      preferredName,
      details: `Added ${formatCurrency(amount)}. Balance: ${formatCurrency(currentBalance)} -> ${formatCurrency(newBalance)}`,
      status: 'SUCCESS'
    });

    return {
      success: true,
      newBalance: newBalance,
      message: `Added ${formatCurrency(amount)}. New balance: ${formatCurrency(newBalance)}`
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Redeems store credit from a player.
 * @param {string} preferredName - Player name
 * @param {number} amount - Amount to redeem
 * @param {string} reference - Reference/reason
 * @return {Object} Result {success, newBalance, message}
 */
function redeemStoreCredit(preferredName, amount, reference = '') {
  try {
    const current = getPlayerStoreCredit(preferredName);

    if (current.balance < amount) {
      return {
        success: false,
        message: `Insufficient balance. Available: ${formatCurrency(current.balance)}`
      };
    }

    ensureStoreCreditLedgerSchema();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Store_Credit');
    const ledger = ss.getSheetByName('Store_Credit_Ledger');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameCol = headers.indexOf('PreferredName');
    const balanceCol = headers.indexOf('Balance');
    const updatedCol = headers.indexOf('LastUpdated');

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameCol] === preferredName) {
        const newBalance = current.balance - amount;
        sheet.getRange(i + 1, balanceCol + 1).setValue(newBalance);
        if (updatedCol !== -1) {
          sheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());
        }

        // Log to ledger
        ledger.appendRow([
          preferredName,
          dateISO(),
          'Redeem',
          -amount,
          newBalance,
          reference,
          currentUser()
        ]);

        logIntegrityAction('STORE_CREDIT_REDEEM', {
          preferredName,
          details: `Redeemed ${formatCurrency(amount)}. Balance: ${formatCurrency(current.balance)} -> ${formatCurrency(newBalance)}`,
          status: 'SUCCESS'
        });

        return {
          success: true,
          newBalance: newBalance,
          message: `Redeemed ${formatCurrency(amount)}. New balance: ${formatCurrency(newBalance)}`
        };
      }
    }

    return { success: false, message: 'Player not found' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Safely formats a date value.
 * @param {*} value - Date value
 * @return {string} Formatted date or empty string
 * @private
 */
function formatDateSafe_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (typeof value === 'string') {
    if (value.includes('T')) {
      return value.split('T')[0];
    }
    return value;
  }
  return String(value);
}