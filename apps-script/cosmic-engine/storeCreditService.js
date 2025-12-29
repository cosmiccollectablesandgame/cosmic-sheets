/**
 * Store Credit Service
 * @fileoverview Handles store credit transactions for the Store_Credit_Ledger
 */

const STORE_CREDIT_SHEET_NAME = 'Store_Credit_Ledger';

/**
 * Logs a store credit transaction to the ledger
 * @param {Object} payload - Transaction payload from the UI
 * @param {string} payload.preferred_name_id - Player identifier (required)
 * @param {string} payload.direction - "IN" or "OUT" (required)
 * @param {number|string} payload.amount - Transaction amount (required, must be positive)
 * @param {string} payload.reason - Reason for transaction
 * @param {string} payload.category - Category (e.g., "Prize Payout", "Manual Adjust")
 * @param {string} payload.tenderType - Tender type (e.g., "Store Credit", "Gift Card")
 * @param {string} payload.description - Additional description
 * @param {string} payload.posRefType - POS reference type (e.g., "Invoice", "TicketID")
 * @param {string} payload.posRefId - POS reference ID
 * @return {Object} Result object with success status and transaction details
 */
function logStoreCreditTransaction(payload) {
  try {
    // Validate required fields
    if (!payload.preferred_name_id || String(payload.preferred_name_id).trim() === '') {
      throw new Error('preferred_name_id is required');
    }

    const direction = String(payload.direction).trim().toUpperCase();
    if (direction !== 'IN' && direction !== 'OUT') {
      throw new Error('direction must be "IN" or "OUT"');
    }

    const amount = parseFloat(payload.amount);
    if (isNaN(amount) || amount <= 0) {
      throw new Error('amount must be a positive number');
    }

    // Get the ledger sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(STORE_CREDIT_SHEET_NAME);

    if (!sheet) {
      throw new Error('Store_Credit_Ledger sheet not found. Please create it first.');
    }

    // Calculate signed amount
    const signedAmount = direction === 'IN' ? amount : -amount;

    // Get player's last running balance
    const preferredNameId = String(payload.preferred_name_id).trim();
    const lastBalance = getLastRunningBalance_(sheet, preferredNameId);
    const newBalance = lastBalance + signedAmount;

    // Generate unique row ID and timestamp
    const rowId = Utilities.getUuid();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSSZ");

    // Build the row data matching ledger headers:
    // Timestamp, preferred_name_id, InOut, Amount, Reason, Category, TenderType,
    // Description, POSRefType, POSRefId, RunningBalance, RowId
    const rowData = [
      timestamp,
      preferredNameId,
      direction,
      signedAmount,
      payload.reason || '',
      payload.category || '',
      payload.tenderType || '',
      payload.description || '',
      payload.posRefType || '',
      payload.posRefId || '',
      newBalance,
      rowId
    ];

    // Append the row
    sheet.appendRow(rowData);

    // Flag that the ledger was updated (for polling/checkForLedgerUpdates)
    PropertiesService.getScriptProperties()
      .setProperty('LEDGER_LAST_UPDATED', String(now.getTime()));

    // Log to integrity log if available
    try {
      logIntegrityAction('STORE_CREDIT_' + direction, {
        preferred_name_id: preferredNameId,
        amount: signedAmount,
        newBalance: newBalance,
        rowId: rowId
      });
    } catch (logError) {
      // Integrity logging is optional - don't fail the transaction
      console.warn('Failed to log to integrity log:', logError);
    }

    return {
      success: true,
      preferred_name_id: preferredNameId,
      newBalance: newBalance,
      direction: direction,
      amount: signedAmount,
      timestamp: timestamp,
      rowId: rowId
    };

  } catch (error) {
    console.error('logStoreCreditTransaction error:', error);
    throw new Error('Failed to log store credit transaction: ' + error.message);
  }
}

/**
 * Gets the last running balance for a player from the ledger
 * Scans from bottom up for efficiency
 * @param {Sheet} sheet - The Store_Credit_Ledger sheet
 * @param {string} preferredNameId - The player identifier
 * @return {number} Last running balance (0 if none found)
 * @private
 */
function getLastRunningBalance_(sheet, preferredNameId) {
  const lastRow = sheet.getLastRow();

  // If only header row exists, return 0
  if (lastRow <= 1) {
    return 0;
  }

  // Get all data (excluding header)
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Scan from bottom up looking for matching preferred_name_id (column B = index 1)
  for (let i = data.length - 1; i >= 0; i--) {
    const rowPreferredNameId = String(data[i][1]).trim();
    if (rowPreferredNameId === preferredNameId) {
      // Column K (RunningBalance) = index 10
      const balance = parseFloat(data[i][10]);
      return isNaN(balance) ? 0 : balance;
    }
  }

  // No matching rows found
  return 0;
}

/**
 * Gets the current store credit balance for a player
 * @param {string} preferredNameId - The player identifier
 * @return {Object} Balance information
 */
function getStoreCreditBalance(preferredNameId) {
  try {
    if (!preferredNameId || String(preferredNameId).trim() === '') {
      throw new Error('preferred_name_id is required');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(STORE_CREDIT_SHEET_NAME);

    if (!sheet) {
      return {
        success: true,
        preferred_name_id: preferredNameId,
        balance: 0,
        found: false
      };
    }

    const balance = getLastRunningBalance_(sheet, String(preferredNameId).trim());

    return {
      success: true,
      preferred_name_id: preferredNameId,
      balance: balance,
      found: balance !== 0 || hasTransactions_(sheet, preferredNameId)
    };

  } catch (error) {
    console.error('getStoreCreditBalance error:', error);
    throw new Error('Failed to get store credit balance: ' + error.message);
  }
}

/**
 * Checks if a player has any transactions in the ledger
 * @param {Sheet} sheet - The Store_Credit_Ledger sheet
 * @param {string} preferredNameId - The player identifier
 * @return {boolean} True if player has transactions
 * @private
 */
function hasTransactions_(sheet, preferredNameId) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;

  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column B only
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === preferredNameId) {
      return true;
    }
  }
  return false;
}

/**
 * Ensures the Store_Credit_Ledger sheet exists with proper headers
 * @return {Sheet} The ledger sheet
 */
function ensureStoreCreditLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(STORE_CREDIT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(STORE_CREDIT_SHEET_NAME);
    sheet.appendRow([
      'Timestamp',
      'preferred_name_id',
      'InOut',
      'Amount',
      'Reason',
      'Category',
      'TenderType',
      'Description',
      'POSRefType',
      'POSRefId',
      'RunningBalance',
      'RowId'
    ]);
    sheet.setFrozenRows(1);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, 12);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');
  }

  return sheet;
}


// ============================================================================
// UI WRAPPER FUNCTIONS (called by store_credit.html)
// ============================================================================

/**
 * Called by ui/store_credit.html when the user submits the form.
 * Delegates to logStoreCreditTransaction() and returns newBalance.
 *
 * @param {Object} payload - Form data from HTML
 * @param {string} payload.playerName - Player name/ID
 * @param {string} payload.direction - "IN" or "OUT"
 * @param {number} payload.amount - Transaction amount
 * @param {string} payload.reason - Reason for transaction
 * @param {string} payload.category - Category
 * @param {string} payload.tenderType - Tender type
 * @param {string} payload.description - Description
 * @param {string} payload.posRefType - POS reference type
 * @param {string} payload.posRefId - POS reference ID
 * @return {Object} { newBalance: number }
 */
function submitStoreCredit(payload) {
  if (!payload || !payload.playerName) {
    throw new Error('Missing player name.');
  }

  var amount = Number(payload.amount);
  if (!amount || amount <= 0) {
    throw new Error('Amount must be greater than zero.');
  }

  // Normalize to "IN" / "OUT"
  var direction = String(payload.direction || '').toUpperCase();
  if (direction !== 'IN' && direction !== 'OUT') {
    throw new Error('Direction must be IN or OUT.');
  }

  // Build payload expected by logStoreCreditTransaction
  var ledgerPayload = {
    preferred_name_id: payload.playerName,
    direction: direction,
    amount: amount,
    reason: payload.reason || '',
    category: payload.category || '',
    tenderType: payload.tenderType || '',
    description: payload.description || '',
    posRefType: payload.posRefType || '',
    posRefId: payload.posRefId || ''
  };

  // Use core service to write the row + calculate running balance
  var result = logStoreCreditTransaction(ledgerPayload);

  // Return numeric newBalance for the HTML
  return {
    newBalance: result.newBalance
  };
}

/**
 * Used by the HTML "Balance Inquiry" button.
 * Returns a NUMBER (store credit balance) for the selected player.
 * @param {string} playerName - Player name/ID
 * @return {number} Current balance
 */
function getCurrentBalance(playerName) {
  if (!playerName) {
    throw new Error('Player name is required.');
  }

  // Delegate to core service
  var result = getStoreCreditBalance(playerName);
  return Number(result.balance) || 0;
}

/**
 * Used by ui/store_credit.html to show the player's recent history.
 * Reads from Store_Credit_Ledger using header indexing.
 *
 * @param {string} playerName - Player name/ID
 * @param {number} limit - Max transactions to return (default 5)
 * @return {Array} Array of transaction objects, newest first
 */
function getPlayerHistory(playerName, limit) {
  if (!playerName) return [];
  limit = limit || 5;

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(STORE_CREDIT_SHEET_NAME);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];

  // Find column indices dynamically
  var tsCol        = headers.indexOf('Timestamp');
  var nameCol      = headers.indexOf('preferred_name_id');
  var directionCol = headers.indexOf('InOut');
  var amountCol    = headers.indexOf('Amount');
  var reasonCol    = headers.indexOf('Reason');
  var categoryCol  = headers.indexOf('Category');
  var descCol      = headers.indexOf('Description');
  var balanceCol   = headers.indexOf('RunningBalance');

  if (nameCol === -1 || directionCol === -1 || amountCol === -1) {
    console.warn('getPlayerHistory: Required columns not found in header');
    return [];
  }

  var txs = [];

  // Walk from bottom up (newest rows last in sheet)
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var name = String(row[nameCol] || '').trim();
    if (name !== playerName) continue;

    txs.push({
      timestamp: row[tsCol] || '',
      direction: row[directionCol] || '',
      amount: Number(row[amountCol]) || 0,
      balance: balanceCol !== -1 ? (Number(row[balanceCol]) || 0) : 0,
      reason: reasonCol !== -1 ? (row[reasonCol] || '') : '',
      category: categoryCol !== -1 ? (row[categoryCol] || '') : '',
      description: descCol !== -1 ? (row[descCol] || '') : ''
    });

    if (txs.length >= limit) break;
  }

  return txs;
}