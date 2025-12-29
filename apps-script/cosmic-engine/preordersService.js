/**
 * Preorder Service - Preorder Management for Cosmic Event Manager
 * @fileoverview Handles preorder buckets, sales, and customer search
 *
 * Sheet Schemas:
 * - Preorders_Buckets: Set_Name, Item_Name, Item_Code, Unit_Cost, Preorder_Price, Quantity,
 *                      Status, Date_Added, Notes, Reserved, Available, Release_Date,
 *                      LastUpdated, MSRP_Price, Retail_Price, Sold_Per_Pack_Price, Sold_Per_Box_Price
 * - Preorders_Sold: Preorder_ID, PreferredName, Contact_Info, Set_Name, Item_Name, Item_Code,
 *                   Qty, Unit_Price, Line_Total, Total_Due, Deposit_Paid, Balance_Due,
 *                   Target_Payoff, Status, Notes, Created_At, Created_By
 */

// ============================================================================
// SCHEMA HELPERS
// ============================================================================

/**
 * Ensures the Preorders_Buckets sheet exists with correct headers
 */
function ensurePreordersBucketsSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preorders_Buckets');

  const headers = [
    'Set_Name',
    'Item_Name',
    'Item_Code',
    'Unit_Cost',
    'Preorder_Price',
    'Quantity',
    'Status',
    'Date_Added',
    'Notes',
    'Reserved',
    'Available',
    'Release_Date',
    'LastUpdated',
    'MSRP_Price',
    'Retail_Price',
    'Sold_Per_Pack_Price',
    'Sold_Per_Box_Price'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Preorders_Buckets');
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Ensures the Preorders_Sold sheet exists with correct headers
 */
function ensurePreordersSoldSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preorders_Sold');

  const headers = [
    'Preorder_ID',
    'PreferredName',
    'Contact_Info',
    'Set_Name',
    'Item_Name',
    'Item_Code',
    'Qty',
    'Unit_Price',
    'Line_Total',
    'Total_Due',
    'Deposit_Paid',
    'Balance_Due',
    'Target_Payoff',
    'Status',
    'Notes',
    'Created_At',
    'Created_By'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('Preorders_Sold');
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return sheet;
}

// ============================================================================
// UTILITY HELPERS
// ============================================================================

/**
 * Finds the index of a header from an array of synonyms
 * NOTE: This finds the first COLUMN that matches ANY synonym (column-order priority)
 * @param {Array<string>} headers - The header row from the sheet
 * @param {Array<string>} synonyms - Array of possible header names to match
 * @return {number} Column index (0-based) or -1 if not found
 */
function findHeaderIndex(headers, synonyms) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').trim().toLowerCase();
    for (const synonym of synonyms) {
      if (header === synonym.toLowerCase()) {
        return i;
      }
    }
  }
  return -1;
}

/**
 * Finds the first matching column from a prioritized list of header names
 * NOTE: This finds the first SYNONYM that exists in headers (synonym-order priority)
 * Use this when you need fallback logic (e.g., prefer Preorder_Price, fallback to Unit_Cost)
 * @param {Array<string>} headers - The header row from the sheet
 * @param {Array<string>} prioritizedNames - Array of header names in priority order
 * @return {number} Column index (0-based) or -1 if not found
 */
function findHeaderIndexByPriority(headers, prioritizedNames) {
  const normalizedHeaders = headers.map(h => String(h || '').trim().toLowerCase());
  
  for (const name of prioritizedNames) {
    const idx = normalizedHeaders.indexOf(name.toLowerCase());
    if (idx !== -1) {
      return idx;
    }
  }
  return -1;
}

/**
 * Generates a unique preorder ID
 * Format: PO-YYMMDD-XXXXX (where X is alphanumeric)
 * @return {string} Unique preorder ID
 * @private
 */
function generatePreorderId_() {
  const now = new Date();
  const datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyMMdd');
  const randomPart = Math.random().toString(36).substring(2, 7).toUpperCase();
  return `PO-${datePart}-${randomPart}`;
}

/**
 * Searches customers by name/preferred name
 * Searches Key_Tracker and BP_Total sheets for matching customers
 * @param {string} query - Search query (name or partial name)
 * @return {Array<Object>} Array of matching customers
 */
function searchCustomers(query) {
  if (!query || query.trim().length < 2) {
    return [];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queryLower = query.toLowerCase().trim();
  const results = [];
  const seenNames = new Set();

  // Search Key_Tracker for customers
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet && keySheet.getLastRow() > 1) {
    const keyData = keySheet.getDataRange().getValues();
    const keyHeaders = keyData[0];
    const nameCol = findHeaderIndex(keyHeaders, ['PreferredName', 'Preferred_Name', 'Name', 'Player_Name', 'Player']);
    const contactCol = findHeaderIndex(keyHeaders, ['Contact_Info', 'ContactInfo', 'Contact', 'Email', 'Phone']);

    if (nameCol !== -1) {
      for (let i = 1; i < keyData.length; i++) {
        const name = String(keyData[i][nameCol] || '').trim();
        const nameLower = name.toLowerCase();

        if (name && nameLower.includes(queryLower) && !seenNames.has(nameLower)) {
          seenNames.add(nameLower);
          results.push({
            preferredName: name,
            contactInfo: contactCol !== -1 ? String(keyData[i][contactCol] || '') : '',
            source: 'Key_Tracker'
          });
        }
      }
    }
  }

  // Search BP_Total for customers
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet && bpSheet.getLastRow() > 1) {
    const bpData = bpSheet.getDataRange().getValues();
    const bpHeaders = bpData[0];
    const nameCol = findHeaderIndex(bpHeaders, ['PreferredName', 'Preferred_Name', 'Name', 'Player_Name', 'Player']);

    if (nameCol !== -1) {
      for (let i = 1; i < bpData.length; i++) {
        const name = String(bpData[i][nameCol] || '').trim();
        const nameLower = name.toLowerCase();

        if (name && nameLower.includes(queryLower) && !seenNames.has(nameLower)) {
          seenNames.add(nameLower);
          results.push({
            preferredName: name,
            contactInfo: '',
            source: 'BP_Total'
          });
        }
      }
    }
  }

  // Sort by name
  results.sort((a, b) => a.preferredName.localeCompare(b.preferredName));

  return results.slice(0, 20); // Limit to 20 results
}


// ============================================================================
// MAIN API FUNCTIONS
// ============================================================================

/**
 * Gets preorder buckets for UI display
 * Returns an object shaped exactly the way the HTML expects:
 * {
 *   sets: [ 'Set A', 'Set B', ... ],
 *   items: [
 *     {
 *       Set_Name,
 *       Item_Name,
 *       Item_Code,
 *       Unit_Price,
 *       Qty_Remaining,
 *       Quantity_Total,
 *       Reserved,
 *       Release_Date,
 *       Notes
 *     },
 *     ...
 *   ],
 *   summary: {
 *     [setName]: { totalAllocated, totalRemaining }
 *   }
 * }
 */
function getPreorderBucketsForUI() {
  ensurePreordersBucketsSchema();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Buckets');

  // If nothing there, return empty structure (so UI doesn't explode)
  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      sets: [],
      items: [],
      summary: {}
    };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const setNameCol   = findHeaderIndex(headers, ['Set_Name', 'SetName', 'Set']);
  const itemNameCol  = findHeaderIndex(headers, ['Item_Name', 'ItemName', 'Name', 'Product_Name', 'Product']);
  const itemCodeCol  = findHeaderIndex(headers, ['Item_Code', 'ItemCode', 'Code', 'SKU']);
  const quantityCol  = findHeaderIndex(headers, ['Quantity', 'Qty', 'Total_Qty']);
  const reservedCol  = findHeaderIndex(headers, ['Reserved']);
  const availableCol = findHeaderIndex(headers, ['Available']);
  const releaseDateCol = findHeaderIndex(headers, ['Release_Date', 'ReleaseDate']);
  const notesCol     = findHeaderIndex(headers, ['Notes']);

  // Price: prefer Preorder_Price, then Unit_Price, then Unit_Cost, then Retail_Price
  // FIXED: Use findHeaderIndexByPriority so it checks synonyms in order, not columns
  const unitPriceCol = findHeaderIndexByPriority(
    headers,
    ['Preorder_Price', 'PreorderPrice', 'Unit_Price', 'UnitPrice', 'Unit_Cost', 'UnitCost', 'Retail_Price', 'RetailPrice']
  );

  const setsSet = new Set();
  const items = [];
  const summary = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const setName  = setNameCol   !== -1 ? String(row[setNameCol]   || '').trim() : '';
    const itemName = itemNameCol  !== -1 ? String(row[itemNameCol]  || '').trim() : '';
    const itemCode = itemCodeCol  !== -1 ? String(row[itemCodeCol]  || '').trim() : '';

    // Skip completely empty rows
    if (!setName && !itemName && !itemCode) continue;

    const quantity = quantityCol  !== -1 ? coerceNumber(row[quantityCol], 0) : 0;
    const reserved = reservedCol  !== -1 ? coerceNumber(row[reservedCol], 0) : 0;

    let available = availableCol !== -1 ? coerceNumber(row[availableCol], 0) : (quantity - reserved);

    // Recalculate available if explicit column is blank
    if (availableCol === -1 || row[availableCol] === '' || row[availableCol] == null) {
      available = Math.max(0, quantity - reserved);
    }

    const unitPrice   = unitPriceCol   !== -1 ? coerceNumber(row[unitPriceCol], 0)   : 0;
    const releaseDate = releaseDateCol !== -1 ? row[releaseDateCol] : '';
    const notes       = notesCol       !== -1 ? row[notesCol]       : '';

    if (setName) {
      setsSet.add(setName);

      // Build summary by set
      if (!summary[setName]) {
        summary[setName] = { totalAllocated: 0, totalRemaining: 0 };
      }
      summary[setName].totalAllocated += quantity;
      summary[setName].totalRemaining += available;
    }

    // Shape the item exactly how the HTML expects
    items.push({
      Set_Name      : setName,
      Item_Name     : itemName,
      Item_Code     : itemCode,
      Unit_Price    : unitPrice,
      Qty_Remaining : available,
      Quantity_Total: quantity,
      Reserved      : reserved,
      Release_Date  : releaseDate,
      Notes         : notes
    });
  }

  return {
    sets   : Array.from(setsSet).filter(s => s).sort(),
    items  : items,
    summary: summary
  };
}


/**
 * Sells a preorder (creates entries in Preorders_Sold and decrements Preorders_Buckets)
 * FIXED: Handles both qty and qtyWanted from the basket
 * @param {Object} payload - The preorder payload
 * @param {string} payload.customerName - Customer's preferred name
 * @param {string} payload.contactInfo - Customer contact info
 * @param {Array<Object>} payload.basket - Array of basket items
 * @param {number} payload.totalDue - Total amount due
 * @param {number} payload.depositAmount - Deposit paid
 * @param {string} payload.targetPayoffDate - Target payoff date
 * @param {string} payload.notes - Additional notes
 * @return {Object} Result with status, preorderId, message, etc.
 */
function sellPreorder(payload) {
  try {
    // Validate payload
    if (!payload.customerName || !payload.customerName.trim()) {
      return { status: 'ERROR', message: 'Customer name is required' };
    }

    if (!payload.basket || payload.basket.length === 0) {
      return { status: 'ERROR', message: 'Basket cannot be empty' };
    }

    ensurePreordersSoldSchema();
    ensurePreordersBucketsSchema();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const soldSheet = ss.getSheetByName('Preorders_Sold');

    // Generate preorder ID
    const preorderId = generatePreorderId_();
    const now = dateISO();

    // Calculate totals
    const totalDue = coerceNumber(payload.totalDue, 0);
    const depositPaid = coerceNumber(payload.depositAmount, 0);
    const balanceDue = Math.max(0, totalDue - depositPaid);

    // Determine status
    let status = 'Pending';
    if (depositPaid > 0 && balanceDue > 0) {
      status = 'Deposit_Paid';
    } else if (depositPaid >= totalDue && totalDue > 0) {
      status = 'Paid_In_Full';
    }

    // Build rows for Preorders_Sold
    // Headers: Preorder_ID, PreferredName, Contact_Info, Set_Name, Item_Name, Item_Code,
    //          Qty, Unit_Price, Line_Total, Total_Due, Deposit_Paid, Balance_Due,
    //          Target_Payoff, Status, Notes, Created_At, Created_By
    const rows = payload.basket.map((item, idx) => {
      // Support both item.qty and item.qtyWanted from the UI
      const qty = coerceNumber(
        item.qty != null ? item.qty : item.qtyWanted,
        0
      );
      const lineTotal = qty * coerceNumber(item.unitPrice, 0);

      return [
        preorderId,
        payload.customerName.trim(),
        payload.contactInfo || '',
        item.setName || '',
        item.itemName || '',
        item.itemCode || '',
        qty,
        coerceNumber(item.unitPrice, 0),
        lineTotal,
        idx === 0 ? totalDue : '',     // Only first row gets total_due
        idx === 0 ? depositPaid : '',  // Only first row gets deposit_paid
        idx === 0 ? balanceDue : '',   // Only first row gets balance_due
        payload.targetPayoffDate || '',
        status,
        payload.notes || '',
        now,
        Session.getActiveUser().getEmail() || 'system'
      ];
    });

    // Write to Preorders_Sold
    const startRow = soldSheet.getLastRow() + 1;
    soldSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // Decrement bucket quantities
    decrementBucketQuantities(ss, payload.basket);

    // Log integrity action (if function exists)
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('PREORDER_CREATED', {
        preferredName: payload.customerName,
        details: `Preorder ${preorderId}: ${payload.basket.length} items, Total: ${formatCurrency(totalDue)}, Deposit: ${formatCurrency(depositPaid)}`,
        status: 'SUCCESS'
      });
    }

    return {
      status: 'OK',
      preorderId: preorderId,
      customerName: payload.customerName,
      itemCount: payload.basket.length,
      totalDue: totalDue,
      balance: balanceDue,
      message: `Preorder ${preorderId} created successfully. Balance due: ${formatCurrency(balanceDue)}`
    };

  } catch (e) {
    console.error('sellPreorder error:', e);
    return { status: 'ERROR', message: e.message };
  }
}


/**
 * Decrements bucket quantities for sold items
 * FIXED: Handles both qty and qtyWanted from basket
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {Array<Object>} basket - Array of basket items
 */
function decrementBucketQuantities(ss, basket) {
  const sheet = ss.getSheetByName('Preorders_Buckets');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const setNameCol = findHeaderIndex(headers, ['Set_Name', 'SetName', 'Set']);
  const itemNameCol = findHeaderIndex(headers, ['Item_Name', 'ItemName', 'Name', 'Product_Name']);
  const itemCodeCol = findHeaderIndex(headers, ['Item_Code', 'ItemCode', 'Code', 'SKU']);
  const quantityCol = findHeaderIndex(headers, ['Quantity', 'Qty']);
  const reservedCol = findHeaderIndex(headers, ['Reserved']);
  const availableCol = findHeaderIndex(headers, ['Available']);
  const updatedCol = findHeaderIndex(headers, ['LastUpdated']);

  for (const item of basket) {
    // Support both qty and qtyWanted
    const itemQty = coerceNumber(
      item.qty != null ? item.qty : item.qtyWanted,
      0
    );
    if (itemQty <= 0) continue;

    // Find matching bucket row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowSetName = String(row[setNameCol] || '').toLowerCase().trim();
      const rowItemName = String(row[itemNameCol] || '').toLowerCase().trim();
      const rowItemCode = String(row[itemCodeCol] || '').toLowerCase().trim();

      const itemSetName = String(item.setName || '').toLowerCase().trim();
      const itemItemName = String(item.itemName || '').toLowerCase().trim();
      const itemItemCode = String(item.itemCode || '').toLowerCase().trim();

      // Match on setName + (itemCode OR itemName)
      const setMatch = !itemSetName || rowSetName === itemSetName;
      const codeMatch = itemItemCode && rowItemCode === itemItemCode;
      const nameMatch = itemItemName && rowItemName === itemItemName;

      if (setMatch && (codeMatch || nameMatch)) {
        // Update Reserved and Available
        const currentReserved = coerceNumber(row[reservedCol], 0);
        const currentQty = coerceNumber(row[quantityCol], 0);

        const newReserved = currentReserved + itemQty;
        const newAvailable = Math.max(0, currentQty - newReserved);

        if (reservedCol !== -1) {
          sheet.getRange(i + 1, reservedCol + 1).setValue(newReserved);
        }

        if (availableCol !== -1) {
          sheet.getRange(i + 1, availableCol + 1).setValue(newAvailable);
        }

        if (updatedCol !== -1) {
          sheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());
        }

        break;
      }
    }
  }
}


/**
 * Increments bucket quantities (for cancellation/returns)
 * FIXED: Handles both qty and qtyWanted
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {Array<Object>} items - Array of items to return
 */
function incrementBucketQuantities(ss, items) {
  const sheet = ss.getSheetByName('Preorders_Buckets');
  if (!sheet || sheet.getLastRow() <= 1) return;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const setNameCol = findHeaderIndex(headers, ['Set_Name', 'SetName', 'Set']);
  const itemNameCol = findHeaderIndex(headers, ['Item_Name', 'ItemName', 'Name', 'Product_Name']);
  const itemCodeCol = findHeaderIndex(headers, ['Item_Code', 'ItemCode', 'Code', 'SKU']);
  const quantityCol = findHeaderIndex(headers, ['Quantity', 'Qty']);
  const reservedCol = findHeaderIndex(headers, ['Reserved']);
  const availableCol = findHeaderIndex(headers, ['Available']);
  const updatedCol = findHeaderIndex(headers, ['LastUpdated']);

  for (const item of items) {
    // Support both qty and qtyWanted
    const itemQty = coerceNumber(
      item.qty != null ? item.qty : item.qtyWanted,
      0
    );
    if (itemQty <= 0) continue;

    // Find matching bucket row
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowSetName = String(row[setNameCol] || '').toLowerCase().trim();
      const rowItemName = String(row[itemNameCol] || '').toLowerCase().trim();
      const rowItemCode = String(row[itemCodeCol] || '').toLowerCase().trim();

      const itemSetName = String(item.setName || '').toLowerCase().trim();
      const itemItemName = String(item.itemName || '').toLowerCase().trim();
      const itemItemCode = String(item.itemCode || '').toLowerCase().trim();

      // Match on setName + (itemCode OR itemName)
      const setMatch = !itemSetName || rowSetName === itemSetName;
      const codeMatch = itemItemCode && rowItemCode === itemItemCode;
      const nameMatch = itemItemName && rowItemName === itemItemName;

      if (setMatch && (codeMatch || nameMatch)) {
        // Update Reserved and Available
        const currentReserved = coerceNumber(row[reservedCol], 0);
        const currentQty = coerceNumber(row[quantityCol], 0);

        const newReserved = Math.max(0, currentReserved - itemQty);
        const newAvailable = Math.max(0, currentQty - newReserved);

        if (reservedCol !== -1) {
          sheet.getRange(i + 1, reservedCol + 1).setValue(newReserved);
        }

        if (availableCol !== -1) {
          sheet.getRange(i + 1, availableCol + 1).setValue(newAvailable);
        }

        if (updatedCol !== -1) {
          sheet.getRange(i + 1, updatedCol + 1).setValue(dateISO());
        }

        break;
      }
    }
  }
}


/**
 * Enhanced customer search - UI calls this function name
 * Just wraps searchCustomers()
 * @param {string} query - Search query
 * @return {Array<Object>} Array of matching customers
 */
function searchCustomersWithPreferred(query) {
  return searchCustomers(query);
}


// ============================================================================
// ADDITIONAL PREORDER MANAGEMENT FUNCTIONS
// ============================================================================

/**
 * Gets all preorders for a customer
 * @param {string} customerName - Customer's preferred name
 * @return {Array<Object>} Array of preorder objects
 */
function getPreordersForCustomer(customerName) {
  if (!customerName || !customerName.trim()) {
    return [];
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Sold');

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nameCol = findHeaderIndex(headers, ['PreferredName', 'Preferred_Name', 'Customer_Name', 'Name']);
  if (nameCol === -1) return [];

  const customerLower = customerName.toLowerCase().trim();
  const preorders = [];

  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][nameCol] || '').toLowerCase().trim();
    if (rowName === customerLower) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      preorders.push(obj);
    }
  }

  return preorders;
}

/**
 * Gets a single preorder by ID
 * @param {string} preorderId - The preorder ID
 * @return {Object|null} Preorder object or null
 */
function getPreorderById(preorderId) {
  if (!preorderId) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Sold');

  if (!sheet || sheet.getLastRow() <= 1) {
    return null;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = findHeaderIndex(headers, ['Preorder_ID', 'PreorderID', 'ID']);
  if (idCol === -1) return null;

  const items = [];
  let foundPreorder = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === preorderId) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      items.push(obj);
    }
  }

  if (items.length > 0) {
    // Return first item with all details, plus items array
    foundPreorder = {
      ...items[0],
      items: items
    };
  }

  return foundPreorder;
}

/**
 * Updates the status of a preorder
 * @param {string} preorderId - The preorder ID
 * @param {string} newStatus - New status value
 * @return {Object} Result with status and message
 */
function updatePreorderStatus(preorderId, newStatus) {
  if (!preorderId || !newStatus) {
    return { status: 'ERROR', message: 'Preorder ID and new status are required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Sold');

  if (!sheet || sheet.getLastRow() <= 1) {
    return { status: 'ERROR', message: 'Preorders_Sold sheet not found or empty' };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = findHeaderIndex(headers, ['Preorder_ID', 'PreorderID', 'ID']);
  const statusCol = findHeaderIndex(headers, ['Status']);

  if (idCol === -1 || statusCol === -1) {
    return { status: 'ERROR', message: 'Required columns not found' };
  }

  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === preorderId) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      updatedCount++;
    }
  }

  if (updatedCount === 0) {
    return { status: 'ERROR', message: `Preorder ${preorderId} not found` };
  }

  if (typeof logIntegrityAction === 'function') {
    logIntegrityAction('PREORDER_STATUS_UPDATE', {
      details: `Preorder ${preorderId} status updated to ${newStatus}`,
      status: 'SUCCESS'
    });
  }

  return {
    status: 'OK',
    message: `Preorder ${preorderId} updated to ${newStatus}`,
    updatedRows: updatedCount
  };
}

/**
 * Cancels a preorder and restores inventory
 * @param {string} preorderId - The preorder ID to cancel
 * @param {string} reason - Cancellation reason
 * @return {Object} Result with status and message
 */
function cancelPreorder(preorderId, reason) {
  if (!preorderId) {
    return { status: 'ERROR', message: 'Preorder ID is required' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const preorder = getPreorderById(preorderId);

  if (!preorder) {
    return { status: 'ERROR', message: `Preorder ${preorderId} not found` };
  }

  if (preorder.Status === 'Cancelled') {
    return { status: 'ERROR', message: `Preorder ${preorderId} is already cancelled` };
  }

  // Build items array for inventory restoration
  const itemsToRestore = preorder.items.map(item => ({
    setName: item.Set_Name,
    itemName: item.Item_Name,
    itemCode: item.Item_Code,
    qty: coerceNumber(item.Qty, 0)
  }));

  // Restore inventory
  incrementBucketQuantities(ss, itemsToRestore);

  // Update status to Cancelled
  const result = updatePreorderStatus(preorderId, 'Cancelled');

  if (result.status === 'OK') {
    if (typeof logIntegrityAction === 'function') {
      logIntegrityAction('PREORDER_CANCELLED', {
        preferredName: preorder.PreferredName,
        details: `Preorder ${preorderId} cancelled. Reason: ${reason || 'Not specified'}. Inventory restored.`,
        status: 'SUCCESS'
      });
    }

    return {
      status: 'OK',
      message: `Preorder ${preorderId} cancelled successfully. Inventory has been restored.`,
      itemsRestored: itemsToRestore.length
    };
  }

  return result;
}