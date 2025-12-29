/**
 * Preorder Cancel Service - Cancel Preorder Operations
 * @fileoverview Handles preorder cancellation and bucket adjustment
 */

// ============================================================================
// GET CANCELABLE PREORDERS
// ============================================================================

/**
 * Gets list of preorders eligible for cancellation
 * A preorder is cancellable if Status is not "Cancelled" and not "Picked Up" (case-insensitive)
 * @return {Array<Object>} List of cancellable preorder objects
 */
function getCancelablePreorders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Sold');

  if (!sheet) {
    throw new Error('Preorders_Sold sheet not found. Please create the sheet first.');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return []; // Only header or empty
  }

  const headers = data[0];

  // Find required columns
  const preorderIdCol = headers.indexOf('Preorder_ID');
  const preferredNameCol = headers.indexOf('preferred_name_id');
  const qtyCol = headers.indexOf('Qty');
  const statusCol = headers.indexOf('Status');

  // Validate required columns exist
  if (preorderIdCol === -1) {
    throw new Error('Required column "Preorder_ID" not found in Preorders_Sold sheet.');
  }
  if (preferredNameCol === -1) {
    throw new Error('Required column "preferred_name_id" not found in Preorders_Sold sheet.');
  }
  if (qtyCol === -1) {
    throw new Error('Required column "Qty" not found in Preorders_Sold sheet.');
  }
  if (statusCol === -1) {
    throw new Error('Required column "Status" not found in Preorders_Sold sheet.');
  }

  // Find optional columns (use defaults if missing)
  const balanceDueCol = headers.indexOf('Balance_Due');
  const setNameCol = headers.indexOf('Set_Name');
  const itemNameCol = headers.indexOf('Item_Name');
  const itemCodeCol = headers.indexOf('Item_Code');

  const cancelablePreorders = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[statusCol] || '').toLowerCase().trim();

    // Skip if already Cancelled or Picked Up
    if (status === 'cancelled' || status === 'picked up') {
      continue;
    }

    // Skip if no preorder ID
    const preorderId = String(row[preorderIdCol] || '').trim();
    if (!preorderId) {
      continue;
    }

    // Build preorder object
    const preorder = {
      preorderId: preorderId,
      preferredName: String(row[preferredNameCol] || ''),
      setName: setNameCol !== -1 ? String(row[setNameCol] || '') : '',
      itemName: itemNameCol !== -1 ? String(row[itemNameCol] || '') : '',
      itemCode: itemCodeCol !== -1 ? String(row[itemCodeCol] || '') : '',
      qty: parseInt(row[qtyCol], 10) || 0,
      status: String(row[statusCol] || ''),
      balanceDue: balanceDueCol !== -1 ? (parseFloat(row[balanceDueCol]) || 0) : 0,
      rowNumber: i + 1 // 1-indexed row number
    };

    cancelablePreorders.push(preorder);
  }

  return cancelablePreorders;
}

// ============================================================================
// CANCEL PREORDER
// ============================================================================

/**
 * Cancels a preorder and returns its quantity to the appropriate bucket
 * @param {string} preorderId - The Preorder_ID to cancel
 * @param {string} staffInitials - Staff initials (optional but recommended)
 * @param {string} reason - Cancellation reason (optional)
 * @return {Object} Result object with success status and details
 */
function cancelPreorder(preorderId, staffInitials, reason) {
  // Validate required input
  if (!preorderId || String(preorderId).trim() === '') {
    throw new Error('Preorder ID is required.');
  }

  preorderId = String(preorderId).trim();
  staffInitials = staffInitials ? String(staffInitials).trim() : '';
  reason = reason ? String(reason).trim() : '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // -------------------------------------------------------------------------
  // Step 1: Find and validate the preorder in Preorders_Sold
  // -------------------------------------------------------------------------
  const preordersSheet = ss.getSheetByName('Preorders_Sold');
  if (!preordersSheet) {
    throw new Error('Preorders_Sold sheet not found.');
  }

  let data = preordersSheet.getDataRange().getValues();
  let headers = data[0];

  // Find required columns
  const preorderIdCol = headers.indexOf('Preorder_ID');
  const qtyCol = headers.indexOf('Qty');
  const statusCol = headers.indexOf('Status');
  const setNameCol = headers.indexOf('Set_Name');
  const itemCodeCol = headers.indexOf('Item_Code');
  const itemNameCol = headers.indexOf('Item_Name');
  const balanceDueCol = headers.indexOf('Balance_Due');

  if (preorderIdCol === -1 || qtyCol === -1 || statusCol === -1) {
    throw new Error('Required columns (Preorder_ID, Qty, Status) not found in Preorders_Sold.');
  }

  // Find the preorder row
  let preorderRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][preorderIdCol]).trim() === preorderId) {
      preorderRowIndex = i;
      break;
    }
  }

  if (preorderRowIndex === -1) {
    throw new Error(`Preorder not found for ID: ${preorderId}`);
  }

  const preorderRow = data[preorderRowIndex];
  const qty = parseInt(preorderRow[qtyCol], 10) || 0;
  const setName = setNameCol !== -1 ? String(preorderRow[setNameCol] || '').trim() : '';
  const itemCode = itemCodeCol !== -1 ? String(preorderRow[itemCodeCol] || '').trim() : '';
  const itemName = itemNameCol !== -1 ? String(preorderRow[itemNameCol] || '').trim() : '';
  const currentStatus = String(preorderRow[statusCol] || '').toLowerCase().trim();
  const balanceDue = balanceDueCol !== -1 ? (parseFloat(preorderRow[balanceDueCol]) || 0) : 0;

  // Check if already cancelled
  if (currentStatus === 'cancelled') {
    throw new Error(`Preorder ${preorderId} is already cancelled.`);
  }

  // Check if already picked up
  if (currentStatus === 'picked up') {
    throw new Error(`Preorder ${preorderId} has already been picked up and cannot be cancelled.`);
  }

  // -------------------------------------------------------------------------
  // Step 2: Ensure cancellation metadata columns exist
  // -------------------------------------------------------------------------
  let cancelledDateCol = headers.indexOf('Cancelled_Date');
  let cancelledByCol = headers.indexOf('Cancelled_By');
  let cancelReasonCol = headers.indexOf('Cancel_Reason');

  let headersModified = false;

  if (cancelledDateCol === -1) {
    cancelledDateCol = headers.length;
    preordersSheet.getRange(1, cancelledDateCol + 1).setValue('Cancelled_Date');
    headers.push('Cancelled_Date');
    headersModified = true;
  }

  if (cancelledByCol === -1) {
    cancelledByCol = headers.length;
    preordersSheet.getRange(1, cancelledByCol + 1).setValue('Cancelled_By');
    headers.push('Cancelled_By');
    headersModified = true;
  }

  if (cancelReasonCol === -1) {
    cancelReasonCol = headers.length;
    preordersSheet.getRange(1, cancelReasonCol + 1).setValue('Cancel_Reason');
    headers.push('Cancel_Reason');
    headersModified = true;
  }

  // Re-read data if headers were modified
  if (headersModified) {
    SpreadsheetApp.flush();
    data = preordersSheet.getDataRange().getValues();
  }

  // -------------------------------------------------------------------------
  // Step 3: Update the preorder row
  // -------------------------------------------------------------------------
  const preorderRowNumber = preorderRowIndex + 1; // 1-indexed

  // Update Status to "Cancelled"
  preordersSheet.getRange(preorderRowNumber, statusCol + 1).setValue('Cancelled');

  // Update cancellation metadata
  preordersSheet.getRange(preorderRowNumber, cancelledDateCol + 1).setValue(new Date());
  preordersSheet.getRange(preorderRowNumber, cancelledByCol + 1).setValue(staffInitials);
  preordersSheet.getRange(preorderRowNumber, cancelReasonCol + 1).setValue(reason);

  // -------------------------------------------------------------------------
  // Step 4: Adjust preorder_buckets
  // -------------------------------------------------------------------------
  let bucketMatched = false;
  let bucketRow = null;

  const bucketsSheet = ss.getSheetByName('preorder_buckets');

  if (bucketsSheet && qty > 0) {
    const bucketResult = adjustPreorderBucket_(bucketsSheet, setName, itemCode, itemName, qty);
    bucketMatched = bucketResult.matched;
    bucketRow = bucketResult.row;
  }

  // -------------------------------------------------------------------------
  // Step 5: Log the action
  // -------------------------------------------------------------------------
  try {
    logIntegrityAction('PREORDER_CANCEL', {
      details: `Cancelled preorder ${preorderId}. Qty: ${qty}. Bucket adjusted: ${bucketMatched}. Staff: ${staffInitials || 'N/A'}. Reason: ${reason || 'N/A'}`,
      status: 'SUCCESS'
    });
  } catch (e) {
    // Don't fail the cancel if logging fails
    console.error('Failed to log preorder cancel action:', e);
  }

  // -------------------------------------------------------------------------
  // Step 6: Return result
  // -------------------------------------------------------------------------
  const result = {
    success: true,
    preorderId: preorderId,
    row: preorderRowNumber,
    qtyReturned: qty,
    bucketMatched: bucketMatched,
    bucketRow: bucketRow
  };

  // Add warning if bucket not found
  if (!bucketMatched && qty > 0) {
    result.warning = 'No matching bucket found in preorder_buckets. Inventory allocation was not restored.';
  }

  return result;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Adjusts the preorder bucket by returning cancelled quantity
 * @param {Sheet} bucketsSheet - The preorder_buckets sheet
 * @param {string} setName - Set name to match
 * @param {string} itemCode - Item code to match (primary)
 * @param {string} itemName - Item name to match (fallback)
 * @param {number} qty - Quantity to return to bucket
 * @return {Object} {matched: boolean, row: number|null}
 * @private
 */
function adjustPreorderBucket_(bucketsSheet, setName, itemCode, itemName, qty) {
  const data = bucketsSheet.getDataRange().getValues();
  const headers = data[0];

  // Find required bucket columns
  const bucketSetNameCol = headers.indexOf('Set_Name');
  const bucketItemNameCol = headers.indexOf('Item_Name');
  const bucketItemCodeCol = headers.indexOf('Item_Code');
  const reservedCol = headers.indexOf('Reserved');
  const availableCol = headers.indexOf('Available');

  // Validate required columns
  if (bucketSetNameCol === -1 || bucketItemNameCol === -1 || bucketItemCodeCol === -1 ||
      reservedCol === -1 || availableCol === -1) {
    // Missing required columns - can't adjust bucket
    return { matched: false, row: null };
  }

  // Find matching bucket row
  let matchedRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const bucketSetName = String(data[i][bucketSetNameCol] || '').trim();
    const bucketItemCode = String(data[i][bucketItemCodeCol] || '').trim();
    const bucketItemName = String(data[i][bucketItemNameCol] || '').trim();

    // Primary match: Set_Name AND Item_Code
    if (setName && itemCode && bucketSetName.toLowerCase() === setName.toLowerCase() &&
        bucketItemCode.toLowerCase() === itemCode.toLowerCase()) {
      matchedRowIndex = i;
      break;
    }

    // Fallback match: Set_Name AND Item_Name (if Item_Code is missing/blank)
    if (setName && (!itemCode || itemCode === '') && itemName &&
        bucketSetName.toLowerCase() === setName.toLowerCase() &&
        bucketItemName.toLowerCase() === itemName.toLowerCase()) {
      matchedRowIndex = i;
      break;
    }
  }

  if (matchedRowIndex === -1) {
    return { matched: false, row: null };
  }

  // Adjust Reserved and Available
  const bucketRowNumber = matchedRowIndex + 1; // 1-indexed
  const currentReserved = parseFloat(data[matchedRowIndex][reservedCol]) || 0;
  const currentAvailable = parseFloat(data[matchedRowIndex][availableCol]) || 0;

  // Reserved = Reserved - qty (but never below 0)
  const newReserved = Math.max(0, currentReserved - qty);

  // Available = Available + qty
  const newAvailable = currentAvailable + qty;

  // Write updated values
  bucketsSheet.getRange(bucketRowNumber, reservedCol + 1).setValue(newReserved);
  bucketsSheet.getRange(bucketRowNumber, availableCol + 1).setValue(newAvailable);

  // Update LastUpdated column if it exists
  const lastUpdatedCol = headers.indexOf('LastUpdated');
  if (lastUpdatedCol !== -1) {
    bucketsSheet.getRange(bucketRowNumber, lastUpdatedCol + 1).setValue(new Date());
  }

  return { matched: true, row: bucketRowNumber };
}