/**
 * Preorder Status Service
 * @fileoverview Backend service for View Preorder Status dashboard
 */

// ============================================================================
// PREORDER STATUS SERVICE
// ============================================================================

/**
 * Returns preorders grouped by status buckets for the status sidebar.
 * @param {Object} filters (optional)
 *   - search: string | null   (search text; matches name, set, item)
 *   - includeCancelled: boolean (default false)
 *   - sortBy: string | null ('Created_At' | 'Target_Payoff' | 'Balance_Due')
 *   - sortDir: 'asc' | 'desc' (default 'asc')
 * @return {Object} {
 *   pendingPayment: Array<Object>,
 *   readyForPickup: Array<Object>,
 *   completed: Array<Object>,
 *   summary: {
 *     pendingPaymentCount: number,
 *     readyForPickupCount: number,
 *     completedCount: number,
 *     totalBalanceDue: number
 *   }
 * }
 */
function getPreordersByStatus(filters) {
  filters = filters || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preorders_Sold');

  if (!sheet) {
    throw new Error('Preorders_Sold sheet not found');
  }

  const data = sheet.getDataRange().getValues();

  if (data.length === 0) {
    return buildEmptyResult_();
  }

  // Parse headers
  const headers = data[0];
  const headerIndex = buildHeaderIndex_(headers);

  // Parse all rows into preorder objects
  const allPreorders = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const preorder = parsePreorderRow_(row, i + 1, headerIndex);

    // Skip empty rows (no Preorder_ID)
    if (!preorder.Preorder_ID && !preorder.preferred_name_id) {
      continue;
    }

    allPreorders.push(preorder);
  }

  // Apply search filter
  let filteredPreorders = allPreorders;
  if (filters.search && filters.search.trim()) {
    const searchLower = filters.search.trim().toLowerCase();
    filteredPreorders = allPreorders.filter(p => {
      const name = (p.preferred_name_id || '').toLowerCase();
      const setName = (p.Set_Name || '').toLowerCase();
      const itemName = (p.Item_Name || '').toLowerCase();
      return name.includes(searchLower) ||
             setName.includes(searchLower) ||
             itemName.includes(searchLower);
    });
  }

  // Group into buckets
  const pendingPayment = [];
  const readyForPickup = [];
  const completed = [];
  let totalBalanceDue = 0;

  for (const preorder of filteredPreorders) {
    const statusNorm = normalizeStatus_(preorder.Status);
    const balance = coerceNumber(preorder.Balance_Due, 0);

    const isCancelled = statusNorm === 'CANCELLED' || statusNorm === 'CANCELED';
    const isPickedUp = statusNorm === 'PICKED UP' || statusNorm === 'COMPLETED';

    // Track total balance due (for all non-cancelled with balance > 0)
    if (!isCancelled && balance > 0) {
      totalBalanceDue += balance;
    }

    // Group according to rules
    if (isCancelled) {
      // Cancelled goes to completed bucket (if includeCancelled is true)
      if (filters.includeCancelled) {
        completed.push(preorder);
      }
    } else if (balance > 0) {
      // Has balance due → pending payment
      pendingPayment.push(preorder);
    } else if (isPickedUp) {
      // Picked up or completed → completed bucket
      completed.push(preorder);
    } else {
      // Balance <= 0 and not picked up → ready for pickup
      readyForPickup.push(preorder);
    }
  }

  // Sort each bucket
  const sortBy = filters.sortBy || 'Created_At';
  const sortDir = filters.sortDir || 'asc';

  sortBucket_(pendingPayment, sortBy, sortDir);
  sortBucket_(readyForPickup, sortBy, sortDir);
  sortBucket_(completed, sortBy, sortDir);

  return {
    pendingPayment: pendingPayment,
    readyForPickup: readyForPickup,
    completed: completed,
    summary: {
      pendingPaymentCount: pendingPayment.length,
      readyForPickupCount: readyForPickup.length,
      completedCount: completed.length,
      totalBalanceDue: totalBalanceDue
    }
  };
}

// ============================================================================
// PRIVATE HELPERS
// ============================================================================

/**
 * Builds empty result object
 * @private
 */
function buildEmptyResult_() {
  return {
    pendingPayment: [],
    readyForPickup: [],
    completed: [],
    summary: {
      pendingPaymentCount: 0,
      readyForPickupCount: 0,
      completedCount: 0,
      totalBalanceDue: 0
    }
  };
}

/**
 * Builds a header-to-index mapping
 * @param {Array} headers - Header row
 * @return {Object} Map of header name to column index
 * @private
 */
function buildHeaderIndex_(headers) {
  const index = {};
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).trim();
    if (header) {
      index[header] = i;
    }
  }
  return index;
}

/**
 * Parses a single row into a preorder object
 * @param {Array} row - Row data
 * @param {number} rowNumber - Sheet row number (1-based)
 * @param {Object} headerIndex - Header-to-index mapping
 * @return {Object} Preorder object
 * @private
 */
function parsePreorderRow_(row, rowNumber, headerIndex) {
  const getValue = (headerName) => {
    const idx = headerIndex[headerName];
    return idx !== undefined ? row[idx] : undefined;
  };

  // Format Created_At for display
  let createdAt = getValue('Created_At');
  if (createdAt instanceof Date) {
    createdAt = Utilities.formatDate(createdAt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (createdAt) {
    createdAt = String(createdAt);
  }

  // Format Target_Payoff for display
  let targetPayoff = getValue('Target_Payoff');
  if (targetPayoff instanceof Date) {
    targetPayoff = Utilities.formatDate(targetPayoff, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } else if (targetPayoff) {
    targetPayoff = String(targetPayoff);
  }

  return {
    rowNumber: rowNumber,
    Preorder_ID: getValue('Preorder_ID'),
    preferred_name_id: getValue('preferred_name_id'),
    Contact_Info: getValue('Contact_Info'),
    Set_Name: getValue('Set_Name'),
    Item_Name: getValue('Item_Name'),
    Item_Code: getValue('Item_Code'),
    Qty: getValue('Qty'),
    Unit_Price: getValue('Unit_Price'),
    Total_Due: getValue('Total_Due'),
    Deposit_Paid: getValue('Deposit_Paid'),
    Balance_Due: getValue('Balance_Due'),
    Target_Payoff: targetPayoff,
    Status: getValue('Status'),
    Notes: getValue('Notes'),
    Created_At: createdAt,
    Created_By: getValue('Created_By')
  };
}

/**
 * Normalizes status string for comparison
 * @param {*} status - Raw status value
 * @return {string} Uppercase trimmed status
 * @private
 */
function normalizeStatus_(status) {
  if (!status) return '';
  return String(status).trim().toUpperCase();
}

/**
 * Sorts a bucket array by specified field
 * @param {Array} bucket - Array of preorder objects
 * @param {string} sortBy - Field to sort by
 * @param {string} sortDir - 'asc' or 'desc'
 * @private
 */
function sortBucket_(bucket, sortBy, sortDir) {
  const multiplier = sortDir === 'desc' ? -1 : 1;

  bucket.sort((a, b) => {
    let aVal = a[sortBy];
    let bVal = b[sortBy];

    // Handle missing values - push to end
    if (aVal === undefined || aVal === null || aVal === '') {
      return 1;
    }
    if (bVal === undefined || bVal === null || bVal === '') {
      return -1;
    }

    // Handle numeric fields
    if (sortBy === 'Balance_Due' || sortBy === 'Qty' || sortBy === 'Total_Due') {
      aVal = Number(aVal) || 0;
      bVal = Number(bVal) || 0;
      return (aVal - bVal) * multiplier;
    }

    // Handle date fields
    if (sortBy === 'Created_At' || sortBy === 'Target_Payoff') {
      const aDate = parseDate_(aVal);
      const bDate = parseDate_(bVal);

      if (!aDate) return 1;
      if (!bDate) return -1;

      return (aDate.getTime() - bDate.getTime()) * multiplier;
    }

    // String comparison fallback
    return String(aVal).localeCompare(String(bVal)) * multiplier;
  });
}

/**
 * Parses various date formats
 * @param {*} value - Date value (string or Date)
 * @return {Date|null} Parsed date or null
 * @private
 */
function parseDate_(value) {
  if (!value) return null;

  if (value instanceof Date) {
    return value;
  }

  // Try parsing string
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  return null;
}