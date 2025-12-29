/**
 * Daily Close Checklist Service
 * @fileoverview Handles daily close checklist submission and logging
 */

// ============================================================================
// DAILY CLOSE CHECKLIST SERVICE
// ============================================================================

/**
 * Logs a daily close checklist submission.
 * @param {Object} payload
 *  - staffInitials: string (required)
 *  - completedItems: string[] - labels of checked items
 *  - notes: string - optional notes
 * @return {Object} result with success status and metadata
 */
function submitDailyCloseChecklist(payload) {
  // Validate payload
  if (!payload) {
    throw new Error('Payload is required.');
  }

  const staffInitials = (payload.staffInitials || '').trim();
  if (!staffInitials) {
    throw new Error('Staff initials are required.');
  }

  // Normalize completedItems to array
  const completedItems = Array.isArray(payload.completedItems)
    ? payload.completedItems
    : [];

  const notes = (payload.notes || '').trim();

  // Ensure Daily_Close_Log sheet exists
  const sheet = getOrCreateDailyCloseLog_();

  // Prepare row data
  const now = new Date();
  const store = 'Main'; // Hardcoded for now

  let createdBy;
  try {
    createdBy = Session.getActiveUser().getEmail() || 'Unknown';
  } catch (e) {
    createdBy = 'Unknown';
  }

  const row = [
    now,                                    // Date
    store,                                  // Store
    staffInitials,                          // Staff_Initials
    JSON.stringify(completedItems),         // Completed_Items (JSON)
    notes,                                  // Notes
    now,                                    // Created_At
    createdBy                               // Created_By
  ];

  // Append row
  try {
    sheet.appendRow(row);
  } catch (e) {
    throw new Error('Failed to save checklist: ' + e.message);
  }

  // Log to Integrity_Log for audit trail
  try {
    logIntegrityAction('DAILY_CLOSE', {
      details: `Staff: ${staffInitials} | Items: ${completedItems.length} | Notes: ${notes.length > 0 ? 'Yes' : 'No'}`,
      status: 'SUCCESS'
    });
  } catch (e) {
    // Don't fail if integrity logging fails
    console.error('Failed to log daily close action:', e);
  }

  return {
    success: true,
    loggedAt: now,
    itemsCount: completedItems.length
  };
}

// ============================================================================
// PRIVATE HELPERS
// ============================================================================

/**
 * Gets or creates the Daily_Close_Log sheet with proper headers.
 * @return {Sheet} The Daily_Close_Log sheet
 * @private
 */
function getOrCreateDailyCloseLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Daily_Close_Log');

  if (!sheet) {
    sheet = ss.insertSheet('Daily_Close_Log');

    const headers = [
      'Date',
      'Store',
      'Staff_Initials',
      'Completed_Items',
      'Notes',
      'Created_At',
      'Created_By'
    ];

    sheet.appendRow(headers);
    sheet.setFrozenRows(1);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f3f3f3');

    // Set column widths for readability
    sheet.setColumnWidth(1, 100);  // Date
    sheet.setColumnWidth(2, 60);   // Store
    sheet.setColumnWidth(3, 100);  // Staff_Initials
    sheet.setColumnWidth(4, 350);  // Completed_Items
    sheet.setColumnWidth(5, 200);  // Notes
    sheet.setColumnWidth(6, 150);  // Created_At
    sheet.setColumnWidth(7, 180);  // Created_By
  }

  return sheet;
}