/**
 * UI Handlers - Server-side functions called from HTML
 * @fileoverview Bridge between HTML UIs and backend services
 */

// ============================================================================
// EVENT UI HANDLERS
// ============================================================================

/**
 * Creates event from UI input
 * @param {Object} meta - Event metadata from UI
 * @return {Object} Result with eventId
 */
function createEventFromUI(meta) {
  return createEvent(meta);
}

/**
 * Gets integrity log for viewer
 * @return {Array<Object>} Log entries
 */
function getIntegrityLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Integrity_Log');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return toObjects(data);
}

/**
 * Gets spent pool for viewer
 * @return {Array<Object>} Spent pool entries
 */
function getSpentPool() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Spent_Pool');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return toObjects(data);
}

/**
 * Looks up player profile
 * @param {string} name - Player name
 * @return {Object} Profile object
 */
function lookupPlayerProfile(name) {
  return {
    name: name,
    bp: getPlayerBP(name),
    keys: getPlayerKeys(name) || { Red: 0, Blue: 0, Green: 0, Yellow: 0, Purple: 0, RainbowEligible: 0 }
  };
}

// ============================================================================
// PRIZE UI HANDLERS
// ============================================================================

/**
 * Previews end prizes from UI
 * @param {string} eventId - Event ID
 * @return {Object} Preview object
 */
function previewEndPrizesFromUI(eventId) {
  const preview = previewEndPrizes(eventId);

  // Store artifact
  storePreviewArtifact(eventId, preview.seed, preview.hash);

  return preview;
}

/**
 * Commits end prizes from UI
 * @param {string} eventId - Event ID
 * @param {string} previewHash - Preview hash
 * @return {Object} Commit result
 */
function commitEndPrizesFromUI(eventId, previewHash) {
  return commitEndPrizes(eventId, previewHash);
}

/**
 * Previews Commander round from UI
 * @param {string} eventId - Event ID
 * @param {number} roundId - Round number
 * @return {Object} Preview object
 */
function previewCommanderRoundFromUI(eventId, roundId) {
  return previewCommanderRound(eventId, roundId);
}

/**
 * Commits Commander round from UI
 * @param {string} eventId - Event ID
 * @param {number} roundId - Round number
 * @param {string} previewHash - Preview hash
 * @return {Object} Commit result
 */
function commitCommanderRoundFromUI(eventId, roundId, previewHash) {
  return commitCommanderRound(eventId, roundId, previewHash);
}

// ============================================================================
// ADDITIONAL HELPER FUNCTIONS
// ============================================================================

/**
 * Gets all preview artifacts (for testing/debugging)
 * @return {Array<Object>} Artifacts
 */
function getAllPreviewArtifacts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preview_Artifacts');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return toObjects(data);
}

/**
 * Creates or ensures all required sheets exist
 */
function ensureAllSheets() {
  ensureCatalogSchema();
  ensureKeyTrackerSchema();
  ensureBPTotalSchema();
  createThrottleSheet_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Integrity_Log
  if (!ss.getSheetByName('Integrity_Log')) {
    const sheet = ss.insertSheet('Integrity_Log');
    sheet.appendRow(['Timestamp', 'Event_ID', 'Action', 'Operator', 'PreferredName', 'Seed', 'Checksum_Before', 'Checksum_After', 'RL_Band', 'DF_Tags', 'Details', 'Status']);
    sheet.setFrozenRows(1);
  }

  // Spent_Pool
  if (!ss.getSheetByName('Spent_Pool')) {
    const sheet = ss.insertSheet('Spent_Pool');
    sheet.appendRow(['Event_ID', 'Item_Code', 'Item_Name', 'Level', 'Qty', 'COGS', 'Total', 'Timestamp', 'Batch_ID', 'Reverted', 'Event_Type']);
    sheet.setFrozenRows(1);
  }

  // Attendance_Missions
  if (!ss.getSheetByName('Attendance_Missions')) {
    const sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow(['PreferredName', 'Points_From_Dice_Rolls', 'Bonus_Points_From_Top4', 'C_Total', 'Badges', 'LastUpdated']);
    sheet.setFrozenRows(1);
  }

  // Players_Prize-Wall-Points
  if (!ss.getSheetByName('Players_Prize-Wall-Points')) {
    const sheet = ss.insertSheet('Players_Prize-Wall-Points');
    sheet.appendRow(['PreferredName', 'Dice_Points_Available', 'Dice_Points_Spent', 'Last_Event', 'LastUpdated']);
    sheet.setFrozenRows(1);
  }
}

// ============================================================================
// DAILY CLOSE CHECKLIST SERVICE
// ============================================================================

/**
 * Saves a Daily Close checklist run into Daily_Close_Log.
 * @param {Object} payload
 *   - location: string
 *   - itemsChecked: string[]   // IDs or labels of checked items
 *   - itemsUnchecked: string[] // IDs or labels of unchecked items
 *   - notes: string
 * @return {Object} result
 */
function saveDailyCloseChecklist(payload) {
  if (!payload) {
    throw new Error('Payload is required');
  }

  const checked = payload.itemsChecked || [];
  const unchecked = payload.itemsUnchecked || [];
  const allDone = unchecked.length === 0 && checked.length > 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Daily_Close_Log');

  // Create sheet with headers if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Daily_Close_Log');
    sheet.appendRow([
      'Timestamp',
      'Run_By',
      'Location',
      'All_Checks_Completed',
      'Items_Checked',
      'Items_Unchecked',
      'Notes'
    ]);
    sheet.setFrozenRows(1);
  }

  const userEmail = getCurrentUserEmail();
  const row = [
    new Date(),                  // Timestamp
    userEmail,                   // Run_By
    payload.location || '',      // Location
    allDone,                     // All_Checks_Completed
    checked.join(', '),          // Items_Checked
    unchecked.join(', '),        // Items_Unchecked
    payload.notes || ''          // Notes
  ];

  sheet.appendRow(row);

  return {
    success: true,
    allChecksCompleted: allDone,
    itemsChecked: checked.length,
    itemsUnchecked: unchecked.length
  };
}

/**
 * Gets current user email for signature display
 * @return {string} User email or 'Unknown'
 */
function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || 'Unknown';
  } catch (e) {
    return 'Unknown';
  }
}

// ============================================================================
// WORKBOOK INITIALIZATION
// ============================================================================

/**
 * Initializes workbook on first run
 */
function initializeWorkbook() {
  // Run auto-fix for all gates
  const gates = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
  gates.forEach(gate => {
    try {
      runAutoFix(gate);
    } catch (e) {
      console.error(`Failed to fix gate ${gate}:`, e);
    }
  });

  // Ensure all sheets
  ensureAllSheets();

  // Log initialization
  logIntegrityAction('WORKBOOK_INIT', {
    details: 'Workbook initialized with all required sheets',
    status: 'SUCCESS'
  });

  SpreadsheetApp.getUi().alert(
    'Initialization Complete',
    'Cosmic Event Manager v7.9.6 is ready!\n\nRun "Build / Repair" from Ops menu to verify health.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}