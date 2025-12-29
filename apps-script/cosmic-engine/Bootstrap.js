/**
 * Bootstrap Script - Run this FIRST before adding HTML files
 * This creates all required sheets and validates the basic setup
 * without needing any UI files.
 */

/**
 * STEP 1: Run this function first (before adding HTML files)
 * This will create all required sheets and basic setup
 */
function bootstrapCosmicEngine() {
  const ui = SpreadsheetApp.getUi();

  try {
    ui.alert(
      'Bootstrap Starting',
      'This will create all required sheets and initialize the Cosmic Event Manager.\n\nThis may take 10-15 seconds...',
      ui.ButtonSet.OK
    );

    // Run all the fixes sequentially
    const results = [];

    // Fix A: Headers/Schema
    try {
      ensureCatalogSchema();
      ensureKeyTrackerSchema();
      ensureBPTotalSchema();
      results.push('✓ Gate A: Schemas created');
    } catch (e) {
      results.push('✗ Gate A: ' + e.message);
    }

    // Fix B: Canonical Names
    try {
      ensureKeyTrackerSchema();
      ensureBPTotalSchema();
      results.push('✓ Gate B: Player tracking ready');
    } catch (e) {
      results.push('✗ Gate B: ' + e.message);
    }

    // Fix C: Required Sheets
    try {
      createAllRequiredSheets();
      results.push('✓ Gate C: All sheets created');
    } catch (e) {
      results.push('✗ Gate C: ' + e.message);
    }

    // Fix D: Throttle Config
    try {
      const throttle = getThrottleKV();
      results.push('✓ Gate D: Throttle configured');
    } catch (e) {
      results.push('✗ Gate D: ' + e.message);
    }

    // Fix E: Inventory
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const catalogSheet = ss.getSheetByName('Prize_Catalog');
      if (catalogSheet) {
        results.push('✓ Gate E: Catalog ready (add items manually)');
      }
    } catch (e) {
      results.push('✗ Gate E: ' + e.message);
    }

    // Log the bootstrap
    logIntegrityAction('BOOTSTRAP', {
      details: 'Initial bootstrap completed',
      status: 'SUCCESS'
    });

    results.push('\n✓ Bootstrap Complete!');
    results.push('\nNext Steps:');
    results.push('1. Add HTML files to Apps Script (see INSTALL.md)');
    results.push('2. Refresh the sheet');
    results.push('3. Use the menu: Cosmic Tournament v7.9');

    ui.alert(
      'Bootstrap Complete!',
      results.join('\n'),
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert(
      'Bootstrap Error',
      'Error: ' + e.message + '\n\nCheck the script logs for details.',
      ui.ButtonSet.OK
    );
    console.error('Bootstrap failed:', e);
  }
}

/**
 * Creates all required sheets
 */
function createAllRequiredSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Integrity_Log
  if (!ss.getSheetByName('Integrity_Log')) {
    const sheet = ss.insertSheet('Integrity_Log');
    sheet.appendRow(['Timestamp', 'Event_ID', 'Action', 'Operator', 'PreferredName', 'Seed',
                     'Checksum_Before', 'Checksum_After', 'RL_Band', 'DF_Tags', 'Details', 'Status']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:L1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }

  // Spent_Pool
  if (!ss.getSheetByName('Spent_Pool')) {
    const sheet = ss.insertSheet('Spent_Pool');
    sheet.appendRow(['Event_ID', 'Item_Code', 'Item_Name', 'Level', 'Qty', 'COGS', 'Total',
                     'Timestamp', 'Batch_ID', 'Reverted', 'Event_Type']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }

  // Attendance_Missions
  if (!ss.getSheetByName('Attendance_Missions')) {
    const sheet = ss.insertSheet('Attendance_Missions');
    sheet.appendRow(['PreferredName', 'Points_From_Dice_Rolls', 'Bonus_Points_From_Top4',
                     'C_Total', 'Badges', 'LastUpdated']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }

  // Players_Prize-Wall-Points
  if (!ss.getSheetByName('Players_Prize-Wall-Points')) {
    const sheet = ss.insertSheet('Players_Prize-Wall-Points');
    sheet.appendRow(['PreferredName', 'Dice_Points_Available', 'Dice_Points_Spent',
                     'Last_Event', 'LastUpdated']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  }

  // Preview_Artifacts (hidden)
  if (!ss.getSheetByName('Preview_Artifacts')) {
    const sheet = ss.insertSheet('Preview_Artifacts');
    sheet.appendRow(['Artifact_ID', 'Event_ID', 'Seed', 'Preview_Hash', 'Created_At', 'Expires_At']);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
}

/**
 * Simple health check without HTML
 * Returns a text summary
 */
function checkHealthSimple() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];

  // Check required sheets
  const requiredSheets = [
    'Prize_Catalog',
    'Prize_Throttle',
    'Integrity_Log',
    'Spent_Pool',
    'Key_Tracker',
    'BP_Total'
  ];

  requiredSheets.forEach(name => {
    const exists = ss.getSheetByName(name) !== null;
    results.push(`${exists ? '✓' : '✗'} ${name}`);
  });

  return results.join('\n');
}

/**
 * Show simple health status
 */
function showSimpleHealth() {
  const ui = SpreadsheetApp.getUi();
  const status = checkHealthSimple();

  ui.alert(
    'System Health Check',
    'Sheet Status:\n\n' + status + '\n\n' +
    'For full health dashboard, add HTML files and use:\nOps → Build / Repair',
    ui.ButtonSet.OK
  );
}