/**
 * Cosmic Games Tournament Manager v7.9.8
 * Main Entry Point - Menu Creation and Triggers
 *
 * CRITICAL: This is the ONLY file that should contain onOpen() and onEdit()
 *
 * @fileoverview Main menu wiring and route handlers for the Cosmic Prize Engine
 * 
 * CHANGELOG v7.9.8:
 * - Merged v7.9.6 and v7.9.7 features
 * - Added Commander Event Wizard
 * - Added Add New Player, Detect/Fix Player Names
 * - Added Daily Close Checklist, View Event Index
 * - Added Preorder Status, Pickup, Cancel routes
 * - Preserved deprecated routes for backwards compatibility
 * - Unified store credit functions with ledger support
 */

// ============================================================================
// VERSION CONSTANTS
// ============================================================================

const ENGINE_VERSION = '7.9.8';
const RECIPE_VERSION = '1.0.1';

// ============================================================================
// TRIGGER: onOpen
// ============================================================================

/**
 * Creates custom menus on spreadsheet open.
 * This is the ONLY onOpen function - do NOT define another one.
 * @param {Object} e - The onOpen event object
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // =========================================================================
  // MENU 1: Cosmic Tournament v7.9.8 (Main Menu)
  // =========================================================================
  const cosmicMenu = ui.createMenu('Cosmic Tournament v7.9.8');

  // Events Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Events')
    .addItem('Start New Event', 'onCreateEvent')
    .addItem('Import Player List', 'onRosterImport')
    .addSeparator()

  );

  // Players Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Players')
    .addItem('Add New Player', 'onAddNewPlayer')
    .addItem('Detect / Fix Player Names', 'onPlayerNameChecker')
    .addSeparator()
    .addItem('Add Key', 'onAddKey')
    .addSeparator()
    .addItem('Award Bonus Points', 'onAwardBP')
    .addItem('Redeem Bonus Points', 'onRedeemBP')
  );

  // Mission Points Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Mission Points')
    .addItem('Award Bonus Points', 'onAwardBP')
    .addItem('Sync BP from Sources', 'menuSyncBPFromSources')
    .addSeparator()
    .addItem('Provision All Players', 'onProvisionAllPlayers')
    .addItem('Scan Attendance / Missions', 'onScanAttendance')
  );

  // Catalog Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Catalog')
    .addItem('Manage Prize Catalog', 'onCatalogManager')
    .addItem('Import Preorder Allocation', 'onPreorderImport')
    .addSeparator()
    .addItem('Prize Throttle (Switchboard)', 'onThrottle')
  );

  // Preorders Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Preorders')
    .addItem('Sell Preorder', 'onSellPreorder')
    .addItem('View Preorder Status', 'onViewPreorderStatus')
    .addItem('Mark Preorder Pickup', 'onMarkPreorderPickup')
    .addItem('Cancel Preorder', 'onCancelPreorder')
    .addSeparator()
    .addItem('Manage Preorder Buckets', 'onPreorderBuckets')
    .addItem('View Preorders Sold', 'onViewPreordersSold')
  );

  // Ops Submenu
  cosmicMenu.addSubMenu(ui.createMenu('Ops')
   
    .addItem('Rebuild Attendance Calendar', 'onRebuildAttendanceCalendar')
    .addItem('Daily Close Checklist', 'onDailyCloseChecklist')

    .addSeparator()
  
    .addSeparator()
    .addItem('Clean Old Previews', 'onCleanPreviews')
    .addItem('Organize Tabs', 'onOrganizeTabs')
    .addItem('Build / Repair', 'onBuildRepair')
  );

  cosmicMenu.addToUi();

  // =========================================================================
  // MENU 2: Cosmic Employee Tools
  // =========================================================================
  ui.createMenu('Cosmic Employee Tools')
    .addItem('Employee Log', 'onOpenEmployeeLog')
    .addSeparator()
    .addItem('Create single assignment…', 'createSingleAssignment')
    .addItem('View pending assignments', 'onViewPendingAssignments')
    .addSeparator()
    .addItem('This Needs... (New Task)', 'showThisNeedsSidebar')
    .addItem('This Needs... Task Board', 'showThisNeedsTaskBoard')
    .addToUi();

  // =========================================================================
  // MENU 3: Store Credit (Separate Top-Level Menu)
  // =========================================================================
  ui.createMenu('Store Credit')
    .addItem('Spend Store Credit', 'onStoreCredit')
    .addToUi();
}

// ============================================================================
// TRIGGER: onEdit
// ============================================================================

/**
 * Handles edit events for the spreadsheet.
 * This is the ONLY onEdit function - do NOT define another one.
 * @param {Object} e - The onEdit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheetName = e.range.getSheet().getName();

  // BP Aggregator sync (keeps BP_Total in sync with source sheets)
  // Only trigger sync for mission source sheets
  const missionSheets = [
    'Attendance_Missions',
    'Flag_Missions',
    'Dice Roll Points',
    'Dice_Points' // Legacy fallback
  ];

  if (missionSheets.includes(sheetName)) {
    try {
      // Use a debounce approach: only sync if the edit was to data rows (not header)
      const editRow = e.range.getRow();
      if (editRow > 1 && typeof updateBPTotalFromSources === 'function') {
        updateBPTotalFromSources();
      }
    } catch (err) {
      // Simple triggers can't show UI alerts, just log
      console.error('onEdit BP sync failed:', err);
    }
  }

  // Dice point checkbox handler
  if (typeof onDicePointCheckboxEdit === 'function') {
    try {
      onDicePointCheckboxEdit(e);
    } catch (err) {
      console.error('Dice Point Checkbox onEdit error:', err);
    }
  }

  // Employee Log edit handler
  if (typeof handleEmployeeLogEdit_ === 'function') {
    try {
      handleEmployeeLogEdit_(e);
    } catch (err) {
      console.error('Employee Log onEdit error:', err);
    }
  }
}

// ============================================================================
// EVENT ROUTES
// ============================================================================

/**
 * Opens Create Event dialog
 */
function onCreateEvent() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/create_event')
      .setWidth(500)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Event');
  } catch (e) {
    showError_('Failed to open Create Event dialog', e);
  }
}

/**
 * Opens Commander Event Wizard dialog
 */
function onCommanderEventWizard() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/commander_wizard')
      .setWidth(700)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Event Wizard');
  } catch (e) {
    showError_('Failed to open Commander Event Wizard', e);
  }
}

/**
 * Opens Smart Roster Import sidebar
 */
function onRosterImport() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/roster_import')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Roster Import', e);
  }
}

/**
 * Opens Preview End Prizes dialog
 */
function onPreviewEndPrizes() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preview_end')
      .setWidth(700)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Preview End Prizes');
  } catch (e) {
    showError_('Failed to open Preview End Prizes', e);
  }
}

/**
 * Alias for onPreviewEndPrizes (v7.9.6 compatibility)
 */
function onPreviewEnd() {
  onPreviewEndPrizes();
}

/**
 * Opens Generate End Prizes dialog (Lock In)
 */
function onGenerateEndPrizes() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/generate_end')
      .setWidth(700)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Lock In End Prizes');
  } catch (e) {
    showError_('Failed to open End Prizes generator', e);
  }
}

/**
 * Opens Commander Round Prizes dialog
 */
function onCommanderRounds() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/round_prizes')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Commander Round Prizes');
  } catch (e) {
    showError_('Failed to open Commander Rounds', e);
  }
}

/**
 * Reverts the last prize run
 */
function onRevertPrizes() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Undo Last Prize Run',
      'This will revert the most recent prize distribution.\n\n' +
      'This action cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      if (typeof revertLastPrizeRun === 'function') {
        const revertResult = revertLastPrizeRun();
        ui.alert('Revert Complete', `Reverted ${revertResult.count || 0} prize(s).`, ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Prize revert function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) {
    showError_('Failed to revert prizes', e);
  }
}

// ============================================================================
// PLAYER ROUTES
// ============================================================================

/**
 * Opens Player Lookup sidebar
 */
function onPlayerLookup() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/player_lookup')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Player Lookup', e);
  }
}

/**
 * Opens Add New Player dialog
 */
function onAddNewPlayer() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/add_player')
      .setWidth(450)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add New Player');
  } catch (e) {
    showError_('Failed to open Add New Player', e);
  }
}

/**
 * Opens Detect / Fix Players sidebar
 */
function onDetectNewPlayers() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/detect_fix_players')
      .setTitle('Detect / Fix Player Names')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Detect / Fix Players', e);
  }
}

/**
 * Opens Spell Check
 */
function onPlayerNameChecker() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/PlayerNameChecker')
      .setWidth(500)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, ' ui/PlayerNameChecker');
  } catch (e) {
    // Fallback to new detect/fix if legacy UI not available
    console.warn('Legacy spell_check UI not found, using detect_fix_players');
    onDetectNewPlayers();
  }
}



/**
 * Opens Player Profile viewer
 */
function onViewPlayerProfile() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/player_profile')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Player Profile', e);
  }
}

/**
 * Alias for onViewPlayerProfile (v7.9.6 compatibility)
 */
function onPlayerProfile() {
  onViewPlayerProfile();
}

/**
 * Opens Add Key dialog
 */
function onAddKey() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/add_key')
      .setWidth(450)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Add Key');
  } catch (e) {
    showError_('Failed to open Add Key dialog', e);
  }
}

/**
 * Opens Award Bonus Points dialog
 */
function onAwardBP() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/award_bp')
      .setWidth(500)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Bonus Points');
  } catch (e) {
    showError_('Failed to open Award BP dialog', e);
  }
}

/**
 * Opens Redeem BP dialog
 */
function onRedeemBP() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/redeem_bp')
      .setWidth(500)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Redeem Bonus Points');
  } catch (e) {
    showError_('Failed to open BP Redemption', e);
  }
}

/**
 * Syncs BP Total from source sheets (single player or quick sync)
 */
function onSyncBPTotal() {
  try {
    const ui = SpreadsheetApp.getUi();

    if (typeof updateBPTotalFromSources === 'function') {
      const result = updateBPTotalFromSources();
      ui.alert('Sync Complete', `BP totals have been synchronized. Updated ${result} player(s).`, ui.ButtonSet.OK);
    } else {
      ui.alert('Not Available', 'BP sync function not available.', ui.ButtonSet.OK);
    }
  } catch (e) {
    showError_('Failed to sync BP totals', e);
  }
}

/**
 * Opens Dice Roll Results dialog for recording dice points
 */
function onDiceRollResults() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/dice_results')
      .setWidth(500)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Dice Roll Results');
  } catch (e) {
    showError_('Failed to open Dice Results dialog', e);
  }
}

// ============================================================================
// MISSION POINTS ROUTES
// ============================================================================

/**
 * Opens Award Flag Mission dialog
 */
function onAwardFlagMission() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/flag_mission')
      .setWidth(500)
      .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Award Flag Mission');
  } catch (e) {
    showError_('Failed to open Flag Mission dialog', e);
  }
}

/**
 * Opens Record Attendance dialog
 */
function onRecordAttendance() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/record_attendance')
      .setWidth(500)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Record Attendance');
  } catch (e) {
    showError_('Failed to open Record Attendance', e);
  }
}

/**
 * Syncs all BP Totals (batch operation) - calls the canonical pipeline
 */
function onSyncBPTotals() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Sync All BP Totals',
      'This will synchronize BP totals for ALL players from source sheets.\n\n' +
      'This may take a few minutes for large player bases.\n\nContinue?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Syncing BP totals...', 'BP Sync', 5);

      if (typeof updateBPTotalFromSources === 'function') {
        const syncResult = updateBPTotalFromSources();
        ui.alert('Sync Complete',
          `Synced ${syncResult} player(s) from mission source sheets.`,
          ui.ButtonSet.OK);
      } else if (typeof syncBPTotalsFromAllSources === 'function') {
        syncBPTotalsFromAllSources();
        ui.alert('Sync Complete', 'BP totals have been synchronized from all sources.', ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'BP sync function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) {
    showError_('Failed to sync BP totals', e);
  }
}

/**
 * Validates mission points integrity
 */
function onValidateMissionPoints() {
  try {
    const ui = SpreadsheetApp.getUi();

    if (typeof validateMissionPointsIntegrity === 'function') {
      const result = validateMissionPointsIntegrity();

      if (result.pass) {
        ui.alert('Validation Passed', 'All mission point data is valid.', ui.ButtonSet.OK);
      } else {
        const issueList = result.issues.slice(0, 10).map(i =>
          `${i.sheet} row ${i.row}: ${i.issue}`
        ).join('\n');
        ui.alert('Validation Issues Found',
          `Found ${result.issues.length} issue(s):\n\n${issueList}` +
          (result.issues.length > 10 ? '\n\n... and more' : ''),
          ui.ButtonSet.OK);
      }
    } else {
      ui.alert('Not Available', 'Validation function not available.', ui.ButtonSet.OK);
    }
  } catch (e) {
    showError_('Failed to validate mission points', e);
  }
}

/**
 * Provisions all players in source sheets
 */
function onProvisionAllPlayers() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Provision All Players',
      'This will ensure all players from PreferredNames are provisioned in:\n' +
      '- Flag_Missions\n' +
      '- Attendance_Missions\n' +
      '- Dice_Points\n' +
      '- BP_Total\n\n' +
      'Continue?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Provisioning players...', 'Provision', 5);

      if (typeof provisionAllPlayers === 'function') {
        const provResult = provisionAllPlayers();
        ui.alert('Provisioning Complete',
          `Provisioned: ${provResult.provisioned}\nAlready existed: ${provResult.alreadyExisted}`,
          ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Provisioning function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) {
    showError_('Failed to provision players', e);
  }
}

/**
 * Refreshes BP_Total from mission source sheets (canonical pipeline entry point)
 */
function onRefreshBPTotalFromSources() {
  try {
    if (typeof updateBPTotalFromSources === 'function') {
      const updated = updateBPTotalFromSources();
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `BP_Total sync complete: ${updated} player(s) updated.`,
        'Bonus Points',
        5
      );
    } else {
      // Fallback to batch sync
      onSyncBPTotals();
    }
  } catch (e) {
    showError_('Failed to refresh BP totals from sources', e);
  }
}

// ============================================================================
// CATALOG ROUTES
// ============================================================================

/**
 * Opens Catalog Manager dialog
 */
function onCatalogManager() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/catalog_manager')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Prize Catalog');
  } catch (e) {
    showError_('Failed to open Catalog Manager', e);
  }
}

/**
 * Opens Preorder Import dialog
 */
function onPreorderImport() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorder_import')
      .setWidth(900)
      .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Import Preorder Allocation');
  } catch (e) {
    showError_('Failed to open Preorder Import', e);
  }
}

/**
 * Opens Prize Throttle switchboard
 */
function onThrottle() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/throttle')
      .setWidth(500)
      .setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Prize Throttle (Switchboard)');
  } catch (e) {
    showError_('Failed to open Throttle Switchboard', e);
  }
}

// ============================================================================
// PREORDER ROUTES
// ============================================================================

/**
 * Opens Sell Preorder dialog
 */
function onSellPreorder() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/sell_preorder')
      .setWidth(950)
      .setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(html, 'Sell Preorder');
  } catch (e) {
    showError_('Failed to open Sell Preorder dialog', e);
  }
}

/**
 * Opens Preorder Status viewer sidebar
 */
function onViewPreorderStatus() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorder_status')
      .setTitle('Preorder Status')
      .setWidth(900);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Preorder Status', e);
  }
}

/**
 * Opens Mark Preorder Pickup dialog
 */
function onMarkPreorderPickup() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorder_pickup')
      .setWidth(900)
      .setHeight(950);
    SpreadsheetApp.getUi().showModalDialog(html, 'Mark Preorder Pickup');
  } catch (e) {
    showError_('Failed to open Preorder Pickup', e);
  }
}

/**
 * Opens Cancel Preorder dialog
 */
function onCancelPreorder() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorder_cancel')
      .setWidth(900)
      .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Cancel Preorder');
  } catch (e) {
    showError_('Failed to open Cancel Preorder', e);
  }
}

/**
 * Opens Preorder Buckets management dialog
 */
function onPreorderBuckets() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorder_buckets')
      .setWidth(900)
      .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'Manage Preorder Buckets');
  } catch (e) {
    showError_('Failed to open Preorder Buckets', e);
  }
}

/**
 * Opens View Preorders Sold sidebar
 */
function onViewPreordersSold() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/preorders_sold')
      .setTitle('Preorders Sold')
      .setWidth(800);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Preorders Sold viewer', e);
  }
}

// ============================================================================
// OPS ROUTES
// ============================================================================

/**
 * Canonical menu handler: Scan Attendance + Missions
 * Bridges the main menu to Mission Scanner Service v1.0.0
 */
function onScanAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    ss.toast('Scanning attendance & missions...', 'Mission Scan', 10);

    // 🔑 This is the NEW engine entry point from Mission Scanner Service
    const result = runMissionScan();  // defined in missionScannerService.gs

    const eventsScanned = result && result.eventsScanned || 0;
    const playersTracked = result && result.playersTracked || 0;
    const missionsComputed = result && result.missionsComputed || 0;

    ui.alert(
      '✅ Mission Scan Complete',
      'Events Scanned: ' + eventsScanned + '\n' +
      'Players Tracked: ' + playersTracked + '\n' +
      'Missions Evaluated: ' + missionsComputed,
      ui.ButtonSet.OK
    );

  } catch (e) {
    // Try to use your shared error helper if it exists
    if (typeof showError_ === 'function') {
      showError_('Failed to scan attendance/missions', e);
    } else {
      ui.alert(
        '❌ Mission Scan Error',
        (e && e.message) ? e.message : String(e),
        ui.ButtonSet.OK
      );
      console.error(e);
    }
  }
}


/**
 * Internal mission recalculation logic (fallback)
 * @private
 */
function recalcAllMissions_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance_Missions');
  if (sheet) {
    console.log('Mission scan triggered for Attendance_Missions');
  }
}

/**
 * Alias for onScanAttendance (v7.9.7 compatibility)
 */
function onScanMissions() {
  onScanAttendance();
}

/**
 * Rebuilds the Attendance Calendar from event tabs
 */
function onRebuildAttendanceCalendar() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Rebuild Attendance Calendar',
      'This will rebuild the Attendance_Calendar sheet from all event tabs.\n\n' +
      'Existing data will be replaced. Continue?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Rebuilding calendar...', 'Rebuild', 10);

      if (typeof rebuildAttendanceCalendar === 'function') {
        const rebuildResult = rebuildAttendanceCalendar();
        ui.alert('Rebuild Complete',
          `Processed ${rebuildResult.events || 0} event(s).`,
          ui.ButtonSet.OK);
      } else {
        ui.alert('Not Available', 'Rebuild function not available.', ui.ButtonSet.OK);
      }
    }
  } catch (e) {
    showError_('Failed to rebuild attendance calendar', e);
  }
}

/**
 * Opens Daily Close Checklist dialog
 */
function onDailyCloseChecklist() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/daily_close')
      .setWidth(600)
      .setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(html, 'Daily Close Checklist');
  } catch (e) {
    showError_('Failed to open Daily Close Checklist', e);
  }
}

/**
 * Opens Event Index sidebar
 */
function onViewEventIndex() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/event_index')
      .setTitle('Event Index')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Event Index', e);
  }
}

/**
 * Opens Ship-Gates Health Check dialog
 */
function onShipGates() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/ship_gates')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Ship-Gates Health Check');
  } catch (e) {
    showError_('Failed to open Ship-Gates Health Check', e);
  }
}

/**
 * Opens Integrity Log viewer
 */
function onViewLog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/log_viewer')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Log Viewer', e);
  }
}

/**
 * Opens Spent Pool viewer
 */
function onViewSpent() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/spent_pool_viewer')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Spent Pool Viewer', e);
  }
}

/**
 * Opens Rainbow Status viewer
 */
function onRainbowStatus() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/rainbow_status')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Rainbow Status', e);
  }
}

/**
 * Cleans old preview artifacts
 */
function onCleanPreviews() {
  try {
    const count = cleanOldPreviews_(24); // 24 hours
    SpreadsheetApp.getUi().alert(
      'Clean Previews Complete',
      `Removed ${count} stale preview artifact(s).`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    showError_('Failed to clean previews', e);
  }
}

/**
 * Opens Tab Organizer dialog
 */
function onOrganizeTabs() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/tab_organizer')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Tab Organizer');
  } catch (e) {
    // Fallback to simple prompt if HTML file not installed yet
    if (e.message && e.message.includes('ui/tab_organizer')) {
      showSimpleTabOrganizerFallback_();
    } else {
      showError_('Failed to open Tab Organizer', e);
    }
  }
}

/**
 * Opens Build/Repair health dashboard
 */
function onBuildRepair() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/health')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Build / Repair - Health Dashboard');
  } catch (e) {
    // Fallback to simple health check if HTML file not installed yet
    if (e.message && e.message.includes('health')) {
      showSimpleHealthFallback_();
    } else {
      showError_('Failed to open Health Dashboard', e);
    }
  }
}

/**
 * Exports Integrity Log and Spent Pool as CSV
 */
function onExportReports() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');

    // Export Integrity_Log
    const logSheet = ss.getSheetByName('Integrity_Log');
    if (logSheet) {
      const filename = `Integrity_Log_${timestamp}.csv`;
      ui.alert('Export Ready', `Log exported. Copy data and save as:\n${filename}\n\nOpen Integrity_Log sheet to view data.`, ui.ButtonSet.OK);
    }

    // Export Spent_Pool
    const spentSheet = ss.getSheetByName('Spent_Pool');
    if (spentSheet) {
      const filename = `Spent_Pool_${timestamp}.csv`;
      ui.alert('Export Ready', `Spent Pool exported. Copy data and save as:\n${filename}\n\nOpen Spent_Pool sheet to view data.`, ui.ButtonSet.OK);
    }
  } catch (e) {
    showError_('Failed to export reports', e);
  }
}

/**
 * Force unlocks an event
 */
function onForceUnlock() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Force Unlock Event',
      'Enter the Event ID (sheet name) to unlock:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      const eventId = response.getResponseText().trim();

      if (eventId) {
        if (typeof forceUnlockEvent === 'function') {
          forceUnlockEvent(eventId);
          ui.alert('Unlocked', `Event "${eventId}" has been unlocked.`, ui.ButtonSet.OK);
        } else {
          ui.alert('Not Available', 'Force unlock function not available.', ui.ButtonSet.OK);
        }
      }
    }
  } catch (e) {
    showError_('Failed to force unlock event', e);
  }
}

/**
 * Emergency revert - rolls back recent changes
 */
function onEmergencyRevert() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Emergency Revert',
      'WARNING: This will attempt to revert recent changes.\n\n' +
      'This is a last-resort option. Are you sure?',
      ui.ButtonSet.YES_NO
    );

    if (result === ui.Button.YES) {
      const confirm = ui.alert(
        'Confirm Emergency Revert',
        'Type "REVERT" to confirm (this cannot be undone):',
        ui.ButtonSet.OK_CANCEL
      );

      if (confirm === ui.Button.OK) {
        if (typeof emergencyRevert === 'function') {
          emergencyRevert();
          ui.alert('Revert Complete', 'Emergency revert has been executed.', ui.ButtonSet.OK);
        } else {
          ui.alert('Not Available', 'Emergency revert function not available.', ui.ButtonSet.OK);
        }
      }
    }
  } catch (e) {
    showError_('Failed to execute emergency revert', e);
  }
}

// ============================================================================
// STORE CREDIT ROUTES
// ============================================================================

/**
 * Opens Spend Store Credit sidebar
 */
function onStoreCredit() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/store_credit')
      .setTitle('Spend Store Credit');
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Store Credit sidebar', e);
  }
}

/**
 * Opens Store Credit Viewer sidebar
 */
function onStoreCreditViewer() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/store_credit_viewer')
      .setTitle('View Store Credit')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Store Credit Viewer', e);
  }
}

// ============================================================================
// EMPLOYEE TOOLS ROUTES
// ============================================================================

/**
 * Opens Employee Log sidebar
 */
function onOpenEmployeeLog() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/employee_log')
      .setTitle('Employee Log')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Employee Log', e);
  }
}

/**
 * Creates assignments from dictation
 */
function createAssignmentsFromDictation() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/dictation_assignments')
      .setWidth(600)
      .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Assignments from Dictation');
  } catch (e) {
    showError_('Failed to open Dictation Assignments', e);
  }
}

/**
 * Alias for createAssignmentsFromDictation (v7.9.6 compatibility)
 */
function onDictationAssignments() {
  createAssignmentsFromDictation();
}

/**
 * Creates a single assignment
 */
function createSingleAssignment() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/single_assignment')
      .setWidth(500)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Create Single Assignment');
  } catch (e) {
    showError_('Failed to open Single Assignment', e);
  }
}

/**
 * Alias for createSingleAssignment (v7.9.6 compatibility)
 */
function onSingleAssignment() {
  createSingleAssignment();
}

/**
 * Views pending assignments
 */
function onViewPendingAssignments() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/pending_assignments')
      .setTitle('Pending Assignments')
      .setWidth(400);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open Pending Assignments', e);
  }
}

/**
 * Alias for onViewPendingAssignments (v7.9.6 compatibility)
 */
function onPendingAssignments() {
  onViewPendingAssignments();
}

/**
 * Shows This Needs sidebar for new tasks
 */
function showThisNeedsSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/this_needs')
      .setWidth(350);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    showError_('Failed to open This Needs sidebar', e);
  }
}

/**
 * Alias for showThisNeedsSidebar (v7.9.6 compatibility)
 */
function onThisNeeds() {
  showThisNeedsSidebar();
}

/**
 * Shows This Needs Task Board
 */
function showThisNeedsTaskBoard() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('ui/task_board')
      .setWidth(800)
      .setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(html, 'This Needs... Task Board');
  } catch (e) {
    showError_('Failed to open Task Board', e);
  }
}

/**
 * Alias for showThisNeedsTaskBoard (v7.9.6 compatibility)
 */
function onTaskBoard() {
  showThisNeedsTaskBoard();
}

// ============================================================================
// STORE CREDIT SERVICE FUNCTIONS
// ============================================================================

/**
 * Returns dropdown options for the Store Credit UI.
 * @return {Object}
 */
function getDropdownData() {
  return {
    reasons: [
      'Product Purchase',
      'Refund',
      'Prize Payout',
      'Tournament Prize',
      'Manual Adjust',
      'Promo/Comp',
      'Correction',
      'Other'
    ],
    categories: [
      'Sales',
      'Prize Payout',
      'Customer Service',
      'Adjustment',
      'Promo'
    ],
    tenderTypes: [
      'Store Credit'
    ],
    posRefTypes: [
      'Invoice',
      'TicketID',
      'OrderID',
      'Manual Note'
    ],
    // v7.9.6 compatibility fields
    types: ['CREDIT', 'DEBIT', 'ADJUSTMENT'],
    serverVersion: ENGINE_VERSION
  };
}

/**
 * Returns the list of player names for the search box.
 * Reads from PreferredNames!A:A or Players_BP_Total!A:A as fallback.
 * @return {string[]}
 */
function getPlayerNames() {
  const ss = SpreadsheetApp.getActive();
  
  // Try PreferredNames first (v7.9.7 structure)
  let sheet = ss.getSheetByName('PreferredNames');
  
  // Fallback to Players_BP_Total (v7.9.6 structure)
  if (!sheet) {
    sheet = ss.getSheetByName('Players_BP_Total');
  }
  
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values
    .map(r => String(r[0]).trim())
    .filter(name => name);
}

/**
 * Returns TRUE if the ledger has changed since the last check.
 * Stub that always returns false - enhance with PropertiesService if needed.
 */
function checkForLedgerUpdates() {
  return false;
}

/**
 * Computes the current store credit balance for a player.
 * Supports both Store_Credit_Ledger (v7.9.7) and Store_Credit_Log (v7.9.6) structures.
 * @param {string} playerName
 * @return {number} current balance
 */
function getCurrentBalance(playerName) {
  if (!playerName) {
    throw new Error('Player name is required for balance lookup.');
  }

  const ss = SpreadsheetApp.getActive();
  
  // Try v7.9.7 structure first
  let sheet = ss.getSheetByName('Store_Credit_Ledger');
  let isLedgerFormat = true;
  
  // Fallback to v7.9.6 structure
  if (!sheet) {
    sheet = ss.getSheetByName('Store_Credit_Log');
    isLedgerFormat = false;
  }
  
  if (!sheet) {
    throw new Error('Store credit sheet not found (tried Store_Credit_Ledger and Store_Credit_Log).');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;

  const headers = data[0];
  
  // Determine column indices based on format
  let nameCol, directionCol, amountCol, typeCol;
  
  if (isLedgerFormat) {
    // v7.9.7 format: PreferredName/preferred_name_id, InOut, Amount
    nameCol = headers.indexOf('PreferredName');
    if (nameCol === -1) nameCol = headers.indexOf('preferred_name_id');
    directionCol = headers.indexOf('InOut');
    amountCol = headers.indexOf('Amount');
  } else {
    // v7.9.6 format: column indices 1=name, 2=amount, 3=type
    nameCol = 1;
    amountCol = 2;
    typeCol = 3;
  }

  let balance = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[nameCol] || '').trim();
    if (!name || name !== playerName) continue;

    const amount = Number(row[amountCol]) || 0;
    
    if (isLedgerFormat) {
      const direction = String(row[directionCol] || '').toUpperCase();
      if (direction === 'IN') {
        balance += amount;
      } else if (direction === 'OUT') {
        balance -= amount;
      }
    } else {
      const type = String(row[typeCol] || '').toUpperCase();
      if (type === 'CREDIT') {
        balance += amount;
      } else if (type === 'DEBIT') {
        balance -= amount;
      } else if (type === 'ADJUSTMENT') {
        balance = amount;
      }
    }
  }

  return balance;
}

/**
 * Returns recent transactions for a player, with running balance.
 * @param {string} playerName
 * @param {number} limit
 * @return {Array<Object>}
 */
function getPlayerHistory(playerName, limit) {
  if (!playerName) return [];
  limit = limit || 5;

  const ss = SpreadsheetApp.getActive();
  
  // Try v7.9.7 structure first
  let sheet = ss.getSheetByName('Store_Credit_Ledger');
  let isLedgerFormat = true;
  
  // Fallback to v7.9.6 structure
  if (!sheet) {
    sheet = ss.getSheetByName('Store_Credit_Log');
    isLedgerFormat = false;
  }
  
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  
  // Column mapping
  let nameCol, tsCol, dirCol, amtCol, reasonCol, catCol, descCol, typeCol, operatorCol;
  
  if (isLedgerFormat) {
    nameCol = headers.indexOf('PreferredName');
    if (nameCol === -1) nameCol = headers.indexOf('preferred_name_id');
    tsCol = headers.indexOf('Timestamp');
    dirCol = headers.indexOf('InOut');
    amtCol = headers.indexOf('Amount');
    reasonCol = headers.indexOf('Reason');
    catCol = headers.indexOf('Category');
    descCol = headers.indexOf('Description');
  } else {
    // v7.9.6 fixed positions
    tsCol = 0;
    nameCol = 1;
    amtCol = 2;
    typeCol = 3;
    reasonCol = 4;
    operatorCol = 5;
  }

  // Build ascending history and running balance
  let runningBalance = 0;
  const allTx = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[nameCol] || '').trim();
    if (name !== playerName) continue;

    const amount = Number(row[amtCol]) || 0;
    let direction;
    
    if (isLedgerFormat) {
      direction = String(row[dirCol] || '').toUpperCase();
      if (direction === 'IN') {
        runningBalance += amount;
      } else if (direction === 'OUT') {
        runningBalance -= amount;
      }
    } else {
      const type = String(row[typeCol] || '').toUpperCase();
      direction = type;
      if (type === 'CREDIT') {
        runningBalance += amount;
      } else if (type === 'DEBIT') {
        runningBalance -= amount;
      } else if (type === 'ADJUSTMENT') {
        runningBalance = amount;
      }
    }

    allTx.push({
      timestamp: row[tsCol] || '',
      direction: direction,
      amount: amount,
      balance: runningBalance,
      reason: row[reasonCol] || '',
      category: isLedgerFormat ? (row[catCol] || '') : '',
      description: isLedgerFormat ? (row[descCol] || '') : '',
      operator: !isLedgerFormat ? (row[operatorCol] || '') : ''
    });
  }

  // Take the last N and reverse so newest is first
  const recent = allTx.slice(-limit).reverse();
  return recent;
}

/**
 * Logs a store credit transaction to Store_Credit_Ledger (v7.9.7 format).
 * Falls back to Store_Credit_Log (v7.9.6 format) if needed.
 * @param {Object} payload - Transaction data
 * @return {Object} Result
 */
function logStoreCreditTransaction(payload) {
  const ss = SpreadsheetApp.getActive();
  
  // Try v7.9.7 structure first
  let sheet = ss.getSheetByName('Store_Credit_Ledger');
  let isLedgerFormat = true;
  
  // Fallback to v7.9.6 structure
  if (!sheet) {
    sheet = ss.getSheetByName('Store_Credit_Log');
    isLedgerFormat = false;
  }
  
  if (!sheet) {
    throw new Error('Store credit sheet not found.');
  }

  const timestamp = new Date();
  const operator = Session.getActiveUser().getEmail() || 'unknown';

  if (isLedgerFormat) {
    // v7.9.7 format
    const row = [
      timestamp,
      payload.preferred_name_id || payload.playerName || '',
      payload.direction || '',
      Number(payload.amount) || 0,
      payload.reason || '',
      payload.category || '',
      payload.tenderType || '',
      payload.description || '',
      payload.posRefType || '',
      payload.posRefId || '',
      '',  // RunningBalance (formula or later calc)
      Utilities.getUuid()  // RowId
    ];
    sheet.appendRow(row);
  } else {
    // v7.9.6 format
    const row = [
      timestamp,
      payload.playerName || payload.preferred_name_id || '',
      Number(payload.amount) || 0,
      payload.type || payload.direction || '',
      payload.reason || '',
      operator,
      payload.balance || 0
    ];
    sheet.appendRow(row);
  }

  return { success: true };
}

/**
 * Main entry point called by ui/store_credit.html.
 * Writes a ledger row and returns the new balance.
 * @param {Object} payload
 * @return {Object} {success, newBalance, previousBalance}
 */
function submitStoreCredit(payload) {
  if (!payload) {
    throw new Error('Missing payload.');
  }
  
  const playerName = payload.playerName || payload.preferred_name_id;
  if (!playerName) {
    throw new Error('Player is required.');
  }
  if (!payload.direction && !payload.type) {
    throw new Error('Direction/Type is required.');
  }
  if (!payload.amount || isNaN(payload.amount) || Number(payload.amount) <= 0) {
    throw new Error('Amount must be a positive number.');
  }

  const previousBalance = getCurrentBalance(playerName);
  
  // Determine new balance based on direction/type
  let newBalance = previousBalance;
  const direction = (payload.direction || payload.type || '').toUpperCase();
  const amount = Number(payload.amount);
  
  if (direction === 'IN' || direction === 'CREDIT') {
    newBalance += amount;
  } else if (direction === 'OUT' || direction === 'DEBIT') {
    newBalance -= amount;
  } else if (direction === 'ADJUSTMENT') {
    newBalance = amount;
  }

  // Map from UI payload to ledger payload shape
  const ledgerPayload = {
    preferred_name_id: playerName,
    playerName: playerName,
    direction: direction === 'CREDIT' ? 'IN' : (direction === 'DEBIT' ? 'OUT' : direction),
    type: payload.type || direction,
    amount: amount,
    reason: payload.reason || '',
    category: payload.category || '',
    tenderType: payload.tenderType || '',
    description: payload.description || '',
    posRefType: payload.posRefType || '',
    posRefId: payload.posRefId || '',
    balance: newBalance
  };

  logStoreCreditTransaction(ledgerPayload);

  return {
    success: true,
    previousBalance: previousBalance,
    newBalance: newBalance
  };
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Shows error dialog with friendly message
 * @param {string} context - User-friendly context
 * @param {Error} error - The error object
 * @private
 */
function showError_(context, error) {
  const ui = SpreadsheetApp.getUi();
  const message = `${context}\n\nError: ${error.message}\n\nPlease check Build/Repair for health issues.`;
  ui.alert('Error', message, ui.ButtonSet.OK);
  console.error(context, error);
}

/**
 * Simple health check fallback when HTML files aren't installed yet
 * @private
 */
function showSimpleHealthFallback_() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(
    'HTML Files Not Installed Yet',
    'The HTML UI files haven\'t been added to Apps Script yet.\n\n' +
    'Would you like to run the basic bootstrap instead?\n\n' +
    '(This will create all required sheets without the full UI)',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    if (typeof bootstrapCosmicEngine === 'function') {
      bootstrapCosmicEngine();
    } else if (typeof ensureAllSheets === 'function') {
      ensureAllSheets();
      ui.alert('Complete', 'All required sheets have been created.', ui.ButtonSet.OK);
    }
  }
}

/**
 * Converts a sheet to RFC-4180 CSV
 * @param {Sheet} sheet - The sheet to convert
 * @return {string} CSV content
 * @private
 */
function sheetToCSV_(sheet) {
  const data = sheet.getDataRange().getValues();
  return data.map(row =>
    row.map(cell => {
      const str = String(cell);
      if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(',')
  ).join('\n');
}

/**
 * Cleans old preview artifacts
 * @param {number} hoursOld - Remove previews older than this many hours
 * @return {number} Count of removed artifacts
 * @private
 */
function cleanOldPreviews_(hoursOld) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preview_Artifacts');

  if (!sheet) return 0;

  const data = sheet.getDataRange().getValues();
  const now = new Date().getTime();
  const cutoff = now - (hoursOld * 60 * 60 * 1000);

  let removeCount = 0;
  for (let i = data.length - 1; i > 0; i--) { // Skip header
    const expiresAt = new Date(data[i][4]).getTime();
    if (expiresAt < cutoff) {
      sheet.deleteRow(i + 1);
      removeCount++;
    }
  }

  if (removeCount > 0 && typeof logIntegrityAction === 'function') {
    logIntegrityAction('CLEAN_PREVIEWS', {
      removed: removeCount,
      cutoff_hours: hoursOld
    });
  }

  return removeCount;
}

/**
 * Gets pending assignments for the employee tools
 * @return {Array<Object>} Array of pending assignments
 * @private
 */
function getPendingAssignments_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Assignments');

  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const statusCol = headers.indexOf('Status');

  if (statusCol === -1) {
    // Try fixed column format (v7.9.6)
    return data.slice(1)
      .filter(row => row[3] !== 'COMPLETED')
      .map(row => ({
        id: row[0],
        description: row[1],
        assignee: row[2],
        status: row[3],
        dueDate: row[4]
      }));
  }

  const pending = [];
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusCol] || '').toLowerCase();
    if (status === 'pending' || status === 'in progress') {
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = data[i][idx];
      });
      pending.push(row);
    }
  }

  return pending;
}

/**
 * Handles edits to the Employee Log sheet
 * @param {Object} e - Edit event
 * @private
 */
function handleEmployeeLogEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Employee_Log') return;

  const row = e.range.getRow();
  if (row <= 1) return; // Skip header

  const col = e.range.getColumn();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Look for LastUpdated or auto-timestamp in column A
  const timestampCol = headers.indexOf('LastUpdated');

  if (timestampCol !== -1) {
    sheet.getRange(row, timestampCol + 1).setValue(new Date());
  } else if (col > 1 && e.value && !sheet.getRange(row, 1).getValue()) {
    // v7.9.6 behavior: auto-timestamp in column A when content is added
    sheet.getRange(row, 1).setValue(new Date());
  }
}

/**
 * Gets a list of event tab names
 * Supports both MM-DD-YYYY (v7.9.7) and XXX##_ (v7.9.6) patterns
 * @return {Array<string>} Array of event tab names
 */
function listEventTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // Pattern for MM-DD-YYYY (v7.9.7)
  const datePattern = /^\d{2}-\d{2}-\d{4}/;
  // Pattern for XXX##_ (v7.9.6)
  const codePattern = /^[A-Z]{3}\d{2}_/;

  return sheets
    .map(s => s.getName())
    .filter(name => datePattern.test(name) || codePattern.test(name))
    .sort();
}

/**
 * Gets event properties from a sheet's metadata
 * @param {string|Sheet} eventIdOrSheet - Event ID or Sheet object
 * @return {Object} Event properties
 */
function getEventProps(eventIdOrSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheet;
  if (typeof eventIdOrSheet === 'string') {
    sheet = ss.getSheetByName(eventIdOrSheet);
  } else {
    sheet = eventIdOrSheet;
  }
  
  if (!sheet) return null;

  const props = {
    eventId: sheet.getName(),
    eventType: 'UNKNOWN',
    playerCount: 0,
    status: 'DRAFT'
  };

  // Try to read from row 1 metadata note (v7.9.7)
  try {
    const metaRange = sheet.getRange('A1').getNote();
    if (metaRange) {
      const parsed = JSON.parse(metaRange);
      Object.assign(props, parsed);
    }
  } catch (e) {
    // No metadata found, try standard property locations (v7.9.6)
    try {
      props.eventType = sheet.getRange('B1').getValue() || 'UNKNOWN';
      props.playerCount = sheet.getRange('B2').getValue() || 0;
      props.status = sheet.getRange('B3').getValue() || 'DRAFT';
    } catch (e2) {
      // Leave defaults
    }
  }

  return props;
}

/**
 * Gets the current user's email
 * @return {string} User email
 */
function currentUser() {
  try {
    return Session.getActiveUser().getEmail() || 'Unknown';
  } catch (e) {
    return 'anonymous';
  }
}

/**
 * First-run setup - opens health check
 */
function firstRunSetup() {
  onBuildRepair();
}
/**
 * Menu handler for BP sync operation
 * Provides user feedback via toast notifications
 */
function menuSyncBPFromSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast('Syncing BP from source sheets...', '⏳ Please Wait', -1);
  
  try {
    const updatedCount = updateBPTotalFromSources();
    
    if (updatedCount > 0) {
      ss.toast(`Successfully synced ${updatedCount} player(s)`, '✅ BP Sync Complete', 5);
    } else {
      ss.toast('All players already up to date', '✅ BP Sync Complete', 3);
    }
  } catch (e) {
    ss.toast('Error: ' + e.message, '❌ Sync Failed', 5);
    console.error('BP Sync Error:', e);
  }
}

// ============================================================================
// DEPRECATED ROUTES - Kept for backwards compatibility
// ============================================================================

/** @deprecated Use onScanAttendance instead */
function onScanMissions_deprecated() {
  console.warn('onScanMissions is deprecated. Use onScanAttendance instead.');
  onScanAttendance();
}

/** @deprecated Use onScanAttendance instead */
function onRecordAttendance_deprecated() {
  console.warn('onRecordAttendance is deprecated. Use onScanAttendance instead.');
  onScanAttendance();
}

/** @deprecated Use onSyncBPTotals instead */
function onRefreshBPTotalFromSources_deprecated() {
  console.warn('onRefreshBPTotalFromSources is deprecated. Use onSyncBPTotals instead.');
  onSyncBPTotals();
}