/**
 * ===================================================================
 * MEGA BOOTSTRAP - Cosmic Event Manager v7.9.6
 * ===================================================================
 *
 * This is a SELF-CONTAINED bootstrap script with ZERO dependencies.
 * Run this FIRST, before adding any other files.
 *
 * INSTRUCTIONS:
 * 1. Copy this entire file to Apps Script as "MegaBootstrap.gs"
 * 2. Run: Functions → megaBootstrapInstall → Run
 * 3. Authorize when prompted
 * 4. Wait ~15 seconds for completion
 * 5. All sheets will be created!
 *
 * After this completes, you can add the other service files and HTML files.
 * ===================================================================
 */

/**
 * RUN THIS FUNCTION FIRST!
 * Creates all required sheets and basic setup
 */
function megaBootstrapInstall() {
  const ui = SpreadsheetApp.getUi();

  try {
    ui.alert(
      '🚀 Mega Bootstrap Starting',
      'This will create all required sheets for the Cosmic Event Manager.\n\n' +
      'This takes about 15 seconds.\n\n' +
      'Click OK to continue...',
      ui.ButtonSet.OK
    );

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    // ====================================================================
    // STEP 1: Create Prize_Catalog
    // ====================================================================
    try {
      if (!ss.getSheetByName('Prize_Catalog')) {
        const sheet = ss.insertSheet('Prize_Catalog');
        sheet.appendRow([
          'Code', 'Name', 'Level', 'Rarity', 'COGS', 'EV_Cost', 'Qty',
          'Eligible_Rounds', 'Eligible_End', 'Player_Threshold', 'InStock',
          'EV_Explanation', 'Round_Weight', 'PV_Multiplier', 'Projected_Qty'
        ]);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:O1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        sheet.setColumnWidth(1, 100); // Code
        sheet.setColumnWidth(2, 200); // Name
        results.push('✓ Prize_Catalog created');
      } else {
        results.push('✓ Prize_Catalog exists');
      }
    } catch (e) {
      results.push('✗ Prize_Catalog: ' + e.message);
    }

    // ====================================================================
    // STEP 2: Create Prize_Throttle
    // ====================================================================
    try {
      if (!ss.getSheetByName('Prize_Throttle')) {
        const sheet = ss.insertSheet('Prize_Throttle');
        sheet.appendRow(['Parameter', 'Value']);

        // Add default throttle parameters
        const defaults = [
          ['RL_Percentage', '0.95'],
          ['EF_Clamp_Min', '0.80'],
          ['EF_Clamp_Max', '2.25'],
          ['Consolation_L1_Ratio', '0.20'],
          ['Night_Mode_Enabled', 'FALSE'],
          ['Night_Mode_Profile', 'STANDARD'],
          ['Resolver_EV_Target', 'L1_EV'],
          ['Rainbow_Rate', '3:1'],
          ['BP_Cap_Per_Event', '20'],
          ['BP_Global_Cap', '100'],
          ['Hybrid_Cap_Enabled', 'TRUE'],
          ['RL_Red_Threshold', '0.95']
        ];

        defaults.forEach(row => sheet.appendRow(row));

        sheet.setFrozenRows(1);
        sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        sheet.setColumnWidth(1, 200);
        sheet.setColumnWidth(2, 100);
        results.push('✓ Prize_Throttle created');
      } else {
        results.push('✓ Prize_Throttle exists');
      }
    } catch (e) {
      results.push('✗ Prize_Throttle: ' + e.message);
    }

    // ====================================================================
    // STEP 3: Create Integrity_Log
    // ====================================================================
    try {
      if (!ss.getSheetByName('Integrity_Log')) {
        const sheet = ss.insertSheet('Integrity_Log');
        sheet.appendRow([
          'Timestamp', 'Event_ID', 'Action', 'Operator', 'PreferredName', 'Seed',
          'Checksum_Before', 'Checksum_After', 'RL_Band', 'DF_Tags', 'Details', 'Status'
        ]);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:L1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');

        // Add bootstrap entry
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
        const operator = Session.getActiveUser().getEmail() || 'unknown';
        sheet.appendRow([
          timestamp,
          'SYSTEM',
          'BOOTSTRAP',
          operator,
          '',
          '',
          '',
          '',
          'GREEN',
          'DF-060',
          'Mega Bootstrap v7.9.6 - Initial setup',
          'SUCCESS'
        ]);

        results.push('✓ Integrity_Log created');
      } else {
        results.push('✓ Integrity_Log exists');
      }
    } catch (e) {
      results.push('✗ Integrity_Log: ' + e.message);
    }

    // ====================================================================
    // STEP 4: Create Spent_Pool
    // ====================================================================
    try {
      if (!ss.getSheetByName('Spent_Pool')) {
        const sheet = ss.insertSheet('Spent_Pool');
        sheet.appendRow([
          'Event_ID', 'Item_Code', 'Item_Name', 'Level', 'Qty', 'COGS', 'Total',
          'Timestamp', 'Batch_ID', 'Reverted', 'Event_Type'
        ]);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        results.push('✓ Spent_Pool created');
      } else {
        results.push('✓ Spent_Pool exists');
      }
    } catch (e) {
      results.push('✗ Spent_Pool: ' + e.message);
    }

    // ====================================================================
    // STEP 5: Create Key_Tracker
    // ====================================================================
    try {
      if (!ss.getSheetByName('Key_Tracker')) {
        const sheet = ss.insertSheet('Key_Tracker');
        sheet.appendRow([
          'PreferredName', 'Red', 'Blue', 'Green', 'Yellow', 'Purple',
          'RainbowEligible', 'LastUpdated'
        ]);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:H1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        sheet.setColumnWidth(1, 150);
        results.push('✓ Key_Tracker created');
      } else {
        results.push('✓ Key_Tracker exists');
      }
    } catch (e) {
      results.push('✗ Key_Tracker: ' + e.message);
    }

    // ====================================================================
    // STEP 6: Create BP_Total (Enhanced Schema)
    // ====================================================================
    try {
      if (!ss.getSheetByName('BP_Total')) {
        const sheet = ss.insertSheet('BP_Total');
        // Enhanced schema with mission point columns
        const headers = [
          'PreferredName',
          'Current_BP',
          'Attendance Mission Points',
          'Flag Mission Points',
          'Dice Roll Points',
          'LastUpdated',
          'BP_Historical'
        ];
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#4285f4')
          .setFontColor('#ffffff');
        sheet.setColumnWidth(1, 150);
        results.push('✓ BP_Total created (enhanced schema)');
      } else {
        results.push('✓ BP_Total exists');
      }
    } catch (e) {
      results.push('✗ BP_Total: ' + e.message);
    }

    // ====================================================================
    // STEP 7: Create Attendance_Missions (Production Schema)
    // ====================================================================
    try {
      if (!ss.getSheetByName('Attendance_Missions')) {
        const sheet = ss.insertSheet('Attendance_Missions');
        // Production schema with all mission badges
        const headers = [
          'PreferredName',
          'Attendance Mission Points',
          'First Contact',
          'Stellar Explorer',
          'Deck Diver',
          'Lunar Loyalty',
          'Meteor Shower',
          'Sealed Voyager',
          'Draft Navigator',
          'Stellar Scholar',
          'cEDH Events',
          'Limited Events',
          'Academy Events',
          'Outreach Events',
          'Free Play Events',
          'Interstellar Strategist',
          'Black Hole Survivor',
          'LastUpdated'
        ];
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#4285f4')
          .setFontColor('#ffffff');
        sheet.setColumnWidth(1, 150);
        results.push('✓ Attendance_Missions created');
      } else {
        results.push('✓ Attendance_Missions exists');
      }
    } catch (e) {
      results.push('✗ Attendance_Missions: ' + e.message);
    }

    // ====================================================================
    // STEP 7b: Create Flag_Missions
    // ====================================================================
    try {
      if (!ss.getSheetByName('Flag_Missions')) {
        const sheet = ss.insertSheet('Flag_Missions');
        // Production schema with all flag missions
        const headers = [
          'PreferredName',
          'Flag Mission Points',
          'Cosmic_Selfie',
          'Review_Writer',
          'Social_Media_Star',
          'App_Explorer',
          'Cosmic_Merchant',
          'Precon_Pioneer',
          'Gravitational_Pull',
          'Rogue_Planet',
          'Quantum_Collector',
          'LastUpdated'
        ];
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#4285f4')
          .setFontColor('#ffffff');
        sheet.setColumnWidth(1, 150);
        results.push('✓ Flag_Missions created');
      } else {
        results.push('✓ Flag_Missions exists');
      }
    } catch (e) {
      results.push('✗ Flag_Missions: ' + e.message);
    }

    // ====================================================================
    // STEP 7c: Create Dice Roll Points
    // ====================================================================
    try {
      if (!ss.getSheetByName('Dice Roll Points')) {
        const sheet = ss.insertSheet('Dice Roll Points');
        const headers = ['PreferredName', 'Dice Roll Points', 'LastUpdated'];
        sheet.appendRow(headers);
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, headers.length)
          .setFontWeight('bold')
          .setBackground('#4285f4')
          .setFontColor('#ffffff');
        sheet.setColumnWidth(1, 150);
        results.push('✓ Dice Roll Points created');
      } else {
        results.push('✓ Dice Roll Points exists');
      }
    } catch (e) {
      results.push('✗ Dice Roll Points: ' + e.message);
    }

    // ====================================================================
    // STEP 8: Create Players_Prize-Wall-Points
    // ====================================================================
    try {
      if (!ss.getSheetByName('Players_Prize-Wall-Points')) {
        const sheet = ss.insertSheet('Players_Prize-Wall-Points');
        sheet.appendRow([
          'PreferredName', 'Dice_Points_Available', 'Dice_Points_Spent',
          'Last_Event', 'LastUpdated'
        ]);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        results.push('✓ Players_Prize-Wall-Points created');
      } else {
        results.push('✓ Players_Prize-Wall-Points exists');
      }
    } catch (e) {
      results.push('✗ Players_Prize-Wall-Points: ' + e.message);
    }

    // ====================================================================
    // STEP 9: Create Preview_Artifacts (hidden)
    // ====================================================================
    try {
      if (!ss.getSheetByName('Preview_Artifacts')) {
        const sheet = ss.insertSheet('Preview_Artifacts');
        sheet.appendRow([
          'Artifact_ID', 'Event_ID', 'Seed', 'Preview_Hash', 'Created_At', 'Expires_At'
        ]);
        sheet.setFrozenRows(1);
        sheet.hideSheet();
        results.push('✓ Preview_Artifacts created (hidden)');
      } else {
        results.push('✓ Preview_Artifacts exists');
      }
    } catch (e) {
      results.push('✗ Preview_Artifacts: ' + e.message);
    }

    // ====================================================================
    // STEP 10: Create optional sheets
    // ====================================================================
    try {
      if (!ss.getSheetByName('Event_Outcomes')) {
        const sheet = ss.insertSheet('Event_Outcomes');
        sheet.appendRow(['PreferredName', 'R1_Result', 'R2_Result', 'R3_Result', 'Notes']);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        results.push('✓ Event_Outcomes created');
      }
    } catch (e) {
      // Optional - don't report error
    }

    try {
      if (!ss.getSheetByName('Prestige_Overflow')) {
        const sheet = ss.insertSheet('Prestige_Overflow');
        sheet.appendRow(['PreferredName', 'Total_Overflow', 'Last_Updated', 'Prestige_Tier']);
        sheet.setFrozenRows(1);
        sheet.getRange('A1:D1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
        results.push('✓ Prestige_Overflow created');
      }
    } catch (e) {
      // Optional - don't report error
    }

    // ====================================================================
    // SUCCESS!
    // ====================================================================
    results.push('');
    results.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    results.push('✅ BOOTSTRAP COMPLETE!');
    results.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    results.push('');
    results.push('📋 Sheets Created:');
    results.push('   • Prize_Catalog');
    results.push('   • Prize_Throttle');
    results.push('   • Integrity_Log');
    results.push('   • Spent_Pool');
    results.push('   • Key_Tracker');
    results.push('   • BP_Total (enhanced schema)');
    results.push('   • Attendance_Missions');
    results.push('   • Flag_Missions');
    results.push('   • Dice Roll Points');
    results.push('   • Players_Prize-Wall-Points');
    results.push('   • Preview_Artifacts (hidden)');
    results.push('');
    results.push('📝 Next Steps:');
    results.push('1. Add service .gs files to Apps Script');
    results.push('2. Add HTML ui/*.html files to Apps Script');
    results.push('3. Save and refresh your sheet');
    results.push('4. Menu will appear: "Cosmic Tournament v7.9"');
    results.push('');
    results.push('📚 See INSTALL.md for detailed instructions');
    results.push('');
    results.push('🎉 You\'re ready to start adding content!');

    ui.alert(
      '✅ Bootstrap Complete!',
      results.join('\n'),
      ui.ButtonSet.OK
    );

  } catch (e) {
    ui.alert(
      '❌ Bootstrap Error',
      'Error: ' + e.message + '\n\n' +
      'Stack: ' + e.stack + '\n\n' +
      'Please report this error.',
      ui.ButtonSet.OK
    );
    console.error('Mega Bootstrap failed:', e);
  }
}

/**
 * Simple health check - run this to see what sheets exist
 */
function megaBootstrapCheckHealth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const requiredSheets = [
    'Prize_Catalog',
    'Prize_Throttle',
    'Integrity_Log',
    'Spent_Pool',
    'Key_Tracker',
    'BP_Total',
    'Attendance_Missions',
    'Flag_Missions',
    'Dice Roll Points',
    'Players_Prize-Wall-Points',
    'Preview_Artifacts'
  ];

  const results = [];
  results.push('📊 System Health Check');
  results.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  results.push('');

  let allGood = true;

  requiredSheets.forEach(sheetName => {
    const exists = ss.getSheetByName(sheetName) !== null;
    results.push((exists ? '✓' : '✗') + ' ' + sheetName);
    if (!exists) allGood = false;
  });

  results.push('');
  results.push('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

  if (allGood) {
    results.push('✅ All required sheets exist!');
    results.push('');
    results.push('Next: Add service files and HTML files');
  } else {
    results.push('⚠️ Some sheets missing');
    results.push('');
    results.push('Run: megaBootstrapInstall()');
  }

  ui.alert('Health Check', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * Add sample data to Prize_Catalog for testing
 */
function megaBootstrapAddSamplePrizes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Catalog');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error', 'Prize_Catalog not found. Run megaBootstrapInstall first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Sample prizes
  const samplePrizes = [
    ['PROMO001', 'Premium Promo Card', 'L3', 'Rare', 5.00, 15.00, 10, 'FALSE', 'TRUE', 8, 'TRUE', 'High value promo', 1.0, 1.0, 0],
    ['BOOST001', 'Booster Pack', 'L2', 'Common', 3.00, 8.00, 50, 'TRUE', 'TRUE', 4, 'TRUE', 'Standard booster', 1.0, 1.0, 0],
    ['DECK001', 'Starter Deck', 'L2', 'Uncommon', 8.00, 20.00, 20, 'FALSE', 'TRUE', 8, 'TRUE', 'Complete deck', 1.0, 1.0, 0],
    ['SLEEP001', 'Card Sleeves', 'L1', 'Common', 2.00, 5.00, 100, 'TRUE', 'TRUE', 1, 'TRUE', 'Protective sleeves', 1.0, 1.0, 0],
    ['DICE001', 'Dice Set', 'L1', 'Common', 1.50, 4.00, 75, 'TRUE', 'TRUE', 1, 'TRUE', 'Six-sided dice', 1.0, 1.0, 0],
    ['PIN001', 'Collector Pin', 'L0', 'Common', 0.50, 2.00, 200, 'TRUE', 'TRUE', 1, 'TRUE', 'Event pin', 1.0, 1.0, 0],
    ['STICK001', 'Sticker Pack', 'L0', 'Common', 0.25, 1.00, 300, 'TRUE', 'TRUE', 1, 'TRUE', 'Promotional stickers', 1.0, 1.0, 0],
    ['PLAY001', 'Playmat', 'L3', 'Rare', 12.00, 30.00, 15, 'FALSE', 'TRUE', 12, 'TRUE', 'Tournament playmat', 1.0, 1.0, 0],
    ['BOX001', 'Storage Box', 'L1', 'Uncommon', 3.50, 8.00, 40, 'TRUE', 'TRUE', 1, 'TRUE', 'Card storage', 1.0, 1.0, 0],
    ['BINDER001', 'Card Binder', 'L2', 'Uncommon', 6.00, 15.00, 25, 'FALSE', 'TRUE', 8, 'TRUE', '9-pocket binder', 1.0, 1.0, 0]
  ];

  // Add sample data
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, samplePrizes.length, samplePrizes[0].length).setValues(samplePrizes);

  SpreadsheetApp.getUi().alert(
    '✓ Sample Data Added',
    'Added 10 sample prizes to Prize_Catalog.\n\n' +
    'You can now test event creation and prize allocation!\n\n' +
    'Levels: L0 (Consolation), L1-L2 (Standard), L3 (Premium)',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Add sample canonical players for testing
 */
function megaBootstrapAddSamplePlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");

  const samplePlayers = [
    { name: 'Alice Smith', att: 15, flag: 8, dice: 5 },
    { name: 'Bob Johnson', att: 10, flag: 12, dice: 3 },
    { name: 'Carol Davis', att: 20, flag: 5, dice: 10 },
    { name: 'Dave Wilson', att: 8, flag: 15, dice: 7 },
    { name: 'Emma Brown', att: 12, flag: 10, dice: 8 }
  ];

  // Add to Key_Tracker
  const keySheet = ss.getSheetByName('Key_Tracker');
  if (keySheet) {
    samplePlayers.forEach(p => {
      keySheet.appendRow([p.name, 0, 0, 0, 0, 0, 0, timestamp]);
    });
  }

  // Add to BP_Total (enhanced schema)
  const bpSheet = ss.getSheetByName('BP_Total');
  if (bpSheet) {
    samplePlayers.forEach(p => {
      const total = p.att + p.flag + p.dice;
      // PreferredName, Current_BP, Attendance Mission Points, Flag Mission Points, Dice Roll Points, LastUpdated, BP_Historical
      bpSheet.appendRow([p.name, total, p.att, p.flag, p.dice, timestamp, total]);
    });
  }

  // Add to Attendance_Missions
  const attSheet = ss.getSheetByName('Attendance_Missions');
  if (attSheet) {
    samplePlayers.forEach(p => {
      // PreferredName, Attendance Mission Points, then placeholder zeros for badges, LastUpdated
      const row = [p.name, p.att, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, timestamp];
      attSheet.appendRow(row);
    });
  }

  // Add to Flag_Missions
  const flagSheet = ss.getSheetByName('Flag_Missions');
  if (flagSheet) {
    samplePlayers.forEach(p => {
      // PreferredName, Flag Mission Points, then placeholder zeros for flags, LastUpdated
      const row = [p.name, p.flag, 0, 0, 0, 0, 0, 0, 0, 0, 0, timestamp];
      flagSheet.appendRow(row);
    });
  }

  // Add to Dice Roll Points
  const diceSheet = ss.getSheetByName('Dice Roll Points');
  if (diceSheet) {
    samplePlayers.forEach(p => {
      diceSheet.appendRow([p.name, p.dice, timestamp]);
    });
  }

  SpreadsheetApp.getUi().alert(
    '✓ Sample Players Added',
    'Added 5 sample players with mission points:\n' +
    '• Alice Smith (Att: 15, Flag: 8, Dice: 5 = 28 BP)\n' +
    '• Bob Johnson (Att: 10, Flag: 12, Dice: 3 = 25 BP)\n' +
    '• Carol Davis (Att: 20, Flag: 5, Dice: 10 = 35 BP)\n' +
    '• Dave Wilson (Att: 8, Flag: 15, Dice: 7 = 30 BP)\n' +
    '• Emma Brown (Att: 12, Flag: 10, Dice: 8 = 30 BP)\n\n' +
    'They appear in Key_Tracker, BP_Total, and all mission sheets.\n\n' +
    'Test the Refresh Bonus Points function to sync from sources!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Complete setup: sheets + sample data
 */
function megaBootstrapCompleteSetup() {
  megaBootstrapInstall();
  Utilities.sleep(2000); // Wait 2 seconds
  megaBootstrapAddSamplePrizes();
  Utilities.sleep(1000); // Wait 1 second
  megaBootstrapAddSamplePlayers();

  SpreadsheetApp.getUi().alert(
    '🎉 Complete Setup Done!',
    'All sheets created + sample data added!\n\n' +
    'You now have:\n' +
    '✓ All required sheets\n' +
    '✓ 10 sample prizes in catalog\n' +
    '✓ 5 sample players in Key_Tracker/BP_Total\n\n' +
    'Next: Add service files and HTML files to Apps Script,\n' +
    'then you can start creating events!\n\n' +
    'See INSTALL.md for next steps.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}