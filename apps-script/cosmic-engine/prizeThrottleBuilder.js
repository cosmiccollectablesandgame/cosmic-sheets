/**
 * Prize_Throttle Sheet Builder
 * @fileoverview Creates Prize_Throttle sheet with all sections and named ranges
 */

/**
 * Creates or rebuilds Prize_Throttle sheet
 * @param {boolean} rebuild - If true, deletes and recreates existing sheet
 * @return {Sheet} Created sheet
 */
function createPrizeThrottleSheet(rebuild = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prize_Throttle');

  // Delete if rebuild requested
  if (rebuild && sheet) {
    ss.deleteSheet(sheet);
    sheet = null;
  }

  // Create sheet if doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet('Prize_Throttle');
  } else {
    sheet.clear();
  }

  // Build all sections
  buildHeader_(sheet);
  buildBudgetRLControls_(sheet);
  buildCommanderTemplate_(sheet);
  buildValidationPanel_(sheet);
  buildFloorsThresholds_(sheet);
  buildRarityWeights_(sheet);
  buildConsolationControls_(sheet);
  buildEconomyControls_(sheet);

  // Create named ranges
  createNamedRanges_(sheet);

  // Apply formatting
  applyFormatting_(sheet);

  Logger.log('Prize_Throttle sheet created successfully');

  return sheet;
}

/**
 * Builds header section (A1:L5)
 * @private
 */
function buildHeader_(sheet) {
  sheet.getRange('A1').setValue('PRIZE_THROTTLE - COMMANDER PRIZE CONTROL CENTER');
  sheet.getRange('A2').setValue('Cosmic Event Manager v7.9.6');
  sheet.getRange('A4').setValue('System Status');
  sheet.getRange('B4').setValue('Ready');
  sheet.getRange('D4').setValue('Last Validated');
  sheet.getRange('E4').setValue(new Date());
  sheet.getRange('G4').setValue('Version');
  sheet.getRange('H4').setValue('1.0.0');

  sheet.getRange('A1').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A2').setFontSize(10).setFontColor('#666666');
}

/**
 * Builds Budget & RL Controls section (A7:L20)
 * @private
 */
function buildBudgetRLControls_(sheet) {
  sheet.getRange('A7').setValue('═══ BUDGET & RL CONTROLS ═══').setFontWeight('bold');

  sheet.getRange('A10').setValue('RL_Baseline_%');
  sheet.getRange('B10').setValue(0.95);
  sheet.getRange('C10').setValue('[LOCKED - DF-010]').setFontColor('#999999');

  sheet.getRange('A15').setValue('RL_Dial_%');
  sheet.getRange('B15').setValue(0.95);
  sheet.getRange('C15').setValue('[EDITABLE]').setFontColor('#0066cc');

  sheet.getRange('A22').setValue('Auto-Trim on Red');
  sheet.getRange('B22').setValue(true);

  sheet.getRange('A23').setValue('Auto-Trim Target');
  sheet.getRange('B23').setValue(0.90);

  sheet.getRange('A24').setValue('Log Bath Events');
  sheet.getRange('B24').setValue(true);

  // Lock RL_Baseline
  const protection = sheet.getRange('B10').protect();
  protection.setDescription('RL Baseline is immutable per DF-010');
}

/**
 * Builds Commander Round Prize Template (A22:F32)
 * @private
 */
function buildCommanderTemplate_(sheet) {
  sheet.getRange('A27').setValue('═══ COMMANDER ROUND PRIZE TEMPLATE ═══').setFontWeight('bold');

  // Headers
  sheet.getRange('A29').setValue('Rank');
  sheet.getRange('B29').setValue('Round 1');
  sheet.getRange('C29').setValue('Round 2');
  sheet.getRange('D29').setValue('Round 3');
  sheet.getRange('E29').setValue('Row Total');
  sheet.getRange('F29').setValue('Notes');

  // Rank labels
  sheet.getRange('A30').setValue('Rank 1');
  sheet.getRange('A31').setValue('Rank 2');
  sheet.getRange('A32').setValue('Rank 3');
  sheet.getRange('A33').setValue('Rank 4');

  // Default template values
  const templateValues = [
    ['L2', 'L2', 'L2'],
    ['L1', 'L0', 'L1'],
    ['L0', 'L0', 'L0'],
    ['L1', 'L0', 'L1']
  ];
  sheet.getRange(30, 2, 4, 3).setValues(templateValues);

  // Row total formulas
  sheet.getRange('E30').setFormula('=GET_LEVEL_AVG_COGS(B30)+GET_LEVEL_AVG_COGS(C30)+GET_LEVEL_AVG_COGS(D30)');
  sheet.getRange('E31').setFormula('=GET_LEVEL_AVG_COGS(B31)+GET_LEVEL_AVG_COGS(C31)+GET_LEVEL_AVG_COGS(D31)');
  sheet.getRange('E32').setFormula('=GET_LEVEL_AVG_COGS(B32)+GET_LEVEL_AVG_COGS(C32)+GET_LEVEL_AVG_COGS(D32)');
  sheet.getRange('E33').setFormula('=GET_LEVEL_AVG_COGS(B33)+GET_LEVEL_AVG_COGS(C33)+GET_LEVEL_AVG_COGS(D33)');

  // Grand total
  sheet.getRange('A35').setValue('Grand Total');
  sheet.getRange('E35').setFormula('=SUM(E30:E33)');

  // RL %
  sheet.getRange('A36').setValue('RL %');
  sheet.getRange('E36').setFormula('=E35/$H$40');

  // RL Band
  sheet.getRange('A37').setValue('RL Band');
  sheet.getRange('E37').setFormula('=IF(E36<=0.90,"GREEN",IF(E36<=0.95,"AMBER","RED"))');

  // Data validation for template cells
  const levelList = ['L0','L1','L2','L3','L4','L5','L6','L7','L8','L9','L10','L11','L12','L13'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(levelList, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B30:D33').setDataValidation(rule);

  // Conditional formatting for RL Band
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('GREEN')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('E37')])
    .build();

  const amberRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('AMBER')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('E37')])
    .build();

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('RED')
    .setBackground('#f4cccc')
    .setRanges([sheet.getRange('E37')])
    .build();

  const rules = sheet.getConditionalFormatRules();
  rules.push(greenRule, amberRule, redRule);
  sheet.setConditionalFormatRules(rules);
}

/**
 * Builds Validation Panel (G22:L40)
 * @private
 */
function buildValidationPanel_(sheet) {
  sheet.getRange('G27').setValue('═══ VALIDATION PANEL ═══').setFontWeight('bold');

  sheet.getRange('G29').setValue('Test Player Count');
  sheet.getRange('H29').setValue(8);

  sheet.getRange('G30').setValue('Test Entry Fee');
  sheet.getRange('H30').setValue(5.00);

  sheet.getRange('G32').setValue('Eligible Net');
  sheet.getRange('H32').setFormula('=H29*H30');

  sheet.getRange('G33').setValue('RL Baseline (95%)');
  sheet.getRange('H33').setFormula('=H32*0.95');

  sheet.getRange('G34').setValue('RL Dial Cap');
  sheet.getRange('H34').setFormula('=H32*$B$15');

  sheet.getRange('G35').setValue('Active Cap');
  sheet.getRange('H35').setFormula('=MIN(H33,H34)');

  sheet.getRange('G37').setValue('Template COGS');
  sheet.getRange('H37').setFormula('=E35');

  sheet.getRange('G38').setValue('RL Usage');
  sheet.getRange('H38').setFormula('=H37/H35');

  sheet.getRange('G39').setValue('Remaining');
  sheet.getRange('H39').setFormula('=H35-H37');

  sheet.getRange('G40').setValue('Status');
  sheet.getRange('H40').setFormula('=E37');

  // Format as currency and percentages
  sheet.getRange('H30').setNumberFormat('$#,##0.00');
  sheet.getRange('H32:H35').setNumberFormat('$#,##0.00');
  sheet.getRange('H37').setNumberFormat('$#,##0.00');
  sheet.getRange('H38').setNumberFormat('0.0%');
  sheet.getRange('H39').setNumberFormat('$#,##0.00');
}

/**
 * Builds Floors & Thresholds (A42:L58)
 * @private
 */
function buildFloorsThresholds_(sheet) {
  sheet.getRange('A42').setValue('═══ FLOORS & THRESHOLDS ═══').setFontWeight('bold');

  const floors = [
    ['Floor_Commander_Fire', 2],
    ['Floor_Commander_Full', 8],
    ['Threshold_L3', 8],
    ['Threshold_Promo_Reg', 8],
    ['Threshold_Promo_Foil', 12],
    ['Threshold_L4', 12],
    ['Threshold_Merch_1', 20],
    ['Threshold_Sealed_24', 24],
    ['Threshold_Big_Bump', 28],
    ['Threshold_Merch_2', 32]
  ];

  let row = 44;
  for (const [name, value] of floors) {
    sheet.getRange(row, 1).setValue(name);
    sheet.getRange(row, 2).setValue(value);
    row++;
  }
}

/**
 * Builds Rarity Weights (A60:L68)
 * @private
 */
function buildRarityWeights_(sheet) {
  sheet.getRange('A60').setValue('═══ RARITY WEIGHTS ═══').setFontWeight('bold');

  const weights = [
    ['Rarity_Weight_01', 0.5],
    ['Rarity_Weight_23', 1.0],
    ['Rarity_Weight_46', 2.0],
    ['Rarity_Weight_79', 3.5],
    ['Rarity_Weight_10', 5.0]
  ];

  let row = 62;
  for (const [name, value] of weights) {
    sheet.getRange(row, 1).setValue(name);
    sheet.getRange(row, 2).setValue(value);
    row++;
  }

  sheet.getRange('A68').setValue('Rarity_Weighting_Active');
  sheet.getRange('B68').setValue(true);

  sheet.getRange('A70').setValue('Allow_Duplicate_Items');
  sheet.getRange('B70').setValue(false);

  sheet.getRange('A71').setValue('Prefer_InStock');
  sheet.getRange('B71').setValue(true);
}

/**
 * Builds Consolation Controls (A73:L82)
 * @private
 */
function buildConsolationControls_(sheet) {
  sheet.getRange('A73').setValue('═══ CONSOLATION CONTROLS ═══').setFontWeight('bold');

  sheet.getRange('A75').setValue('Consolation_Ceiling_Mode');
  sheet.getRange('B75').setValue('STRICT');

  sheet.getRange('A77').setValue('Consolation_L1_Pct');
  sheet.getRange('B77').setValue(0.20);

  sheet.getRange('A78').setValue('Consolation_L0_Pct');
  sheet.getRange('B78').setValue(0.80);

  sheet.getRange('A80').setValue('Auto_Trim_Consolation');
  sheet.getRange('B80').setValue(true);

  sheet.getRange('A82').setValue('Fixed_L2_Seats_Enabled');
  sheet.getRange('B82').setValue(true);
}

/**
 * Builds Economy Controls (A84:L92)
 * @private
 */
function buildEconomyControls_(sheet) {
  sheet.getRange('A84').setValue('═══ ECONOMY CONTROLS ═══').setFontWeight('bold');

  sheet.getRange('A86').setValue('EF_Cap_Pct_of_RL');
  sheet.getRange('B86').setValue(0.35);

  sheet.getRange('A87').setValue('BP_Cap');
  sheet.getRange('B87').setValue(100);

  sheet.getRange('A88').setValue('BP_Prestige_Overflow');
  sheet.getRange('B88').setValue(true);

  sheet.getRange('A90').setValue('Rainbow_Conversion_Rate');
  sheet.getRange('B90').setValue('3:1');

  sheet.getRange('A92').setValue('Hybrid_Cap_Enabled');
  sheet.getRange('B92').setValue(true);
}

/**
 * Creates named ranges for Prize_Throttle
 * @private
 */
function createNamedRanges_(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ranges = {
    'RL_Baseline_Pct': 'B10',
    'RL_Dial_Pct': 'B15',
    'Auto_Trim_Enabled': 'B22',
    'Auto_Trim_Target': 'B23',
    'Log_Bath_Events': 'B24',
    'Commander_Round_Template': 'B30:D33',
    'Test_Player_Count': 'H29',
    'Test_Entry_Fee': 'H30',
    'Active_RL_Cap': 'H35',
    'Floor_Commander_Fire': 'B44',
    'Floor_Commander_Full': 'B45',
    'Threshold_L3': 'B46',
    'Threshold_Promo_Reg': 'B47',
    'Threshold_Promo_Foil': 'B48',
    'Threshold_L4': 'B49',
    'Threshold_Merch_1': 'B50',
    'Threshold_Sealed_24': 'B51',
    'Threshold_Big_Bump': 'B52',
    'Threshold_Merch_2': 'B53',
    'Rarity_Weight_01': 'B62',
    'Rarity_Weight_23': 'B63',
    'Rarity_Weight_46': 'B64',
    'Rarity_Weight_79': 'B65',
    'Rarity_Weight_10': 'B66',
    'Rarity_Weighting_Active': 'B68',
    'Allow_Duplicate_Items': 'B70',
    'Prefer_InStock': 'B71'
  };

  for (const [name, a1Notation] of Object.entries(ranges)) {
    try {
      const existingRange = ss.getRangeByName(name);
      if (existingRange) {
        ss.removeNamedRange(name);
      }
    } catch (e) {
      // Range doesn't exist, continue
    }

    ss.setNamedRange(name, sheet.getRange(a1Notation));
  }

  Logger.log(`Created ${Object.keys(ranges).length} named ranges`);
}

/**
 * Applies formatting to Prize_Throttle sheet
 * @private
 */
function applyFormatting_(sheet) {
  // Header formatting
  sheet.getRange('A1:L1').setBackground('#1a73e8').setFontColor('#ffffff');

  // Section headers
  const sectionHeaders = ['A7', 'A27', 'G27', 'A42', 'A60', 'A73', 'A84'];
  sectionHeaders.forEach(cell => {
    sheet.getRange(cell).setBackground('#e8eaf6').setFontWeight('bold');
  });

  // Template section
  sheet.getRange('A29:F29').setBackground('#d9d9d9').setFontWeight('bold');
  sheet.getRange('B30:D33').setBackground('#f3f3f3');

  // Validation panel
  sheet.getRange('G29:H29').setBackground('#d9d9d9').setFontWeight('bold');

  // Freeze header rows
  sheet.setFrozenRows(5);

  // Auto-resize columns
  sheet.autoResizeColumns(1, 12);

  Logger.log('Formatting applied to Prize_Throttle');
}

/**
 * Creates Bath_Log sheet
 * @return {Sheet} Created sheet
 */
function createBathLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Bath_Log');

  if (sheet) {
    return sheet; // Already exists
  }

  sheet = ss.insertSheet('Bath_Log');

  const headers = [
    'Timestamp', 'Event_ID', 'Event_Date', 'Format', 'Player_Count', 'Entry_Fee',
    'RL_Baseline_95', 'RL_Dial_%', 'RL_Dial_$', 'Preview_Prize_COGS', 'RL_Usage_%',
    'RLbath_$', 'RLbath_%', 'Was_Trimmed', 'Trim_Amount_$', 'Final_Prize_COGS',
    'RL_Final_%', 'RL_Band_Final', 'df_tags', 'Seed', 'Preview_Hash', 'Commit_Hash', 'Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  // Create named range
  ss.setNamedRange('Bath_Log_Data', sheet.getRange('A2:W'));

  Logger.log('Bath_Log sheet created');

  return sheet;
}