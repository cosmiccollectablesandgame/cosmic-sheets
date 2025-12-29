/**
 * Throttle Service - Configuration Management
 * @fileoverview Manages Prize_Throttle (KV store) for global parameters
 */
// ============================================================================
// THROTTLE KV OPERATIONS
// ============================================================================
/**
 * Default throttle parameters
 */
const THROTTLE_DEFAULTS = {
  'RL_Percentage': '0.95',
  'EF_Clamp_Min': '0.80',
  'EF_Clamp_Max': '2.25',
  'Consolation_L1_Ratio': '0.20',
  'Night_Mode_Enabled': 'FALSE',
  'Night_Mode_Profile': 'STANDARD',
  'Resolver_EV_Target': 'L1_EV',
  'Rainbow_Rate': '3:1',
  'BP_Cap_Per_Event': '20',
  'BP_Global_Cap': '100',
  'Hybrid_Cap_Enabled': 'TRUE',
  'RL_Red_Threshold': '0.95'
};
/**
 * Gets all throttle parameters as KV object
 * @return {Object} Throttle parameters
 */
function getThrottleKV() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prize_Throttle');
  // Create sheet with defaults if missing
  if (!sheet) {
    sheet = createThrottleSheet_();
  }
  const data = sheet.getDataRange().getValues();
  const kv = toKV(data);
  // Fill missing defaults
  Object.keys(THROTTLE_DEFAULTS).forEach(key => {
    if (kv[key] === undefined || kv[key] === '') {
      kv[key] = THROTTLE_DEFAULTS[key];
    }
  });
  return kv;
}
/**
 * Sets throttle parameters (validates first)
 * @param {Object} updates - KV updates
 * @throws {Error} If validation fails
 */
function setThrottleKV(updates) {
  // Validate updates
  const errors = validateThrottleUpdates_(updates);
  if (errors.length > 0) {
    throwError('Invalid throttle parameters', 'THROTTLE_INVALID', errors.join('; '));
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prize_Throttle');
  if (!sheet) {
    sheet = createThrottleSheet_();
  }
  const data = sheet.getDataRange().getValues();
  const existingKV = toKV(data);
  // Apply updates
  Object.keys(updates).forEach(key => {
    existingKV[key] = updates[key];
  });
  // Write back to sheet
  const rows = [['Parameter', 'Value']];
  Object.keys(existingKV).forEach(key => {
    rows.push([key, existingKV[key]]);
  });
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  // Log change
  logIntegrityAction('THROTTLE_CHANGE', {
    details: `Updated: ${Object.keys(updates).join(', ')}`,
    status: 'SUCCESS'
  });
}
/**
 * Gets a single throttle parameter
 * @param {string} key - Parameter key
 * @param {*} defaultValue - Default if not found
 * @return {*} Parameter value
 */
function getThrottleParam(key, defaultValue = null) {
  const kv = getThrottleKV();
  return kv[key] !== undefined ? kv[key] : defaultValue;
}
/**
 * Sets a single throttle parameter
 * @param {string} key - Parameter key
 * @param {*} value - Parameter value
 */
function setThrottleParam(key, value) {
  setThrottleKV({ [key]: value });
}
// ============================================================================
// VALIDATION
// ============================================================================
/**
 * Validates throttle updates
 * @param {Object} updates - KV updates
 * @return {Array<string>} Array of error messages (empty if valid)
 * @private
 */
function validateThrottleUpdates_(updates) {
  const errors = [];
  // RL_Percentage: 0.0 - 1.0
  if (updates.RL_Percentage !== undefined) {
    const val = parseFloat(updates.RL_Percentage);
    if (isNaN(val) || val < 0 || val > 1) {
      errors.push('RL_Percentage must be between 0.0 and 1.0');
    }
  }
  // EF_Clamp_Min: 0.5 - 2.0
  if (updates.EF_Clamp_Min !== undefined) {
    const val = parseFloat(updates.EF_Clamp_Min);
    if (isNaN(val) || val < 0.5 || val > 2.0) {
      errors.push('EF_Clamp_Min must be between 0.5 and 2.0');
    }
  }
  // EF_Clamp_Max: 1.0 - 3.0
  if (updates.EF_Clamp_Max !== undefined) {
    const val = parseFloat(updates.EF_Clamp_Max);
    if (isNaN(val) || val < 1.0 || val > 3.0) {
      errors.push('EF_Clamp_Max must be between 1.0 and 3.0');
    }
  }
  // EF_Clamp_Min < EF_Clamp_Max
  if (updates.EF_Clamp_Min && updates.EF_Clamp_Max) {
    const min = parseFloat(updates.EF_Clamp_Min);
    const max = parseFloat(updates.EF_Clamp_Max);
    if (min >= max) {
      errors.push('EF_Clamp_Min must be less than EF_Clamp_Max');
    }
  }
  // Consolation_L1_Ratio: 0.0 - 1.0
  if (updates.Consolation_L1_Ratio !== undefined) {
    const val = parseFloat(updates.Consolation_L1_Ratio);
    if (isNaN(val) || val < 0 || val > 1) {
      errors.push('Consolation_L1_Ratio must be between 0.0 and 1.0');
    }
  }
  // Rainbow_Rate: "X:1" format
  if (updates.Rainbow_Rate !== undefined) {
    const parsed = parseRatio(updates.Rainbow_Rate);
    if (!parsed) {
      errors.push('Rainbow_Rate must be in format "X:1" (e.g., "3:1")');
    } else if (parsed.den !== 1) {
      errors.push('Rainbow_Rate denominator must be 1');
    }
  }
  // BP_Cap_Per_Event: 0 - 50
  if (updates.BP_Cap_Per_Event !== undefined) {
    const val = parseInt(updates.BP_Cap_Per_Event, 10);
    if (isNaN(val) || val < 0 || val > 50) {
      errors.push('BP_Cap_Per_Event must be between 0 and 50');
    }
  }
  // BP_Global_Cap: 0 - 200
  if (updates.BP_Global_Cap !== undefined) {
    const val = parseInt(updates.BP_Global_Cap, 10);
    if (isNaN(val) || val < 0 || val > 200) {
      errors.push('BP_Global_Cap must be between 0 and 200');
    }
  }
  // Night_Mode_Profile: STANDARD or ENHANCED
  if (updates.Night_Mode_Profile !== undefined) {
    const val = String(updates.Night_Mode_Profile).toUpperCase();
    if (!['STANDARD', 'ENHANCED'].includes(val)) {
      errors.push('Night_Mode_Profile must be STANDARD or ENHANCED');
    }
  }
  return errors;
}
// ============================================================================
// DERIVED BUDGET
// ============================================================================
/**
 * Computes derived RL95 budget for an event
 * @param {Object} eventProps - Event metadata
 * @param {string} eventProps.event_type - CONSTRUCTED or LIMITED
 * @param {number} eventProps.entry - Entry fee
 * @param {number} eventProps.kit_cost_per_player - Kit cost (LIMITED only)
 * @param {number} playerCount - Number of players
 * @return {Object} {budget, rl_percentage, hybrid_cap}
 */
function derivedBudgetForEvent(eventProps, playerCount) {
  const throttle = getThrottleKV();
  const rlPercent = parseFloat(throttle.RL_Percentage || 0.95);
  const hybridCapEnabled = coerceBoolean(throttle.Hybrid_Cap_Enabled);
  let eligibleNet = 0;
  if (eventProps.event_type === 'LIMITED') {
    // LIMITED: (Entry - Kit Cost) × Player Count × 0.95
    const entry = eventProps.entry || 0;
    const kitCost = eventProps.kit_cost_per_player || 0;
    eligibleNet = (entry - kitCost) * playerCount * rlPercent;
  } else {
    // CONSTRUCTED: Entry × Player Count × 0.95
    const entry = eventProps.entry || 0;
    eligibleNet = entry * playerCount * rlPercent;
  }
  // Hybrid cap: min(RL × 0.10, 15.00)
  let hybridCap = 0;
  if (hybridCapEnabled && eventProps.event_type === 'HYBRID') {
    hybridCap = Math.min(eligibleNet * 0.10, 15.00);
  }
  return {
    budget: eligibleNet,
    rl_percentage: rlPercent,
    hybrid_cap: hybridCap
  };
}
// ============================================================================
// HELPERS
// ============================================================================
/**
 * Creates Prize_Throttle sheet with defaults
 * @return {Sheet} Created sheet
 * @private
 */
function createThrottleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet('Prize_Throttle');
  const rows = [['Parameter', 'Value']];
  Object.keys(THROTTLE_DEFAULTS).forEach(key => {
    rows.push([key, THROTTLE_DEFAULTS[key]]);
  });
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  // Format
  sheet.setFrozenRows(1);
  sheet.getRange('A1:B1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
  sheet.autoResizeColumns(1, 2);
  logIntegrityAction('THROTTLE_CREATE', {
    details: 'Created Prize_Throttle with defaults',
    status: 'SUCCESS'
  });
  return sheet;
}

/**
 * Builds/rebuilds Prize_Throttle sheet with full Commander template
 * @param {boolean} rebuild - If true, clears and rebuilds entire sheet
 * @return {Sheet} Prize_Throttle sheet
 */
function buildPrizeThrottleSheet(rebuild = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prize_Throttle');

  if (rebuild && sheet) {
    ss.deleteSheet(sheet);
    sheet = null;
  }

  if (!sheet) {
    sheet = ss.insertSheet('Prize_Throttle');
  } else {
    // Clear existing content but keep the sheet
    sheet.clear();
  }

  // ===== SECTION 1: Master Controls (Rows 1-16) =====
  sheet.getRange('A1').setValue('COSMIC PRIZE ENGINE - THROTTLE CONTROL PANEL v7.9.6');
  sheet.getRange('A1:E1').merge().setFontWeight('bold').setFontSize(14)
    .setBackground('#4285f4').setFontColor('#ffffff').setHorizontalAlignment('center');

  sheet.getRange('A2').setValue('Master Controls');
  sheet.getRange('A2:B2').merge().setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

  const controls = [
    ['Parameter', 'Value'],
    ['Commander_Rounds_Enabled', 'TRUE'],
    ['Commander_Floor', '2'],
    ['RL_Percentage', '0.95'],
    ['Default_Entry_Fee', '15'],
    ['EF_Clamp_Min', '0.80'],
    ['EF_Clamp_Max', '2.25'],
    ['Allow_Duplicates', 'FALSE'],
    ['Preview_Required', 'TRUE'],
    ['Consolation_L1_Ratio', '0.20'],
    ['Rainbow_Rate', '3:1'],
    ['BP_Cap_Per_Event', '20'],
    ['BP_Global_Cap', '100'],
    ['Hybrid_Cap_Enabled', 'TRUE']
  ];

  sheet.getRange(3, 1, controls.length, 2).setValues(controls);
  sheet.getRange('A3:B3').setFontWeight('bold').setBackground('#d9d9d9');

  // ===== SECTION 2: Commander Prize Template Grid (Rows 18-35) =====
  sheet.getRange('A18').setValue('COMMANDER PRIZE TEMPLATE');
  sheet.getRange('A18:N18').merge().setFontWeight('bold').setFontSize(12)
    .setBackground('#ea4335').setFontColor('#ffffff').setHorizontalAlignment('center');

  sheet.getRange('A19').setValue('Round Prize Configuration (Modify these to control Commander round prizes)');
  sheet.getRange('A19:N19').merge().setFontSize(10).setBackground('#f4cccc');

  // Template headers
  const templateHeaders = [
    'Round', 'Seat', 'Level', 'L0', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'L7', 'L8', 'L9', 'L10'
  ];
  sheet.getRange(20, 1, 1, templateHeaders.length).setValues([templateHeaders]);
  sheet.getRange(20, 1, 1, templateHeaders.length).setFontWeight('bold')
    .setBackground('#4285f4').setFontColor('#ffffff');

  // Template data (example configuration - can be modified by user)
  const templateData = [
    ['R1', '1st', 'L2', '', '1', '', '', '', '', '', '', '', '', ''],
    ['R2', '1st', 'L2', '', '1', '', '', '', '', '', '', '', '', ''],
    ['R2', '4th', 'L1', '', '1', '', '', '', '', '', '', '', '', ''],
    ['R3', '1st', 'L3', '', '', '', '1', '', '', '', '', '', '', ''],
    ['End', '1st', 'L4', '', '', '', '', '1', '', '', '', '', '', ''],
    ['End', '2nd', 'L3', '', '', '', '1', '', '', '', '', '', '', ''],
    ['End', '3rd', 'L3', '', '', '', '1', '', '', '', '', '', '', ''],
    ['End', '4th', 'L2', '', '1', '', '', '', '', '', '', '', '', ''],
    ['End', '5-8th', 'L1', '', '1', '', '', '', '', '', '', '', '', '']
  ];

  sheet.getRange(21, 1, templateData.length, templateHeaders.length).setValues(templateData);

  // ===== SECTION 3: Validation Panel (Rows 32-40) =====
  sheet.getRange('A32').setValue('BUDGET VALIDATION (Auto-computed during Preview)');
  sheet.getRange('A32:E32').merge().setFontWeight('bold').setFontSize(11)
    .setBackground('#fbbc04').setHorizontalAlignment('center');

  const validationLabels = [
    ['Metric', 'Value'],
    ['Expected_COGS', '(Run Preview)'],
    ['RL_Budget', '(Run Preview)'],
    ['RL_Spend_Percent', '(Run Preview)'],
    ['Status', '(Run Preview)'],
    ['Last_Preview_Event', '(Run Preview)'],
    ['Last_Preview_Time', '(Run Preview)']
  ];

  sheet.getRange(33, 1, validationLabels.length, 2).setValues(validationLabels);
  sheet.getRange('A33:B33').setFontWeight('bold').setBackground('#d9d9d9');

  // ===== SECTION 4: Instructions (Rows 42+) =====
  sheet.getRange('A42').setValue('INSTRUCTIONS');
  sheet.getRange('A42:E42').merge().setFontWeight('bold').setFontSize(11)
    .setBackground('#6aa84f').setFontColor('#ffffff').setHorizontalAlignment('center');

  const instructions = [
    [''],
    ['1. Modify the Commander Prize Template above to control what prizes are awarded in each round.'],
    ['2. Use Events → Preview Commander Prizes to see expected costs before committing.'],
    ['3. If RL Budget is exceeded, adjust the template (reduce prize levels or quantities).'],
    ['4. Run Events → Commit Commander Prizes to finalize and write prizes to the event sheet.'],
    ['5. The validation panel will update automatically during preview operations.']
  ];

  sheet.getRange(43, 1, instructions.length, 5).setValues(instructions.map(i => [i[0], '', '', '', '']));

  // Column widths
  sheet.setColumnWidth(1, 200);  // Parameter/Round
  sheet.setColumnWidth(2, 120);  // Value/Seat
  sheet.setColumnWidths(3, 12, 70);  // Level columns

  logIntegrityAction('THROTTLE_BUILD', {
    details: 'Built Prize_Throttle with Commander template',
    status: 'SUCCESS'
  });

  return sheet;
}

/**
 * Gets numeric throttle param
 * @param {string} key - Parameter key
 * @param {number} defaultValue - Default value
 * @return {number} Numeric value
 */
function getThrottleNumber(key, defaultValue) {
  const val = getThrottleParam(key, defaultValue);
  return coerceNumber(val, defaultValue);
}
/**
 * Gets boolean throttle param
 * @param {string} key - Parameter key
 * @param {boolean} defaultValue - Default value
 * @return {boolean} Boolean value
 */
function getThrottleBoolean(key, defaultValue) {
  const val = getThrottleParam(key, defaultValue);
  return coerceBoolean(val);
}