/**
 * suffixTests.gs
 * TEST & SANITY CHECKS FOR SUFFIX SYSTEM v7.9.7+
 *
 * Basic tests and sanity checks for suffix parsing and mission triggers.
 * Run testSuffixSystem() to verify all components are working correctly.
 *
 * @fileoverview Test utilities for suffix configuration and mission system
 */

/**
 * Master test runner for suffix system.
 * Run this function to verify all suffix components are working.
 *
 * @return {Object} Test results summary
 */
function testSuffixSystem() {
  Logger.log('========================================');
  Logger.log('SUFFIX SYSTEM TEST SUITE v7.9.7+');
  Logger.log('========================================\n');

  const results = {
    total: 0,
    passed: 0,
    failed: 0,
    errors: []
  };

  // Run all test suites
  runTestSuite_('Suffix Parsing', testSuffixParsing_, results);
  runTestSuite_('Suffix Metadata', testSuffixMetadata_, results);
  runTestSuite_('Mission Triggers', testMissionTriggers_, results);
  runTestSuite_('Commander Brackets', testCommanderBrackets_, results);
  runTestSuite_('Limited Formats', testLimitedFormats_, results);

  // Summary
  Logger.log('\n========================================');
  Logger.log('TEST SUMMARY');
  Logger.log('========================================');
  Logger.log(`Total Tests: ${results.total}`);
  Logger.log(`Passed: ${results.passed}`);
  Logger.log(`Failed: ${results.failed}`);

  if (results.errors.length > 0) {
    Logger.log('\nFailed Tests:');
    results.errors.forEach(err => Logger.log(`  ✗ ${err}`));
  }

  Logger.log('========================================\n');

  return results;
}

/**
 * Test suite runner helper
 * @private
 */
function runTestSuite_(name, testFn, results) {
  Logger.log(`\n--- ${name} ---`);
  try {
    testFn(results);
  } catch (e) {
    results.failed++;
    results.total++;
    results.errors.push(`${name}: ${e.message}`);
    Logger.log(`✗ SUITE ERROR: ${e.message}`);
  }
}

/**
 * Test assertion helper
 * @private
 */
function assert_(condition, message, results) {
  results.total++;
  if (condition) {
    results.passed++;
    Logger.log(`✓ ${message}`);
  } else {
    results.failed++;
    results.errors.push(message);
    Logger.log(`✗ ${message}`);
  }
}

// ============================================================================
// TEST SUITES
// ============================================================================

/**
 * Tests suffix parsing from event IDs
 * @private
 */
function testSuffixParsing_(results) {
  // Test no suffix
  let suffix = getSuffixFromEventId_('11-23-2025');
  assert_(suffix === null, 'Parse "11-23-2025" → null (no suffix)', results);

  // Test single letter suffixes
  suffix = getSuffixFromEventId_('11-23B-2025');
  assert_(suffix === 'B', 'Parse "11-23B-2025" → "B"', results);

  suffix = getSuffixFromEventId_('05-01T-2026');
  assert_(suffix === 'T', 'Parse "05-01T-2026" → "T"', results);

  suffix = getSuffixFromEventId_('12-25A-2025');
  assert_(suffix === 'A', 'Parse "12-25A-2025" → "A"', results);

  suffix = getSuffixFromEventId_('03-15D-2025');
  assert_(suffix === 'D', 'Parse "03-15D-2025" → "D"', results);

  // Test invalid formats
  suffix = getSuffixFromEventId_('invalid');
  assert_(suffix === null, 'Parse "invalid" → null', results);

  suffix = getSuffixFromEventId_(null);
  assert_(suffix === null, 'Parse null → null', results);

  suffix = getSuffixFromEventId_('');
  assert_(suffix === null, 'Parse empty string → null', results);
}

/**
 * Tests suffix metadata retrieval
 * @private
 */
function testSuffixMetadata_(results) {
  // Test valid suffix codes
  let meta = getSuffixMeta_('B');
  assert_(meta !== null, 'Get metadata for "B" (Casual Commander)', results);
  assert_(meta.code === 'B', 'B.code === "B"', results);
  assert_(meta.name === 'Casual Commander (Brk 1–2)', 'B.name correct', results);
  assert_(meta.commanderBracket === 2, 'B.commanderBracket === 2', results);
  assert_(meta.requiresKitPrompt === false, 'B.requiresKitPrompt === false', results);

  meta = getSuffixMeta_('T');
  assert_(meta !== null, 'Get metadata for "T" (cEDH)', results);
  assert_(meta.commanderBracket === 5, 'T.commanderBracket === 5', results);

  meta = getSuffixMeta_('D');
  assert_(meta !== null, 'Get metadata for "D" (Booster Draft)', results);
  assert_(meta.requiresKitPrompt === true, 'D.requiresKitPrompt === true', results);

  // Test invalid suffix
  meta = getSuffixMeta_('999');
  assert_(meta === null, 'Get metadata for invalid suffix → null', results);

  meta = getSuffixMeta_(null);
  assert_(meta === null, 'Get metadata for null → null', results);
}

/**
 * Tests mission trigger mapping
 * @private
 */
function testMissionTriggers_(results) {
  // Test that all 26 suffixes have valid metadata
  const allCodes = getAllSuffixCodes_();
  assert_(allCodes.length === 26, 'SUFFIX_MAP contains 26 entries (A-Z)', results);

  // Test suffix validation
  assert_(isValidSuffix_('B') === true, 'isValidSuffix_("B") === true', results);
  assert_(isValidSuffix_('T') === true, 'isValidSuffix_("T") === true', results);
  assert_(isValidSuffix_('Z') === true, 'isValidSuffix_("Z") === true', results);
  assert_(isValidSuffix_('999') === false, 'isValidSuffix_("999") === false', results);
  assert_(isValidSuffix_(null) === false, 'isValidSuffix_(null) === false', results);

  // Test display name generation
  const displayName = getSuffixDisplayName_('B');
  assert_(displayName === 'B – Casual Commander (Brk 1–2)', 'Display name for B correct', results);
}

/**
 * Tests Commander bracket mappings
 * @private
 */
function testCommanderBrackets_(results) {
  // Test B = Brackets 1-2
  let meta = getSuffixMeta_('B');
  assert_(meta.commanderRange !== null, 'B has commanderRange', results);
  assert_(meta.commanderRange[0] === 1, 'B.commanderRange[0] === 1', results);
  assert_(meta.commanderRange[1] === 2, 'B.commanderRange[1] === 2', results);
  assert_(meta.missionTags.includes('COMMANDER'), 'B has COMMANDER tag', results);
  assert_(meta.missionTags.includes('BRK_1_2'), 'B has BRK_1_2 tag', results);

  // Test C = Brackets 3-4
  meta = getSuffixMeta_('C');
  assert_(meta.commanderRange !== null, 'C has commanderRange', results);
  assert_(meta.commanderRange[0] === 3, 'C.commanderRange[0] === 3', results);
  assert_(meta.commanderRange[1] === 4, 'C.commanderRange[1] === 4', results);
  assert_(meta.missionTags.includes('BRK_3_4'), 'C has BRK_3_4 tag', results);

  // Test T = Bracket 5 (cEDH)
  meta = getSuffixMeta_('T');
  assert_(meta.commanderRange !== null, 'T has commanderRange', results);
  assert_(meta.commanderRange[0] === 5, 'T.commanderRange[0] === 5', results);
  assert_(meta.commanderRange[1] === 5, 'T.commanderRange[1] === 5', results);
  assert_(meta.missionTags.includes('BRK_5'), 'T has BRK_5 tag', results);
  assert_(meta.missionTags.includes('CEDH'), 'T has CEDH tag', results);

  // Test L = Commander League (no specific bracket)
  meta = getSuffixMeta_('L');
  assert_(meta.commanderBracket === null, 'L.commanderBracket === null (league)', results);
  assert_(meta.missionTags.includes('LEAGUE'), 'L has LEAGUE tag', results);
}

/**
 * Tests Limited format identification
 * @private
 */
function testLimitedFormats_(results) {
  // Test D = Booster Draft (Limited, requires kit prompt)
  let meta = getSuffixMeta_('D');
  assert_(meta.formatType === 'LIMITED', 'D.formatType === "LIMITED"', results);
  assert_(meta.requiresKitPrompt === true, 'D requires kit prompt', results);
  assert_(meta.missionTags.includes('LIMITED'), 'D has LIMITED tag', results);

  // Test P = Proxy/Cube Draft (also Limited)
  meta = getSuffixMeta_('P');
  assert_(meta.formatType === 'LIMITED', 'P.formatType === "LIMITED"', results);
  assert_(meta.requiresKitPrompt === true, 'P requires kit prompt', results);

  // Test R = Prerelease Sealed
  meta = getSuffixMeta_('R');
  assert_(meta.formatType === 'LIMITED', 'R.formatType === "LIMITED"', results);
  assert_(meta.requiresKitPrompt === true, 'R requires kit prompt', results);

  // Test S = Sealed
  meta = getSuffixMeta_('S');
  assert_(meta.formatType === 'LIMITED', 'S.formatType === "LIMITED"', results);
  assert_(meta.requiresKitPrompt === true, 'S requires kit prompt', results);

  // Test that all Limited formats are captured
  const limitedFormats = getFilteredSuffixes_({ requiresKitPrompt: true });
  assert_(limitedFormats.length === 4, 'Exactly 4 formats require kit prompt (D/P/R/S)', results);

  // Test Constructed format (should NOT require kit prompt)
  meta = getSuffixMeta_('B');
  assert_(meta.formatType === 'CONSTRUCTED', 'B.formatType === "CONSTRUCTED"', results);
  assert_(meta.requiresKitPrompt === false, 'B does NOT require kit prompt', results);
}

// ============================================================================
// INTEGRATION TEST HELPERS
// ============================================================================

/**
 * Quick sanity check for suffix system (can be called from menu)
 * Logs a simple pass/fail summary to the console.
 */
function quickSuffixCheck() {
  const ui = SpreadsheetApp.getUi();

  try {
    const results = testSuffixSystem();

    const message = `Suffix System Tests\n\n` +
                    `Total: ${results.total}\n` +
                    `Passed: ${results.passed}\n` +
                    `Failed: ${results.failed}\n\n` +
                    (results.failed > 0 ? 'See Logs (Ctrl+Enter) for details.' : 'All tests passed! ✓');

    ui.alert('Suffix System Test Results', message, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('Test Error', `Failed to run suffix tests:\n\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Demonstrates suffix parsing with current event tabs
 */
function demonstrateSuffixParsing() {
  Logger.log('========================================');
  Logger.log('SUFFIX PARSING DEMONSTRATION');
  Logger.log('========================================\n');

  const eventTabs = listEventTabs();

  if (eventTabs.length === 0) {
    Logger.log('No event tabs found. Create some events with suffixes to test.');
    return;
  }

  Logger.log(`Found ${eventTabs.length} event tab(s):\n`);

  eventTabs.forEach(eventId => {
    const suffix = getSuffixFromEventId_(eventId);
    const meta = suffix ? getSuffixMeta_(suffix) : null;

    Logger.log(`Event: ${eventId}`);
    if (meta) {
      Logger.log(`  Suffix: ${suffix} – ${meta.name}`);
      Logger.log(`  Game: ${meta.game}`);
      Logger.log(`  Format: ${meta.formatType}`);
      Logger.log(`  Kit Prompt: ${meta.requiresKitPrompt ? 'YES' : 'NO'}`);
      if (meta.commanderBracket) {
        Logger.log(`  Commander Bracket: ${meta.commanderBracket}`);
      }
      Logger.log(`  Mission Tags: ${meta.missionTags.join(', ')}`);
    } else {
      Logger.log(`  Suffix: None`);
    }
    Logger.log('');
  });

  Logger.log('========================================\n');
}