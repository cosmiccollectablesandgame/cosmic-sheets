/**
 * Custom Functions for Prize_Throttle
 * @fileoverview Sheet formula functions for real-time validation
 */

/**
 * Gets average COGS for a prize level using rarity weighting
 * Used in Prize_Throttle validation formulas
 *
 * @param {string} level - Prize level (L0-L13)
 * @return {number} Average COGS for the level
 * @customfunction
 */
function GET_LEVEL_AVG_COGS(level) {
  if (!level || level === '') return 0;

  try {
    const rarityService = new RaritySelectionService();
    return rarityService.getLevelAverageCOGS(level);
  } catch (e) {
    Logger.log('Error in GET_LEVEL_AVG_COGS: ' + e.message);
    return 0;
  }
}

/**
 * Validates if a prize level exists in the catalog
 *
 * @param {string} level - Prize level (L0-L13)
 * @return {boolean} True if level has items
 * @customfunction
 */
function VALIDATE_PRIZE_LEVEL(level) {
  if (!level || level === '') return false;

  try {
    const rarityService = new RaritySelectionService();
    const eligible = rarityService.getEligibleItems(level, []);
    return eligible.length > 0;
  } catch (e) {
    return false;
  }
}

/**
 * Gets count of items at a prize level
 *
 * @param {string} level - Prize level (L0-L13)
 * @return {number} Count of items
 * @customfunction
 */
function COUNT_LEVEL_ITEMS(level) {
  if (!level || level === '') return 0;

  try {
    const rarityService = new RaritySelectionService();
    const eligible = rarityService.getEligibleItems(level, []);
    return eligible.length;
  } catch (e) {
    return 0;
  }
}

/**
 * Gets total inventory quantity for a level
 *
 * @param {string} level - Prize level (L0-L13)
 * @return {number} Total quantity
 * @customfunction
 */
function GET_LEVEL_TOTAL_QTY(level) {
  if (!level || level === '') return 0;

  try {
    const rarityService = new RaritySelectionService();
    const eligible = rarityService.getEligibleItems(level, []);
    return eligible.reduce((sum, item) => sum + item.qty, 0);
  } catch (e) {
    return 0;
  }
}

/**
 * Calculates RL band from usage percentage
 *
 * @param {number} rlUsagePct - RL usage as percentage (e.g., 0.93 for 93%)
 * @return {string} Band: GREEN, AMBER, or RED
 * @customfunction
 */
function GET_RL_BAND(rlUsagePct) {
  if (typeof rlUsagePct !== 'number') return 'ERROR';

  if (rlUsagePct <= 0.90) return 'GREEN';
  if (rlUsagePct <= 0.95) return 'AMBER';
  return 'RED';
}

/**
 * Formats RL band with color indicator
 *
 * @param {string} band - Band name (GREEN/AMBER/RED)
 * @return {string} Formatted band
 * @customfunction
 */
function FORMAT_RL_BAND(band) {
  switch (band) {
    case 'GREEN': return '🟢 GREEN';
    case 'AMBER': return '🟡 AMBER';
    case 'RED': return '🔴 RED';
    default: return band;
  }
}