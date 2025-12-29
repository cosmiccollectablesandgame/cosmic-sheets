/**
 * Rarity-Weighted Prize Selection Service
 * @fileoverview Selects prizes from catalog using rarity weighting and deterministic seeding
 */

/**
 * Rarity Selection Service Class
 * Handles prize selection with rarity-based weighting
 */
class RaritySelectionService {

  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.catalogSheet = this.ss.getSheetByName('Prize_Catalog');
    this.throttle = new PrizeThrottleService();

    if (!this.catalogSheet) {
      throw new Error('Prize_Catalog sheet not found');
    }
  }

  /**
   * Selects a prize from Prize_Catalog at specified level
   * Uses rarity weighting + seed for deterministic selection
   * @param {string} level - Prize level (L0-L13)
   * @param {string} seed - Seed string for determinism
   * @param {Array<string>} excludeList - Item codes to exclude
   * @return {Object|null} Selected prize object or null
   */
  selectPrizeFromLevel(level, seed, excludeList = []) {

    // Get eligible items from catalog
    const eligible = this.getEligibleItems(level, excludeList);

    if (eligible.length === 0) {
      Logger.log(`ERROR: No eligible items for level ${level}`);
      return null;
    }

    // Build weighted pool
    const weightedPool = this.buildWeightedPool(eligible);

    if (weightedPool.length === 0) {
      Logger.log(`ERROR: Weighted pool empty for level ${level}`);
      return null;
    }

    // Deterministic selection using seed
    const index = this.seededRandom(seed, weightedPool.length);
    const selected = weightedPool[index];

    Logger.log(`Selected ${selected.item_code} (Rarity ${selected.rarity}) from level ${level} using seed ${seed}`);

    return selected;
  }

  /**
   * Gets eligible items from Prize_Catalog
   * @param {string} level - Prize level
   * @param {Array<string>} excludeList - Codes to exclude
   * @return {Array<Object>} Eligible items
   */
  getEligibleItems(level, excludeList) {
    const data = this.catalogSheet.getDataRange().getValues();
    const headers = data[0];

    const levelCol = headers.indexOf('Level');
    const itemCodeCol = headers.indexOf('Code');
    const itemNameCol = headers.indexOf('Name');
    const rarityCol = headers.indexOf('Rarity');
    const qtyCol = headers.indexOf('Qty');
    const cogsCol = headers.indexOf('COGS');
    const eligibleRoundsCol = headers.indexOf('Eligible_Rounds');

    if (levelCol === -1 || itemCodeCol === -1) {
      throw new Error('Required columns not found in Prize_Catalog');
    }

    const eligible = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Check criteria
      const matchesLevel = row[levelCol] === level;
      const hasRoundEligibility = eligibleRoundsCol === -1 || coerceBoolean(row[eligibleRoundsCol]);
      const hasQty = qtyCol === -1 || coerceNumber(row[qtyCol], 0) > 0;
      const notExcluded = !excludeList.includes(row[itemCodeCol]);

      if (matchesLevel && hasRoundEligibility && hasQty && notExcluded) {
        eligible.push({
          item_code: row[itemCodeCol],
          item_name: itemNameCol !== -1 ? row[itemNameCol] : '',
          rarity: rarityCol !== -1 ? coerceNumber(row[rarityCol], 5) : 5,
          qty: qtyCol !== -1 ? coerceNumber(row[qtyCol], 1) : 1,
          cogs: cogsCol !== -1 ? coerceNumber(row[cogsCol], 0) : 0
        });
      }
    }

    return eligible;
  }

  /**
   * Builds weighted pool based on rarity
   * @param {Array<Object>} eligible - Eligible items
   * @return {Array<Object>} Weighted pool
   */
  buildWeightedPool(eligible) {
    const pool = [];

    if (!this.throttle.isRarityWeightingActive()) {
      // No weighting - each item appears once
      return eligible;
    }

    for (const item of eligible) {
      const weight = this.throttle.getRarityWeight(item.rarity);
      const tickets = Math.max(1, Math.round(weight));

      for (let i = 0; i < tickets; i++) {
        pool.push(item);
      }
    }

    return pool;
  }

  /**
   * Seeded random number generator
   * Returns index in range [0, max)
   * @param {string} seed - Seed string
   * @param {number} max - Max value (exclusive)
   * @return {number} Index
   */
  seededRandom(seed, max) {
    // Simple hash function
    let hash = 0;
    for (let i = 0; i < seed.length; i++) {
      const char = seed.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32-bit integer
    }

    // Convert to positive and mod by max
    const positive = Math.abs(hash);
    return positive % max;
  }

  /**
   * Gets average COGS for a level (for validation formulas)
   * @param {string} level - Prize level
   * @return {number} Average COGS
   */
  getLevelAverageCOGS(level) {
    const eligible = this.getEligibleItems(level, []);

    if (eligible.length === 0) return 0;

    if (!this.throttle.isRarityWeightingActive()) {
      // Simple average
      const total = eligible.reduce((sum, item) => sum + item.cogs, 0);
      return total / eligible.length;
    }

    // Weighted average
    let totalWeightedCOGS = 0;
    let totalWeight = 0;

    for (const item of eligible) {
      const weight = this.throttle.getRarityWeight(item.rarity);
      totalWeightedCOGS += item.cogs * weight;
      totalWeight += weight;
    }

    return totalWeight > 0 ? totalWeightedCOGS / totalWeight : 0;
  }
}