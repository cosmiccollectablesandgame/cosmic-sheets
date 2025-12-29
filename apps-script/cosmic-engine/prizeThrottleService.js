/**
 * Prize_Throttle Service - Commander Prize Configuration
 * @fileoverview Handles reading Prize_Throttle controls for Commander round prizes
 * Note: This is separate from the simple KV Prize_Throttle used by other systems
 */

/**
 * Prize Throttle Service Class
 * Reads configuration from Prize_Throttle sheet
 */
class PrizeThrottleService {

  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.throttleSheet = this.ss.getSheetByName('Prize_Throttle');

    if (!this.throttleSheet) {
      throw new Error('Prize_Throttle sheet not found. Run Build/Repair first.');
    }
  }

  /**
   * Gets RL Baseline (DF-010: immutable at 95%)
   * @return {number} 0.95
   */
  getRLBaseline() {
    return 0.95; // Immutable DF-010
  }

  /**
   * Gets RL Dial percentage
   * @return {number} RL dial value
   */
  getRLDial() {
    try {
      const value = this.throttleSheet.getRange('RL_Dial_Pct').getValue();
      return coerceNumber(value, 0.95);
    } catch (e) {
      Logger.log('Warning: RL_Dial_Pct not found, using default 0.95');
      return 0.95;
    }
  }

  /**
   * Gets active RL cap (min of baseline and dial)
   * @return {number} Active RL cap
   */
  getActiveRLCap() {
    return Math.min(this.getRLBaseline(), this.getRLDial());
  }

  /**
   * Gets Commander Round Template (B27:D30)
   * @return {Array<Array>} 4x3 array of prize levels
   */
  getCommanderRoundTemplate() {
    try {
      return this.throttleSheet.getRange('Commander_Round_Template').getValues();
    } catch (e) {
      throw new Error('Commander_Round_Template not found in Prize_Throttle');
    }
  }

  /**
   * Gets test scenario settings
   * @return {Object} {playerCount, entryFee}
   */
  getTestScenario() {
    try {
      return {
        playerCount: coerceNumber(this.throttleSheet.getRange('Test_Player_Count').getValue(), 8),
        entryFee: coerceNumber(this.throttleSheet.getRange('Test_Entry_Fee').getValue(), 5.00)
      };
    } catch (e) {
      return { playerCount: 8, entryFee: 5.00 };
    }
  }

  /**
   * Checks if auto-trim is enabled
   * @return {boolean} Auto-trim enabled
   */
  isAutoTrimEnabled() {
    try {
      return coerceBoolean(this.throttleSheet.getRange('Auto_Trim_Enabled').getValue());
    } catch (e) {
      return true;
    }
  }

  /**
   * Gets auto-trim target percentage
   * @return {number} Target percentage
   */
  getAutoTrimTarget() {
    try {
      return coerceNumber(this.throttleSheet.getRange('Auto_Trim_Target').getValue(), 0.90);
    } catch (e) {
      return 0.90;
    }
  }

  /**
   * Checks if Bath logging is enabled
   * @return {boolean} Bath logging enabled
   */
  isBathLoggingEnabled() {
    try {
      return coerceBoolean(this.throttleSheet.getRange('Log_Bath_Events').getValue());
    } catch (e) {
      return true;
    }
  }

  /**
   * Gets rarity weight for a rarity value
   * @param {number} rarity - Rarity (0-10)
   * @return {number} Weight multiplier
   */
  getRarityWeight(rarity) {
    try {
      if (rarity <= 1) return coerceNumber(this.throttleSheet.getRange('Rarity_Weight_01').getValue(), 0.5);
      if (rarity <= 3) return coerceNumber(this.throttleSheet.getRange('Rarity_Weight_23').getValue(), 1.0);
      if (rarity <= 6) return coerceNumber(this.throttleSheet.getRange('Rarity_Weight_46').getValue(), 2.0);
      if (rarity <= 9) return coerceNumber(this.throttleSheet.getRange('Rarity_Weight_79').getValue(), 3.5);
      return coerceNumber(this.throttleSheet.getRange('Rarity_Weight_10').getValue(), 5.0);
    } catch (e) {
      // Fallback to defaults
      if (rarity <= 1) return 0.5;
      if (rarity <= 3) return 1.0;
      if (rarity <= 6) return 2.0;
      if (rarity <= 9) return 3.5;
      return 5.0;
    }
  }

  /**
   * Checks if rarity weighting is active
   * @return {boolean} Rarity weighting active
   */
  isRarityWeightingActive() {
    try {
      return coerceBoolean(this.throttleSheet.getRange('Rarity_Weighting_Active').getValue());
    } catch (e) {
      return true;
    }
  }

  /**
   * Gets floors and thresholds
   * @return {Object} Floor values
   */
  getFloorsAndThresholds() {
    const getNamedValue = (name, defaultVal) => {
      try {
        return coerceNumber(this.throttleSheet.getRange(name).getValue(), defaultVal);
      } catch (e) {
        return defaultVal;
      }
    };

    return {
      commanderFire: getNamedValue('Floor_Commander_Fire', 2),
      commanderFull: getNamedValue('Floor_Commander_Full', 8),
      l3: getNamedValue('Threshold_L3', 8),
      promoReg: getNamedValue('Threshold_Promo_Reg', 8),
      promoFoil: getNamedValue('Threshold_Promo_Foil', 12),
      l4: getNamedValue('Threshold_L4', 12),
      merch1: getNamedValue('Threshold_Merch_1', 20),
      sealed24: getNamedValue('Threshold_Sealed_24', 24),
      bigBump: getNamedValue('Threshold_Big_Bump', 28),
      merch2: getNamedValue('Threshold_Merch_2', 32)
    };
  }

  /**
   * Checks if duplicate items allowed
   * @return {boolean} Allow duplicates
   */
  allowDuplicateItems() {
    try {
      return coerceBoolean(this.throttleSheet.getRange('Allow_Duplicate_Items').getValue());
    } catch (e) {
      return false;
    }
  }

  /**
   * Checks if should prefer in-stock items
   * @return {boolean} Prefer in-stock
   */
  preferInStock() {
    try {
      return coerceBoolean(this.throttleSheet.getRange('Prefer_InStock').getValue());
    } catch (e) {
      return true;
    }
  }
}