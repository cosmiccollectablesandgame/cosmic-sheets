/**
 * Commander Round Prize Service
 * @fileoverview Main entry for generating Commander round prizes
 */

/**
 * Commander Round Prize Service Class
 * Orchestrates prize generation for Commander events
 */
class CommanderRoundPrizeService {

  constructor() {
    this.throttle = new PrizeThrottleService();
    this.bathLog = new BathLogService();
    this.rarityService = new RaritySelectionService();
  }

  /**
   * Main entry: Generate for active event
   */
  generateForActiveEvent() {

    try {
      // Get active event sheet
      const eventSheet = this.getActiveEventSheet();
      if (!eventSheet) {
        throw new Error('No active event sheet found. Please select an event sheet first.');
      }

      // Load metadata
      const eventData = this.loadEventMetadata(eventSheet);

      // Check floors
      if (!this.checkFloors(eventData)) {
        return;
      }

      // Load template
      const template = this.throttle.getCommanderRoundTemplate();

      // Generate preview
      const preview = this.generatePreview(template, eventData);

      // Check RL and trim if needed
      const finalData = this.checkRLAndTrim(preview, eventData);

      // Log bath event
      if (this.throttle.isBathLoggingEnabled()) {
        this.bathLog.logBathEvent(eventData, preview, finalData);
      }

      // Show preview dialog
      this.showPreviewDialog(finalData, eventData);

    } catch (error) {
      SpreadsheetApp.getUi().alert('Error: ' + error.message);
      Logger.log('Error in generateForActiveEvent: ' + error.message);
    }
  }

  /**
   * Generates prize preview from template
   * @param {Array<Array>} template - Template grid (4x3)
   * @param {Object} eventData - Event metadata
   * @return {Object} Preview data
   */
  generatePreview(template, eventData) {

    const prizeGrid = [];
    const usedPrizes = [];
    let totalCOGS = 0;
    const df_tags = ['DF-010', 'DF-020', 'DF-260'];

    for (let rank = 0; rank < 4; rank++) {
      const rowPrizes = [];
      for (let round = 0; round < 3; round++) {
        const level = template[rank][round];
        const seedStr = `${eventData.seed}-R${rank+1}C${round+1}`;

        const prize = this.rarityService.selectPrizeFromLevel(
          level,
          seedStr,
          this.throttle.allowDuplicateItems() ? [] : usedPrizes
        );

        if (prize) {
          rowPrizes.push(prize.item_code);
          if (!this.throttle.allowDuplicateItems()) {
            usedPrizes.push(prize.item_code);
          }
          totalCOGS += prize.cogs;
        } else {
          rowPrizes.push('ERROR');
        }
      }
      prizeGrid.push(rowPrizes);
    }

    const preview_hash = this.calculateHash(prizeGrid, eventData.seed);

    return {
      prizeGrid: prizeGrid,
      totalCOGS: totalCOGS,
      df_tags: df_tags,
      preview_hash: preview_hash
    };
  }

  /**
   * Checks RL compliance and applies auto-trim if needed
   * @param {Object} preview - Preview data
   * @param {Object} eventData - Event metadata
   * @return {Object} Final data with trim status
   */
  checkRLAndTrim(preview, eventData) {

    const eligibleNet = eventData.player_count * eventData.entry_fee;
    const RL_Baseline = eligibleNet * 0.95;
    const RL_Usage = preview.totalCOGS / RL_Baseline;

    let was_trimmed = false;
    let trim_amount = 0;
    let final_cogs = preview.totalCOGS;
    let rl_band = 'GREEN';

    if (RL_Usage > 0.95) {
      rl_band = 'RED';
    } else if (RL_Usage > 0.90) {
      rl_band = 'AMBER';
    }

    // Auto-trim if Red
    if (rl_band === 'RED' && this.throttle.isAutoTrimEnabled()) {
      const target = this.throttle.getAutoTrimTarget();
      const targetCOGS = RL_Baseline * target;
      trim_amount = preview.totalCOGS - targetCOGS;
      final_cogs = targetCOGS;
      was_trimmed = true;
      rl_band = 'GREEN';

      preview.df_tags.push('DF-240');

      Logger.log(`Auto-trimmed from $${preview.totalCOGS.toFixed(2)} to $${final_cogs.toFixed(2)}`);
    }

    return {
      ...preview,
      was_trimmed: was_trimmed,
      trim_amount: trim_amount,
      final_cogs: final_cogs,
      final_rl_usage: final_cogs / RL_Baseline,
      rl_band: rl_band,
      RL_Baseline: RL_Baseline
    };
  }

  /**
   * Shows preview dialog to user
   * @param {Object} finalData - Final prize data
   * @param {Object} eventData - Event metadata
   */
  showPreviewDialog(finalData, eventData) {
    const ui = SpreadsheetApp.getUi();

    let message = `Commander Round Prize Preview\n\n`;
    message += `Event: ${eventData.event_id}\n`;
    message += `Players: ${eventData.player_count}\n\n`;
    message += `Total COGS: $${finalData.final_cogs.toFixed(2)}\n`;
    message += `RL Usage: ${(finalData.final_rl_usage * 100).toFixed(1)}%\n`;
    message += `RL Band: ${finalData.rl_band}\n\n`;

    if (finalData.was_trimmed) {
      message += `⚠ Auto-trimmed from $${(finalData.final_cogs + finalData.trim_amount).toFixed(2)}\n\n`;
    }

    message += `Proceed with commit?`;

    const result = ui.alert('Prize Preview', message, ui.ButtonSet.YES_NO);

    if (result === ui.Button.YES) {
      this.commitPrizes(finalData, eventData);
    }
  }

  /**
   * Commits prizes to event sheet
   * @param {Object} finalData - Final prize data
   * @param {Object} eventData - Event metadata
   */
  commitPrizes(finalData, eventData) {
    const eventSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(eventData.event_id);

    if (!eventSheet) {
      throw new Error('Event sheet not found: ' + eventData.event_id);
    }

    // Write to C2:E5 (Rounds 1-3, Ranks 1-4)
    eventSheet.getRange(2, 3, 4, 3).setValues(finalData.prizeGrid);

    Logger.log(`Committed prizes to ${eventData.event_id}`);

    // Log to Integrity_Log
    logIntegrityAction('COMMANDER_PRIZES_COMMIT', {
      event_id: eventData.event_id,
      total_cogs: finalData.final_cogs.toFixed(2),
      rl_usage: (finalData.final_rl_usage * 100).toFixed(1) + '%',
      rl_band: finalData.rl_band,
      was_trimmed: finalData.was_trimmed,
      seed: eventData.seed,
      hash: finalData.preview_hash
    });

    SpreadsheetApp.getUi().alert('Success', 'Commander round prizes committed!', SpreadsheetApp.getUi().ButtonSet.OK);
  }

  // ============================================================================
  // HELPER METHODS
  // ============================================================================

  /**
   * Gets active event sheet
   * @return {Sheet|null} Event sheet or null
   */
  getActiveEventSheet() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const name = sheet.getName();

    // Check if matches event pattern MM-DD-YYYY
    if (/^\d{2}-\d{2}-\d{4}/.test(name)) {
      return sheet;
    }

    return null;
  }

  /**
   * Loads event metadata from sheet
   * @param {Sheet} sheet - Event sheet
   * @return {Object} Event metadata
   */
  loadEventMetadata(sheet) {
    // Simple metadata extraction
    return {
      event_id: sheet.getName(),
      event_date: new Date(),
      format: 'Commander',
      player_count: this.getPlayerCount(sheet),
      entry_fee: 5.00, // Default, could read from sheet
      seed: this.generateSeed(sheet.getName())
    };
  }

  /**
   * Gets player count from event sheet
   * @param {Sheet} sheet - Event sheet
   * @return {number} Player count
   */
  getPlayerCount(sheet) {
    // Count non-empty rows in column B (preferred_name_id)
    const data = sheet.getRange('B2:B100').getValues();
    let count = 0;
    for (let row of data) {
      if (row[0] !== '') count++;
    }
    return count;
  }

  /**
   * Generates seed for event
   * @param {string} eventId - Event ID
   * @return {string} Seed string
   */
  generateSeed(eventId) {
    return `CMD-${eventId}-${new Date().getTime().toString(36)}`;
  }

  /**
   * Checks if floors are met
   * @param {Object} eventData - Event metadata
   * @return {boolean} True if can proceed
   */
  checkFloors(eventData) {
    const floors = this.throttle.getFloorsAndThresholds();

    if (eventData.player_count < floors.commanderFire) {
      SpreadsheetApp.getUi().alert(
        'Floor Not Met',
        `Commander requires ${floors.commanderFire}+ players to fire. Current: ${eventData.player_count}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }

    if (eventData.player_count < floors.commanderFull) {
      const result = SpreadsheetApp.getUi().alert(
        'Warning',
        `Full Commander program requires ${floors.commanderFull}+ players. Current: ${eventData.player_count}\n\nProceed anyway?`,
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      return result === SpreadsheetApp.getUi().Button.YES;
    }

    return true;
  }

  /**
   * Calculates hash for prize grid
   * @param {Array<Array>} prizeGrid - Prize grid
   * @param {string} seed - Seed string
   * @return {string} Hash
   */
  calculateHash(prizeGrid, seed) {
    const content = JSON.stringify(prizeGrid) + seed;
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, content)
      .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'))
      .join('');
  }
}

/**
 * Global function for menu handler
 * Generates Commander round prizes for active event
 */
function generateCommanderRoundPrizes() {
  const service = new CommanderRoundPrizeService();
  service.generateForActiveEvent();
}