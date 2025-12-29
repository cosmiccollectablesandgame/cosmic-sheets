/**
 * Bath_Log Service - RLbath Event Logging
 * @fileoverview Logs budget near-misses and RL violations to Bath_Log sheet
 */

/**
 * Bath Log Service Class
 * Handles logging of RLbath events
 */
class BathLogService {

  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.bathLogSheet = this.ss.getSheetByName('Bath_Log');

    if (!this.bathLogSheet) {
      this.createBathLogSheet();
    }
  }

  /**
   * Creates Bath_Log sheet with headers
   * @return {Sheet} Created sheet
   */
  createBathLogSheet() {
    this.bathLogSheet = this.ss.insertSheet('Bath_Log');

    const headers = [
      'Timestamp', 'Event_ID', 'Event_Date', 'Format', 'Player_Count', 'Entry_Fee',
      'RL_Baseline_95', 'RL_Dial_%', 'RL_Dial_$', 'Preview_Prize_COGS', 'RL_Usage_%',
      'RLbath_$', 'RLbath_%', 'Was_Trimmed', 'Trim_Amount_$', 'Final_Prize_COGS',
      'RL_Final_%', 'RL_Band_Final', 'df_tags', 'Seed', 'Preview_Hash', 'Commit_Hash', 'Notes'
    ];

    this.bathLogSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    this.bathLogSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    this.bathLogSheet.setFrozenRows(1);

    // Format header
    this.bathLogSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285f4')
      .setFontColor('#ffffff');

    Logger.log('Bath_Log sheet created');

    return this.bathLogSheet;
  }

  /**
   * Logs a Bath event
   * @param {Object} eventData - Event metadata
   * @param {Object} previewData - Preview prize data
   * @param {Object} finalData - Final committed data
   */
  logBathEvent(eventData, previewData, finalData) {

    const throttle = new PrizeThrottleService();

    const eligibleNet = eventData.player_count * eventData.entry_fee;
    const RL_Baseline = eligibleNet * 0.95;
    const RL_Dial_Pct = throttle.getRLDial();
    const RL_Dial_Dollar = eligibleNet * RL_Dial_Pct;

    const RL_Usage_Pct = previewData.totalCOGS / RL_Baseline;
    const RLbath_Dollar = Math.max(0, previewData.totalCOGS - RL_Baseline);
    const RLbath_Pct = Math.max(0, RL_Usage_Pct - 1.0);

    // Log if bath territory OR dial above baseline OR trimmed
    if (RLbath_Dollar > 0 || RL_Dial_Pct > 0.95 || finalData.was_trimmed) {

      const row = [
        new Date(),
        eventData.event_id,
        eventData.event_date,
        eventData.format || 'Commander',
        eventData.player_count,
        eventData.entry_fee,
        RL_Baseline,
        RL_Dial_Pct,
        RL_Dial_Dollar,
        previewData.totalCOGS,
        RL_Usage_Pct,
        RLbath_Dollar,
        RLbath_Pct,
        finalData.was_trimmed || false,
        finalData.trim_amount || 0,
        finalData.final_cogs,
        finalData.final_rl_usage,
        finalData.rl_band,
        Array.isArray(finalData.df_tags) ? finalData.df_tags.join(',') : '',
        eventData.seed || '',
        previewData.preview_hash || '',
        finalData.commit_hash || '',
        finalData.notes || ''
      ];

      this.bathLogSheet.appendRow(row);

      Logger.log(`Bath event logged: RLbath $${RLbath_Dollar.toFixed(2)} (${(RLbath_Pct * 100).toFixed(1)}%)`);

      // Also log to Integrity_Log
      logIntegrityAction('BATH_EVENT', {
        event_id: eventData.event_id,
        rl_usage: (RL_Usage_Pct * 100).toFixed(1) + '%',
        was_trimmed: finalData.was_trimmed,
        rl_band: finalData.rl_band
      });
    }
  }

  /**
   * Gets Bath_Log entries
   * @param {number} limit - Max entries to return
   * @return {Array<Object>} Log entries
   */
  getBathLog(limit = 100) {
    if (!this.bathLogSheet) return [];

    const data = this.bathLogSheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const entries = toObjects(data);

    // Sort by timestamp descending
    entries.sort((a, b) => new Date(b.Timestamp) - new Date(a.Timestamp));

    return entries.slice(0, limit);
  }

  /**
   * Gets Bath_Log statistics
   * @return {Object} Statistics
   */
  getBathStats() {
    const entries = this.getBathLog();

    if (entries.length === 0) {
      return {
        total_events: 0,
        trimmed_events: 0,
        avg_rl_usage: 0,
        max_rl_usage: 0
      };
    }

    const trimmedCount = entries.filter(e => e.Was_Trimmed === true).length;
    const avgUsage = entries.reduce((sum, e) => sum + coerceNumber(e['RL_Usage_%'], 0), 0) / entries.length;
    const maxUsage = Math.max(...entries.map(e => coerceNumber(e['RL_Usage_%'], 0)));

    return {
      total_events: entries.length,
      trimmed_events: trimmedCount,
      avg_rl_usage: avgUsage,
      max_rl_usage: maxUsage
    };
  }
}