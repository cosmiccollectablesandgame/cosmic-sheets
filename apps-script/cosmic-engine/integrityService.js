/**
 * Integrity Service - Logging and Audit Trail
 * @fileoverview Manages Integrity_Log and Spent_Pool writes, checksums, batch IDs
 */
// ============================================================================
// INTEGRITY LOG
// ============================================================================
/**
 * Logs an action to Integrity_Log
 * @param {string} action - Action type (e.g., 'EVENT_CREATE', 'COMMIT')
 * @param {Object} payload - Action details
 * @param {string} payload.storeId - Store ID (optional, defaults to 'MAIN')
 * @param {string} payload.eventId - Event ID
 * @param {string} payload.preferredName - Player name (optional)
 * @param {string} payload.seed - Seed used (optional)
 * @param {string} payload.checksumBefore - Checksum before (optional)
 * @param {string} payload.checksumAfter - Checksum after (optional)
 * @param {string} payload.rlBand - RL band: Green/Amber/Red (optional)
 * @param {Array<string>} payload.dfTags - Decision flags (optional)
 * @param {string} payload.details - Additional details (optional)
 * @param {string} payload.status - Status: SUCCESS/FAILURE/ABORTED (optional)
 */
function logIntegrityAction(action, payload = {}) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Integrity_Log');
    // Create sheet if missing
    if (!sheet) {
      sheet = ss.insertSheet('Integrity_Log');
      sheet.appendRow([
        'Timestamp',
        'StoreID',
        'Event_ID',
        'Action',
        'Operator',
        'PreferredName',
        'Seed',
        'Checksum_Before',
        'Checksum_After',
        'RL_Band',
        'DF_Tags',
        'Details',
        'Status'
      ]);
    }

    const row = [
      dateISO(),
      payload.storeId || 'MAIN',
      payload.eventId || '',
      action,
      currentUser(),
      payload.preferredName || '',
      payload.seed || '',
      payload.checksumBefore || '',
      payload.checksumAfter || '',
      payload.rlBand || '',
      payload.dfTags ? payload.dfTags.join(',') : '',
      payload.details || '',
      payload.status || 'SUCCESS'
    ];
    sheet.appendRow(row);
  } catch (e) {
    console.error('Failed to log integrity action:', action, e);
    // Don't throw - logging failures shouldn't block operations
  }
}
/**
 * Logs a preview action
 * @param {string} eventId - Event ID
 * @param {string} seed - Seed used
 * @param {Object} preview - Preview object
 * @param {string} rlBand - RL band
 */
function logPreview(eventId, seed, preview, rlBand) {
  const hash = computeHash(preview);
  logIntegrityAction('PREVIEW', {
    eventId,
    seed,
    checksumBefore: hash,
    rlBand,
    dfTags: ['DF-050'],
    details: `Preview generated: ${preview.allocations?.length || 0} allocations`,
    status: 'SUCCESS'
  });
}
/**
 * Logs a commit action
 * @param {string} eventId - Event ID
 * @param {string} seed - Seed used
 * @param {string} previewHash - Hash from preview
 * @param {string} commitHash - Hash after commit
 * @param {string} rlBand - RL band
 * @param {number} spentTotal - Total COGS spent
 */
function logCommit(eventId, seed, previewHash, commitHash, rlBand, spentTotal) {
  const matched = previewHash === commitHash;
  logIntegrityAction('COMMIT', {
    eventId,
    seed,
    checksumBefore: previewHash,
    checksumAfter: commitHash,
    rlBand,
    dfTags: ['DF-050', 'DF-010'],
    details: `Spent: ${formatCurrency(spentTotal)} | Hash match: ${matched}`,
    status: matched ? 'SUCCESS' : 'ABORTED'
  });
}
// ============================================================================
// SPENT POOL
// ============================================================================
/**
 * Writes entries to Spent_Pool
 * @param {Array<Object>} entries - Array of spent pool entries
 * @param {string} entries[].eventId - Event ID
 * @param {string} entries[].itemCode - Item code
 * @param {string} entries[].itemName - Item name
 * @param {string} entries[].level - Level (L0-L4)
 * @param {number} entries[].qty - Quantity
 * @param {number} entries[].cogs - COGS per unit
 * @param {string} entries[].eventType - Event type
 * @param {string} batchId - Batch ID for this write
 */
function writeSpentPool(entries, batchId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Spent_Pool');
  // Create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet('Spent_Pool');
    sheet.appendRow([
      'Event_ID',
      'Item_Code',
      'Item_Name',
      'Level',
      'Qty',
      'COGS',
      'Total',
      'Timestamp',
      'Batch_ID',
      'Reverted',
      'Event_Type'
    ]);
  }
  const timestamp = dateISO();
  const rows = entries.map(entry => [
    entry.eventId,
    entry.itemCode,
    entry.itemName,
    entry.level,
    entry.qty,
    entry.cogs,
    entry.qty * entry.cogs,
    timestamp,
    batchId,
    false, // Not reverted
    entry.eventType || 'CONSTRUCTED'
  ]);
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}
/**
 * Generates a new batch ID
 * @return {string} Batch ID (timestamp-based)
 */
function newBatchId() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss-') +
    Math.random().toString(36).substring(2, 7).toUpperCase();
}
/**
 * Gets total spent for an event from Spent_Pool
 * @param {string} eventId - Event ID
 * @return {number} Total COGS spent
 */
function getEventSpent(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Spent_Pool');
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let total = 0;
  for (let i = 1; i < data.length; i++) { // Skip header
    const [evId, , , , qty, cogs, , , , reverted] = data[i];
    if (evId === eventId && !reverted) {
      total += (qty * cogs);
    }
  }
  return total;
}
/**
 * Reverts a batch from Spent_Pool
 * @param {string} batchId - Batch ID to revert
 * @return {number} Count of reverted entries
 */
function revertBatch(batchId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Spent_Pool');
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  let revertCount = 0;
  for (let i = 1; i < data.length; i++) { // Skip header
    const batchCol = 8; // Batch_ID column (0-indexed: 8)
    const revertedCol = 9; // Reverted column
    if (data[i][batchCol] === batchId && !data[i][revertedCol]) {
      sheet.getRange(i + 1, revertedCol + 1).setValue(true);
      revertCount++;
    }
  }
  if (revertCount > 0) {
    logIntegrityAction('REVERT_BATCH', {
      details: `Reverted batch ${batchId}: ${revertCount} entries`,
      status: 'SUCCESS'
    });
  }
  return revertCount;
}
// ============================================================================
// PREVIEW ARTIFACTS
// ============================================================================
/**
 * Stores a preview artifact (for hash verification)
 * @param {string} eventId - Event ID
 * @param {string} seed - Seed used
 * @param {string} previewHash - Preview hash
 * @param {number} expiresHours - Expiration in hours (default: 24)
 * @return {string} Artifact ID
 */
function storePreviewArtifact(eventId, seed, previewHash, expiresHours = 24) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preview_Artifacts');
  // Create hidden sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet('Preview_Artifacts');
    sheet.hideSheet();
    sheet.appendRow([
      'Artifact_ID',
      'Event_ID',
      'Seed',
      'Preview_Hash',
      'Created_At',
      'Expires_At'
    ]);
  }
  const artifactId = newBatchId();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + expiresHours * 60 * 60 * 1000);
  sheet.appendRow([
    artifactId,
    eventId,
    seed,
    previewHash,
    dateISO(),
    Utilities.formatDate(expiresAt, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'")
  ]);
  return artifactId;
}
/**
 * Retrieves a preview artifact by event ID
 * @param {string} eventId - Event ID
 * @return {Object|null} Artifact or null
 */
function getPreviewArtifact(eventId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preview_Artifacts');
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  // Find most recent non-expired artifact for this event
  for (let i = data.length - 1; i > 0; i--) { // Reverse search, skip header
    const [artifactId, evId, seed, previewHash, createdAt, expiresAt] = data[i];
    if (evId === eventId) {
      const expiry = new Date(expiresAt);
      if (expiry > now) {
        return {
          artifactId,
          eventId: evId,
          seed,
          previewHash,
          createdAt,
          expiresAt
        };
      }
    }
  }
  return null;
}
/**
 * Deletes a preview artifact
 * @param {string} artifactId - Artifact ID
 */
function deletePreviewArtifact(artifactId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Preview_Artifacts');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // Skip header
    if (data[i][0] === artifactId) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}
// ============================================================================
// RL BANDS
// ============================================================================
/**
 * Computes RL band based on percentage
 * @param {number} percentUsed - Percentage of budget used (0-1)
 * @return {string} Band: GREEN, AMBER, or RED
 */
function computeRLBand(percentUsed) {
  if (percentUsed <= 0.90) return 'GREEN';
  if (percentUsed <= 0.95) return 'AMBER';
  return 'RED';
}
/**
 * Gets RL band with formatted percentage
 * @param {number} spent - Amount spent
 * @param {number} budget - Total budget
 * @return {Object} {band, percent, percentFormatted}
 */
function getRLBandInfo(spent, budget) {
  if (budget <= 0) {
    return { band: 'RED', percent: 0, percentFormatted: '0%' };
  }
  const percent = spent / budget;
  const band = computeRLBand(percent);
  const percentFormatted = formatPercent(percent, 1);
  return { band, percent, percentFormatted };
}