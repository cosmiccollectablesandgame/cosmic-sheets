/**
 * ======================================================================
 * Undiscovered Names Scan (v2 - simple, bulletproof)
 * ======================================================================
 * - Reads canonical names from PreferredNames!A:A
 *   (header optional; we auto-detect and skip it if present)
 * - Scans ALL event tabs (MM-DD-YYYY or MM-DDX-YYYY)
 * - Reads player names from column B on those event tabs
 * - Writes unknown names to UndiscoveredNames:
 *     Potential_Name | First_Seen_Sheet | Reviewed
 */

/**
 * Main entry – run this from Apps Script or wire to a menu.
 */
function runUndiscoveredNamesScan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const PREFERRED_SHEET_NAME = 'PreferredNames';
  const UNDISCOVERED_SHEET_NAME = 'UndiscoveredNames';

  const preferredSheet = ss.getSheetByName(PREFERRED_SHEET_NAME);
  if (!preferredSheet) {
    throw new Error('Sheet "PreferredNames" not found.');
  }

  // --------------------------------------------------------------------
  // 1) Build canonical name set from PreferredNames!A:A
  // --------------------------------------------------------------------
  const prefRange = preferredSheet.getDataRange();
  const prefValues = prefRange.getValues();
  if (prefValues.length === 0) {
    throw new Error('PreferredNames sheet is empty.');
  }

  const preferredNameSet = new Set();

  // Detect if first row looks like a header (e.g. "PreferredName", "Name", etc.)
  let startRow = 0;
  const firstCell = String(prefValues[0][0] || '').toLowerCase();
  if (firstCell.includes('name') || firstCell.includes('preferred')) {
    startRow = 1; // skip header
  }

  for (let r = startRow; r < prefValues.length; r++) {
    const name = normalizeName_(prefValues[r][0]);
    if (name) preferredNameSet.add(name);
  }

  // --------------------------------------------------------------------
  // 2) Prepare UndiscoveredNames sheet and existing unknown set
  // --------------------------------------------------------------------
  const undiscoveredSheet = getOrCreateUndiscoveredSheet_(ss, UNDISCOVERED_SHEET_NAME);
  const undData = undiscoveredSheet.getDataRange().getValues();

  const existingUnknownSet = new Set();
  if (undData.length > 1) {
    const header = undData[0];
    const potentialIdx = header.indexOf('Potential_Name');
    for (let r = 1; r < undData.length; r++) {
      const name = normalizeName_(undData[r][potentialIdx]);
      if (name) existingUnknownSet.add(name);
    }
  }

  const newUnknownSet = new Set();
  const rowsToAppend = [];

  // --------------------------------------------------------------------
  // 3) Scan all event sheets (MM-DD-YYYY / MM-DDX-YYYY)
  // --------------------------------------------------------------------
  const sheets = ss.getSheets();
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (!isEventSheetName_(sheetName)) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return; // header or empty

    // We assume column B has the player names (preferred_name_id)
    for (let r = 1; r < values.length; r++) {
      const rawName = values[r][1]; // column B (0-based index)
      const name = normalizeName_(rawName);
      if (!name) continue;

      // Skip if already canonical
      if (preferredNameSet.has(name)) continue;

      // Skip if already known unknown (existing or this run)
      if (existingUnknownSet.has(name)) continue;
      if (newUnknownSet.has(name)) continue;

      newUnknownSet.add(name);
      rowsToAppend.push([name, sheetName, false]); // Reviewed = FALSE
    }
  });

  // --------------------------------------------------------------------
  // 4) Append new unknown names
  // --------------------------------------------------------------------
  if (rowsToAppend.length > 0) {
    const lastRow = undiscoveredSheet.getLastRow();
    let writeStartRow = lastRow + 1;

    if (lastRow === 0) {
      // Sheet somehow blank – re-init headers before writing
      initUndiscoveredHeaders_(undiscoveredSheet);
      writeStartRow = 2;
    }

    undiscoveredSheet.insertRowsAfter(lastRow || 1, rowsToAppend.length);
    undiscoveredSheet
      .getRange(writeStartRow, 1, rowsToAppend.length, 3)
      .setValues(rowsToAppend);
  }

  Logger.log('UndiscoveredNames scan complete. New names added: ' + rowsToAppend.length);
}

/**
 * Event sheet detection: MM-DD-YYYY or MM-DDX-YYYY
 */
function isEventSheetName_(name) {
  if (!name) return false;
  const re = /^\d{2}-\d{2}[A-Z]?-\d{4}$/;
  return re.test(name);
}

/**
 * Normalizes a name to a trimmed string (or '' if empty).
 */
function normalizeName_(value) {
  if (value === null || value === undefined) return '';
  const s = String(value).trim();
  return s;
}

/**
 * Get or create UndiscoveredNames with proper headers.
 */
function getOrCreateUndiscoveredSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initUndiscoveredHeaders_(sheet);
    return sheet;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length === 0 || data[0].length === 0 || data[0][0] === '') {
    initUndiscoveredHeaders_(sheet);
  }
  return sheet;
}

/**
 * Initialize headers on UndiscoveredNames.
 */
function initUndiscoveredHeaders_(sheet) {
  sheet.clear();
  const headers = ['Potential_Name', 'First_Seen_Sheet', 'Reviewed'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}
