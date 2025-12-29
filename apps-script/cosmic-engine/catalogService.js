/**
 * Catalog Service - Prize Catalog Operations
 * @fileoverview Manages Prize_Catalog: CRUD, normalization, dry-run, imports
 */
// ============================================================================
// CATALOG OPERATIONS
// ============================================================================
/**
 * Gets Prize_Catalog as array of objects
 * @return {Array<Object>} Catalog items
 */
function getCatalog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Catalog');
  if (!sheet) {
    throwError('Prize_Catalog sheet not found', 'CATALOG_MISSING', 'Run Build/Repair to create it');
  }
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only header or empty
  return toObjects(data);
}
/**
 * Gets catalog as Map keyed by Code
 * @return {Map<string, Object>} Catalog map
 */
function getCatalogMap() {
  const catalog = getCatalog();
  return toMapByKey(catalog, 'Code');
}
/**
 * Normalizes Prize_Catalog headers (synonym mapping)
 * @return {Object} {applied: boolean, mappings: Array}
 */
function normalizeHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Catalog');
  if (!sheet) {
    throwError('Prize_Catalog sheet not found', 'CATALOG_MISSING');
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const synonyms = getHeaderSynonyms();
  const mappings = [];
  let changed = false;
  const normalized = headers.map(header => {
    const canonical = normalizeHeader(header);
    if (canonical !== header) {
      mappings.push({ from: header, to: canonical });
      changed = true;
      return canonical;
    }
    return header;
  });
  if (changed) {
    sheet.getRange(1, 1, 1, normalized.length).setValues([normalized]);
    logIntegrityAction('CATALOG_HEADER_NORMALIZE', {
      details: `Mapped ${mappings.length} synonyms: ${mappings.map(m => m.from + '→' + m.to).join(', ')}`,
      status: 'SUCCESS'
    });
  }
  return { applied: changed, mappings };
}
/**
 * Updates catalog items
 * @param {Array<Object>} edits - Array of edits {code, field, value}
 */
function updateItems(edits) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Prize_Catalog');
  if (!sheet) {
    throwError('Prize_Catalog sheet not found', 'CATALOG_MISSING');
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeCol = headers.indexOf('Code');
  if (codeCol === -1) {
    throwError('Code column not found', 'CATALOG_INVALID');
  }
  let changeCount = 0;
  edits.forEach(edit => {
    const { code, field, value } = edit;
    const fieldCol = headers.indexOf(field);
    if (fieldCol === -1) return;
    // Find row with matching code
    for (let i = 1; i < data.length; i++) {
      if (data[i][codeCol] === code) {
        const oldValue = data[i][fieldCol];
        if (oldValue !== value) {
          sheet.getRange(i + 1, fieldCol + 1).setValue(value);
          changeCount++;
          logIntegrityAction('CATALOG_CHANGE', {
            details: `${code}.${field}: ${oldValue} → ${value}`,
            status: 'SUCCESS'
          });
        }
        break;
      }
    }
  });
  return { changed: changeCount };
}
// ============================================================================
// DRY RUN
// ============================================================================
/**
 * Dry-runs catalog against recent events
 * @param {number} eventCount - Number of recent events to test (default: 3)
 * @return {Object} Dry-run report
 */
function dryRunAgainstEvents(eventCount = 3) {
  const eventTabs = listEventTabs().slice(-eventCount);
  const catalog = getCatalogMap();
  const issues = [];
  eventTabs.forEach(tab => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    // Check End_Prizes column (F)
    const headers = data[0];
    const endCol = headers.indexOf('End_Prizes');
    if (endCol === -1) return;
    for (let i = 1; i < data.length; i++) {
      const endPrizes = String(data[i][endCol]);
      if (!endPrizes) continue;
      // Parse prize codes (comma-separated)
      const codes = endPrizes.split(',').map(c => c.trim()).filter(c => c);
      codes.forEach(code => {
        if (!catalog.has(code)) {
          issues.push({
            event: tab,
            row: i + 1,
            code,
            issue: 'Code not found in catalog'
          });
        }
      });
    }
  });
  return {
    events_tested: eventTabs.length,
    issues_found: issues.length,
    issues
  };
}
// ============================================================================
// PREORDER IMPORT
// ============================================================================
/**
 * Imports preorder allocations (updates Projected_Qty only)
 * @param {string} csvText - CSV content
 * @return {Object} Import result
 */
function importPreorders(csvText) {
  const rows = parseCSV(csvText);
  if (rows.length === 0) {
    throwError('Empty CSV', 'IMPORT_EMPTY');
  }
  // Expected headers: Code, Name, Qty_Reserved, Expected_Date, Vendor
  const headers = rows[0];
  const codeCol = headers.findIndex(h => h.match(/code|sku/i));
  const qtyCol = headers.findIndex(h => h.match(/qty|quantity|reserved/i));
  if (codeCol === -1 || qtyCol === -1) {
    throwError('Missing required columns', 'IMPORT_INVALID', 'CSV must have Code and Qty columns');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const catalogSheet = ss.getSheetByName('Prize_Catalog');
  if (!catalogSheet) {
    throwError('Prize_Catalog not found', 'CATALOG_MISSING');
  }
  const catalogData = catalogSheet.getDataRange().getValues();
  const catalogHeaders = catalogData[0];
  const catalogCodeCol = catalogHeaders.indexOf('Code');
  const projectedCol = catalogHeaders.indexOf('Projected_Qty');
  if (catalogCodeCol === -1) {
    throwError('Code column not found in catalog', 'CATALOG_INVALID');
  }
  // Create Projected_Qty column if missing
  let actualProjectedCol = projectedCol;
  if (projectedCol === -1) {
    actualProjectedCol = catalogHeaders.length;
    catalogSheet.getRange(1, actualProjectedCol + 1).setValue('Projected_Qty');
  }
  let updatedCount = 0;
  for (let i = 1; i < rows.length; i++) {
    const code = rows[i][codeCol];
    const qty = parseInt(rows[i][qtyCol], 10);
    if (!code || isNaN(qty)) continue;
    // Find code in catalog
    for (let j = 1; j < catalogData.length; j++) {
      if (catalogData[j][catalogCodeCol] === code) {
        catalogSheet.getRange(j + 1, actualProjectedCol + 1).setValue(qty);
        updatedCount++;
        break;
      }
    }
  }
  logIntegrityAction('PREORDER_IMPORT', {
    details: `Updated ${updatedCount} items from ${rows.length - 1} preorder rows`,
    status: 'SUCCESS'
  });
  return {
    rows_processed: rows.length - 1,
    items_updated: updatedCount
  };
}

/**
 * Imports preorder allocation from UI (Option 1 - structured data)
 * @param {Object} payload - {setName: string, items: Array<Object>}
 * @return {Object} Import result
 */
function importPreorderAllocationFromUI(payload) {
  if (!payload || !payload.setName || !payload.items || payload.items.length === 0) {
    throwError('Invalid payload', 'IMPORT_INVALID', 'Missing setName or items');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { setName, items } = payload;

  // Ensure Preorders_Buckets sheet exists
  let bucketsSheet = ss.getSheetByName('Preorders_Buckets');
  if (!bucketsSheet) {
    bucketsSheet = ss.insertSheet('Preorders_Buckets');
    bucketsSheet.appendRow([
      'Set_Name',
      'Item_Name',
      'Item_Code',
      'Unit_Cost',
      'Unit_Price',
      'Quantity',
      'Date_Added',
      'Status'
    ]);
    bucketsSheet.setFrozenRows(1);
    bucketsSheet.getRange('A1:H1')
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
  }

  const timestamp = new Date();
  const rowsToAdd = [];

  items.forEach(item => {
    rowsToAdd.push([
      setName,
      item.itemName,
      item.itemCode || '',
      item.unitCost || '',
      item.unitPrice || '',
      item.quantity,
      timestamp,
      'Active'
    ]);
  });

  // Append all rows at once
  if (rowsToAdd.length > 0) {
    bucketsSheet.getRange(
      bucketsSheet.getLastRow() + 1,
      1,
      rowsToAdd.length,
      8
    ).setValues(rowsToAdd);
  }

  // Log the action
  logIntegrityAction('PREORDER_UI_IMPORT', {
    details: `Imported ${items.length} items for set "${setName}" to Preorders_Buckets`,
    status: 'SUCCESS'
  });

  return {
    success: true,
    imported: items.length,
    setName: setName,
    items: items.map(item => ({
      name: item.itemName,
      code: item.itemCode,
      qty: item.quantity
    }))
  };
}

// ============================================================================
// SCHEMA HELPERS
// ============================================================================
/**
 * Ensures Prize_Catalog has required headers
 */
function ensureCatalogSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Prize_Catalog');
  const requiredHeaders = [
    'Code', 'Name', 'Level', 'Rarity', 'COGS', 'EV_Cost', 'Qty',
    'Eligible_Rounds', 'Eligible_End', 'Player_Threshold', 'InStock',
    'EV_Explanation', 'Round_Weight', 'PV_Multiplier', 'Projected_Qty'
  ];
  if (!sheet) {
    sheet = ss.insertSheet('Prize_Catalog');
    sheet.appendRow(requiredHeaders);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:O1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    return;
  }
  // Check and add missing headers
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = requiredHeaders.filter(h => !headers.includes(h));
  if (missing.length > 0) {
    const startCol = headers.length + 1;
    missing.forEach((header, idx) => {
      sheet.getRange(1, startCol + idx).setValue(header);
    });
  }
}