/**
 * Standardize headers on all event tabs.
 *
 * Forces row 1 on every event sheet to be:
 *   Rank | preferred_name_id | R1_Prize | R2_Prize | R3_Prize | End_Prizes
 *
 * Existing headers in row 1 are cleared, then the new ones are written
 * and highlighted green.
 */
function standardizeAllEventHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Event tab name pattern:
  //  MM-DD-YYYY
  //  MM-DDX-YYYY  (X = A–Z suffix)
  // Adjust if you have other valid patterns.
  const eventTabRegex = /^\d{2}-\d{2}[A-Z]?-?\d{4}$/;

  // Target headers you requested
  const HEADERS = [
    'Rank',
    'preferred_name_id',
    'R1_Prize',
    'R2_Prize',
    'R3_Prize',
    'End_Prizes'
  ];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (!eventTabRegex.test(name)) {
      return; // skip non-event sheets
    }

    const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);

    // Clear existing header row content
    if (lastCol > 0) {
      sheet.getRange(1, 1, 1, lastCol).clearContent();
    }

    // Write new headers into A1:F1
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    // Highlight headers green & bold for visibility
    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setBackground('#b7e1cd') // light green
      .setFontWeight('bold');
  });
}
