/**
 * syncBPFromSources
 *
 * Keeps the BP_Total sheet in sync with:
 *  - Attendance_Missions (Attendance Mission Points)
 *  - Flag_Missions      (Flag Mission Points)
 *  - Dice Roll Points   (Points / Dice Roll Points)
 *
 * For each player in BP_Total:
 *  - Pulls their sub-totals from the three source sheets
 *  - Writes them into BP_Total
 *  - Recomputes Current_BP = Attendance + Flag + Dice + Historical_BP
 *  - Updates LastUpdated
 */
function syncBPFromSources() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bpSheet   = ss.getSheetByName('BP_Total');
  const attSheet  = ss.getSheetByName('Attendance_Missions');
  const flagSheet = ss.getSheetByName('Flag_Missions');
  const diceSheet = ss.getSheetByName('Dice Roll Points'); // change if your tab is named differently

  if (!bpSheet || !attSheet || !flagSheet || !diceSheet) {
    throw new Error('One or more required sheets are missing. Needed: BP_Total, Attendance_Missions, Flag_Missions, Dice Roll Points');
  }

  // Helper to find header index from a list of possible names
  function getHeaderIndex(headerRow, namesArray) {
    for (var i = 0; i < namesArray.length; i++) {
      var idx = headerRow.indexOf(namesArray[i]);
      if (idx !== -1) return idx;
    }
    return -1;
  }

  // -------------------------------
  // Build maps from source sheets
  // -------------------------------

  // Attendance_Missions
  var attValues = attSheet.getDataRange().getValues();
  var attHeader = attValues[0];

  var attNameCol = getHeaderIndex(attHeader, ['PreferredName', 'preferred_name_id']);
  var attPointsCol = getHeaderIndex(attHeader, ['Attendance Mission Points', 'Attendance Missions']);

  if (attNameCol === -1 || attPointsCol === -1) {
    throw new Error('Attendance_Missions is missing "PreferredName/preferred_name_id" or "Attendance Mission Points" header.');
  }

  var attendanceMap = {};
  for (var i = 1; i < attValues.length; i++) {
    var row = attValues[i];
    var name = row[attNameCol];
    if (!name) continue;
    var pts = Number(row[attPointsCol]) || 0;
    attendanceMap[name] = pts;
  }

  // Flag_Missions
  var flagValues = flagSheet.getDataRange().getValues();
  var flagHeader = flagValues[0];

  var flagNameCol = getHeaderIndex(flagHeader, ['PreferredName', 'preferred_name_id']);
  var flagPointsCol = getHeaderIndex(flagHeader, ['Flag Mission Points', 'Flag Points', 'Flag_Points']);

  if (flagNameCol === -1 || flagPointsCol === -1) {
    throw new Error('Flag_Missions is missing "PreferredName/preferred_name_id" or "Flag Mission Points / Flag Points" header.');
  }

  var flagMap = {};
  for (var j = 1; j < flagValues.length; j++) {
    var fRow = flagValues[j];
    var fName = fRow[flagNameCol];
    if (!fName) continue;
    var fPts = Number(fRow[flagPointsCol]) || 0;
    flagMap[fName] = fPts;
  }

  // Dice Roll Points
  var diceValues = diceSheet.getDataRange().getValues();
  var diceHeader = diceValues[0];

  var diceNameCol = getHeaderIndex(diceHeader, ['PreferredName', 'preferred_name_id']);
  var dicePointsCol = getHeaderIndex(diceHeader, ['Points', 'Dice Roll Points']);

  if (diceNameCol === -1 || dicePointsCol === -1) {
    throw new Error('Dice Roll Points sheet is missing "PreferredName/preferred_name_id" or "Points/Dice Roll Points" header.');
  }

  var diceMap = {};
  for (var k = 1; k < diceValues.length; k++) {
    var dRow = diceValues[k];
    var dName = dRow[diceNameCol];
    if (!dName) continue;
    var dPts = Number(dRow[dicePointsCol]) || 0;
    diceMap[dName] = dPts;
  }

  // -------------------------------
  // Update BP_Total
  // -------------------------------
  var bpRange = bpSheet.getDataRange();
  var bpValues = bpRange.getValues();
  var bpHeader = bpValues[0];

  var bpNameCol      = getHeaderIndex(bpHeader, ['PreferredName', 'preferred_name_id']);
  var bpAttCol       = getHeaderIndex(bpHeader, ['Attendance Mission Points', 'Attendance Missions']);
  var bpFlagCol      = getHeaderIndex(bpHeader, ['Flag Mission Points', 'Flag Missions']);
  var bpDiceCol      = getHeaderIndex(bpHeader, ['Dice Roll Points']);
  var bpCurrentCol   = getHeaderIndex(bpHeader, ['Current_BP', 'Current BP']);
  var bpHistoricalCol= getHeaderIndex(bpHeader, ['Historical_BP', 'Historical BP']);
  var bpUpdatedCol   = getHeaderIndex(bpHeader, ['LastUpdated', 'Last Updated']);

  if (bpNameCol === -1 || bpAttCol === -1 || bpFlagCol === -1 || bpDiceCol === -1 || bpCurrentCol === -1 || bpHistoricalCol === -1 || bpUpdatedCol === -1) {
    throw new Error('BP_Total is missing one or more required headers (PreferredName, Attendance Mission Points, Flag Mission Points, Dice Roll Points, Current_BP, Historical_BP, LastUpdated).');
  }

  var now = new Date();
  var updatedRows = 0;

  for (var r = 1; r < bpValues.length; r++) {
    var bpRow = bpValues[r];
    var playerName = bpRow[bpNameCol];
    if (!playerName) continue; // skip blank rows

    var attPts  = attendanceMap[playerName] || 0;
    var flagPts = flagMap[playerName] || 0;
    var dicePts = diceMap[playerName] || 0;
    var histPts = Number(bpRow[bpHistoricalCol]) || 0;

    // Write sub-totals
    bpRow[bpAttCol]  = attPts;
    bpRow[bpFlagCol] = flagPts;
    bpRow[bpDiceCol] = dicePts;

    // Recompute Current_BP
    bpRow[bpCurrentCol] = attPts + flagPts + dicePts + histPts;

    // Timestamp
    bpRow[bpUpdatedCol] = now;

    updatedRows++;
  }

  // Bulk write back
  bpRange.setValues(bpValues);

  Logger.log('syncBPFromSources complete. Updated rows: ' + updatedRows);
}
