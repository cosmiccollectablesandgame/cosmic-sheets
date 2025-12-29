/**
 * Gets the current store credit balance for a player
 * @param {string} playerName - The preferred_name_id to look up
 * @returns {number} The current balance
 */
function getCurrentBalance(playerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = ss.getSheetByName('Store_Credit_Ledger');
  
  if (!ledger) {
    throw new Error('Store_Credit_Ledger sheet not found');
  }
  
  const data = ledger.getDataRange().getValues();
  const headers = data[0];
  
  const playerCol = headers.indexOf('preferred_name_id');
  const balanceCol = headers.indexOf('RunningBalance');
  
  if (playerCol === -1 || balanceCol === -1) {
    throw new Error('Required columns not found');
  }
  
  // Search from bottom up for most recent transaction
  for (let i = data.length - 1; i > 0; i--) {
    if (data[i][playerCol] === playerName) {
      return Number(data[i][balanceCol]) || 0;
    }
  }
  
  return 0;
}