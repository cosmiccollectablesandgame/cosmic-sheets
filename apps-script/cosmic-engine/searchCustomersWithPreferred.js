/**
 * Search customers across PreferredNames, Members, and existing Preorders
 * Call this from the Sell Preorder dialog
 * 
 * @param {string} query - Search string (minimum 2 characters)
 * @returns {Array} Array of customer objects with name, source, phone, email, hasPreorders
 */
function searchCustomersWithPreferred(query) {
  if (!query || query.length < 2) return [];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  const seenNames = new Set();
  const queryLower = query.toLowerCase();
  
  // 1. Search PreferredNames tab first (highest priority)
  try {
    const prefSheet = ss.getSheetByName('PreferredNames');
    if (prefSheet) {
      const prefData = prefSheet.getDataRange().getValues();
      const prefHeaders = prefData[0];
      
      // Find column indices (adjust these based on your actual column names)
      const nameCol = findColumnIndex(prefHeaders, ['name', 'preferred name', 'preferredname', 'player name', 'playername']);
      const phoneCol = findColumnIndex(prefHeaders, ['phone', 'phone number', 'phonenumber']);
      const emailCol = findColumnIndex(prefHeaders, ['email', 'email address', 'emailaddress']);
      
      for (let i = 1; i < prefData.length; i++) {
        const row = prefData[i];
        const name = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
        
        if (name && name.toLowerCase().includes(queryLower)) {
          const nameKey = name.toLowerCase();
          if (!seenNames.has(nameKey)) {
            seenNames.add(nameKey);
            results.push({
              name: name,
              source: 'PreferredNames',
              phone: phoneCol >= 0 ? String(row[phoneCol] || '').trim() : '',
              email: emailCol >= 0 ? String(row[emailCol] || '').trim() : '',
              hasPreorders: checkHasPreorders(ss, name)
            });
          }
        }
      }
    }
  } catch (e) {
    console.log('PreferredNames search error: ' + e.message);
  }
  
  // 2. Search Members tab
  try {
    const membersSheet = ss.getSheetByName('Members');
    if (membersSheet) {
      const membersData = membersSheet.getDataRange().getValues();
      const membersHeaders = membersData[0];
      
      const nameCol = findColumnIndex(membersHeaders, ['name', 'member name', 'membername', 'full name', 'fullname']);
      const phoneCol = findColumnIndex(membersHeaders, ['phone', 'phone number', 'phonenumber']);
      const emailCol = findColumnIndex(membersHeaders, ['email', 'email address', 'emailaddress']);
      
      for (let i = 1; i < membersData.length; i++) {
        const row = membersData[i];
        const name = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
        
        if (name && name.toLowerCase().includes(queryLower)) {
          const nameKey = name.toLowerCase();
          if (!seenNames.has(nameKey)) {
            seenNames.add(nameKey);
            results.push({
              name: name,
              source: 'Members',
              phone: phoneCol >= 0 ? String(row[phoneCol] || '').trim() : '',
              email: emailCol >= 0 ? String(row[emailCol] || '').trim() : '',
              hasPreorders: checkHasPreorders(ss, name)
            });
          }
        }
      }
    }
  } catch (e) {
    console.log('Members search error: ' + e.message);
  }
  
  // 3. Search existing Preorders (for customers not in other lists)
  try {
    const preordersSheet = ss.getSheetByName('Preorders') || ss.getSheetByName('Preorders_Log');
    if (preordersSheet) {
      const preordersData = preordersSheet.getDataRange().getValues();
      const preordersHeaders = preordersData[0];
      
      const nameCol = findColumnIndex(preordersHeaders, ['customer', 'customer name', 'customername', 'name']);
      const contactCol = findColumnIndex(preordersHeaders, ['contact', 'contact info', 'contactinfo', 'phone', 'email']);
      
      for (let i = 1; i < preordersData.length; i++) {
        const row = preordersData[i];
        const name = nameCol >= 0 ? String(row[nameCol] || '').trim() : '';
        
        if (name && name.toLowerCase().includes(queryLower)) {
          const nameKey = name.toLowerCase();
          if (!seenNames.has(nameKey)) {
            seenNames.add(nameKey);
            const contact = contactCol >= 0 ? String(row[contactCol] || '').trim() : '';
            
            results.push({
              name: name,
              source: 'Preorders',
              phone: contact.includes('@') ? '' : contact,
              email: contact.includes('@') ? contact : '',
              hasPreorders: true
            });
          }
        }
      }
    }
  } catch (e) {
    console.log('Preorders search error: ' + e.message);
  }
  
  // Sort: PreferredNames first, then Members, then Preorders
  const sourceOrder = { 'PreferredNames': 0, 'Members': 1, 'Preorders': 2 };
  results.sort((a, b) => {
    const orderDiff = sourceOrder[a.source] - sourceOrder[b.source];
    if (orderDiff !== 0) return orderDiff;
    return a.name.localeCompare(b.name);
  });
  
  return results.slice(0, 15); // Limit to 15 results
}

/**
 * Helper: Find column index by checking multiple possible header names
 */
function findColumnIndex(headers, possibleNames) {
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i] || '').toLowerCase().trim();
    if (possibleNames.includes(header)) {
      return i;
    }
  }
  return -1;
}

/**
 * Helper: Check if customer has any active preorders
 */
function checkHasPreorders(ss, customerName) {
  try {
    const preordersSheet = ss.getSheetByName('Preorders') || ss.getSheetByName('Preorders_Log');
    if (!preordersSheet) return false;
    
    const data = preordersSheet.getDataRange().getValues();
    const headers = data[0];
    
    const nameCol = findColumnIndex(headers, ['customer', 'customer name', 'customername', 'name']);
    const statusCol = findColumnIndex(headers, ['status', 'order status', 'orderstatus']);
    
    if (nameCol < 0) return false;
    
    const nameLower = customerName.toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      const rowName = String(data[i][nameCol] || '').toLowerCase().trim();
      if (rowName === nameLower) {
        // If there's a status column, check if it's active
        if (statusCol >= 0) {
          const status = String(data[i][statusCol] || '').toLowerCase();
          if (status === '' || status === 'pending' || status === 'active' || status === 'open') {
            return true;
          }
        } else {
          // No status column, assume all are active
          return true;
        }
      }
    }
    
    return false;
  } catch (e) {
    console.log('checkHasPreorders error: ' + e.message);
    return false;
  }
}