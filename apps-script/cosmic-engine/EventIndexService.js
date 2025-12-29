/**
 * Event Index Service - Event Summary Dashboard
 * @fileoverview Provides high-level event metrics for the Event Index UI
 */

// ============================================================================
// TYPE DEFINITIONS
// ============================================================================

/**
 * @typedef {Object} EventIndexRow
 * @property {string} sheetName        - Event tab name (e.g., "11-29-2025")
 * @property {string} displayName      - Formatted display name with format suffix
 * @property {string} date             - ISO-style date "YYYY-MM-DD"
 * @property {string} suffix           - Suffix after date (e.g., "A", "B", "1", or "")
 * @property {string} format           - Event format (Commander, Draft, etc.)
 * @property {number|null} playerCount - Number of players in roster
 * @property {number|null} rlPercentUsed - Percentage of RL budget used (0-100)
 * @property {number|null} revenue     - Total entry fees collected
 * @property {number|null} prizeSpend  - Actual prize cost from Spent_Pool
 * @property {number|null} margin      - Revenue minus prizeSpend
 * @property {string} status           - "Open", "Locked", "Closed", "Unknown"
 * @property {string|null} lastUpdated - Last modification timestamp
 */

/**
 * @typedef {Object} EventIndexResult
 * @property {EventIndexRow[]} events  - Array of event rows
 * @property {number} totalEvents      - Total count before limit applied
 */

// ============================================================================
// MAIN PUBLIC FUNCTION
// ============================================================================

/**
 * Returns a list of event summary rows for the Event Index UI.
 *
 * @param {Object=} filters - Optional filter parameters
 * @param {number=} filters.limit - Max events to return (default: 100)
 * @param {string=} filters.format - Filter by format (e.g., "Commander", "Draft", or null for all)
 * @param {string=} filters.startDate - "YYYY-MM-DD" inclusive start
 * @param {string=} filters.endDate - "YYYY-MM-DD" inclusive end
 * @return {EventIndexResult} Result with events array and totalEvents count
 */
function getEventIndex(filters) {
  filters = filters || {};
  const limit = filters.limit || 100;
  const formatFilter = filters.format || null;
  const startDate = filters.startDate ? parseFilterDate_(filters.startDate) : null;
  const endDate = filters.endDate ? parseFilterDate_(filters.endDate) : null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Get all event tab names
  const eventTabNames = listEventTabs_();

  // Pre-load Spent_Pool data for efficiency
  const spentByEvent = loadSpentPoolByEvent_(ss);

  // Pre-load Integrity_Log for status/lastUpdated
  const integrityByEvent = loadIntegrityByEvent_(ss);

  // Build event rows
  const allEvents = [];

  eventTabNames.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Parse date and suffix from sheet name
    const parsed = parseEventDateFromName_(sheetName);
    if (!parsed) return;

    // Apply date filters
    if (startDate && parsed.dateObj < startDate) return;
    if (endDate && parsed.dateObj > endDate) return;

    // Get event metadata and compute metrics
    const eventRow = buildEventRow_(
      sheet,
      sheetName,
      parsed,
      spentByEvent,
      integrityByEvent
    );

    // Apply format filter
    if (formatFilter && formatFilter !== 'All') {
      // Keep Unknowns (could be mis-detected), but filter known non-matches
      if (eventRow.format !== 'Unknown' && eventRow.format !== formatFilter) {
        return;
      }
    }

    allEvents.push(eventRow);
  });

  // Sort by date descending (most recent first)
  allEvents.sort((a, b) => {
    const dateA = a.date || '';
    const dateB = b.date || '';
    return dateB.localeCompare(dateA);
  });

  const totalEvents = allEvents.length;

  // Apply limit
  const limitedEvents = allEvents.slice(0, limit);

  return {
    events: limitedEvents,
    totalEvents: totalEvents
  };
}

// ============================================================================
// PRIVATE HELPERS - EVENT TAB DISCOVERY
// ============================================================================

/**
 * Lists all event tabs (MM-DD-YYYY and MM-DDx-YYYY formats).
 * Uses existing listEventTabs if available, otherwise implements locally.
 * @return {Array<string>} Event tab names
 * @private
 */
function listEventTabs_() {
  // Check if global listEventTabs exists (from main/event service)
  if (typeof listEventTabs === 'function') {
    return listEventTabs();
  }

  // Fallback implementation:
  // - MM-DD-YYYY
  // - MM-DDx-YYYY (suffix letter)
  // - MM-DD-YYYY-suffix (legacy trailing suffix separated by dash)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const eventPattern = /^\d{2}-\d{2}([a-zA-Z])?-\d{4}(\b|[-_].*)?$/;

  return sheets
    .map(s => s.getName())
    .filter(name => eventPattern.test(name))
    .sort();
}

// ============================================================================
// PRIVATE HELPERS - DATE PARSING
// ============================================================================

/**
 * Parses event date and suffix from sheet name.
 * Supported:
 * 1) MM-DD-YYYY
 * 2) MM-DD-YYYY-anything (legacy trailing suffix separated by dash)
 * 3) MM-DDx-YYYY (suffix letter between day and year, your standard)
 *
 * @param {string} sheetName - Event sheet name
 * @return {Object|null} {dateStr, dateObj, suffix} or null if invalid
 * @private
 */
function parseEventDateFromName_(sheetName) {
  // Case 3: MM-DDx-YYYY
  let m = sheetName.match(/^(\d{2})-(\d{2})([a-zA-Z])-(\d{4})(.*)?$/);
  if (m) {
    const month = parseInt(m[1], 10);
    const day = parseInt(m[2], 10);
    const suffix = m[3] || '';
    const year = parseInt(m[4], 10);

    if (month < 1 || month > 12 || day < 1 || day > 31 || year < 2000 || year > 2100) return null;

    const dateObj = new Date(year, month - 1, day);
    const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

    return { dateStr, dateObj, suffix };
  }

  // Case 1/2: MM-DD-YYYY or MM-DD-YYYY-...
  m = sheetName.match(/^(\d{2})-(\d{2})-(\d{4})(.*)$/);
  if (!m) return null;

  const month = parseInt(m[1], 10);
  const day = parseInt(m[2], 10);
  const year = parseInt(m[3], 10);
  let suffix = m[4] || '';

  if (suffix.startsWith('-')) suffix = suffix.substring(1);

  if (month < 1 || month > 12 || day < 1 || day > 31 || year < 2000 || year > 2100) return null;

  const dateObj = new Date(year, month - 1, day);
  const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;

  return { dateStr, dateObj, suffix };
}

/**
 * Parses a filter date string (YYYY-MM-DD) to Date object.
 * @param {string} dateStr - Date string in YYYY-MM-DD format
 * @return {Date|null} Date object or null if invalid
 * @private
 */
function parseFilterDate_(dateStr) {
  if (!dateStr) return null;

  const match = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  const year = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const day = parseInt(match[3], 10);

  return new Date(year, month - 1, day);
}

// ============================================================================
// PRIVATE HELPERS - EVENT ROW BUILDER
// ============================================================================

/**
 * Builds a complete EventIndexRow for an event sheet.
 *
 * @param {Sheet} sheet - Event sheet
 * @param {string} sheetName - Sheet name
 * @param {Object} parsed - Parsed date info {dateStr, dateObj, suffix}
 * @param {Map} spentByEvent - Pre-loaded spent pool data
 * @param {Map} integrityByEvent - Pre-loaded integrity log data
 * @return {EventIndexRow} Event row object
 * @private
 */
function buildEventRow_(sheet, sheetName, parsed, spentByEvent, integrityByEvent) {
  // Get event properties (from developer metadata)
  const props = getEventPropsForIndex_(sheet);

  // Get format
  const format = getEventFormat_(sheet, props);

  // Get player count
  const playerCount = getEventPlayerCount_(sheet);

  // Get financials
  const financials = getEventFinancials_(sheet, props, playerCount, spentByEvent.get(sheetName));

  // Get status and lastUpdated
  const statusInfo = getEventStatus_(sheet, sheetName, integrityByEvent.get(sheetName));

  // Build display name
  let displayName = sheetName;
  if (format && format !== 'Unknown') {
    displayName = `${sheetName} - ${format}`;
  }

  return {
    sheetName: sheetName,
    displayName: displayName,
    date: parsed.dateStr,
    suffix: parsed.suffix,
    format: format,
    playerCount: playerCount,
    rlPercentUsed: financials.rlPercentUsed,
    revenue: financials.revenue,
    prizeSpend: financials.prizeSpend,
    margin: financials.margin,
    status: statusInfo.status,
    lastUpdated: statusInfo.lastUpdated
  };
}

// ============================================================================
// PRIVATE HELPERS - EVENT PROPERTIES
// ============================================================================

/**
 * Gets event properties from developer metadata.
 * Uses existing getEventProps if available.
 *
 * @param {Sheet} sheet - Event sheet
 * @return {Object} Event properties
 * @private
 */
function getEventPropsForIndex_(sheet) {
  // Try to use existing function if available
  if (typeof getEventProps === 'function') {
    return getEventProps(sheet);
  }

  // Fallback: read developer metadata directly
  const metadata = sheet.getDeveloperMetadata();
  const props = {};

  metadata.forEach(m => {
    const key = m.getKey();
    const value = m.getValue();

    if (key === 'entry' || key === 'kit_cost_per_player') {
      props[key] = parseFloat(value) || 0;
    } else {
      props[key] = value;
    }
  });

  return props;
}

// ============================================================================
// PRIVATE HELPERS - FORMAT DETECTION
// ============================================================================

/**
 * Infers event format from sheet and properties.
 *
 * @param {Sheet} sheet - Event sheet
 * @param {Object} props - Event properties
 * @return {string} Format name or "Unknown"
 * @private
 */
function getEventFormat_(sheet, props) {
  // Check props first (most reliable)
  const rawType =
    props.event_type ||
    props.eventType ||
    props.type ||
    props.format ||
    props.event_format;

  if (rawType) return formatEventType_(rawType);

  // Check for format in sheet name
  const name = sheet.getName().toLowerCase();
  if (name.includes('commander') || name.includes('cmdr')) return 'Commander';
  if (name.includes('draft')) return 'Draft';
  if (name.includes('sealed')) return 'Sealed';
  if (name.includes('standard')) return 'Standard';
  if (name.includes('modern')) return 'Modern';
  if (name.includes('pioneer')) return 'Pioneer';
  if (name.includes('legacy')) return 'Legacy';
  if (name.includes('pauper')) return 'Pauper';

  // Check A1 note for metadata JSON
  try {
    const a1Note = sheet.getRange('A1').getNote();
    if (a1Note) {
      try {
        const noteData = JSON.parse(a1Note);
        if (noteData.format) return formatEventType_(noteData.format);
        if (noteData.event_type) return formatEventType_(noteData.event_type);
      } catch (e) {}
    }

    // Check for format header in data
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const formatColIdx = headers.findIndex(h =>
      String(h).toLowerCase() === 'format' ||
      String(h).toLowerCase() === 'event_type'
    );

    if (formatColIdx !== -1 && sheet.getLastRow() > 1) {
      const formatValue = sheet.getRange(2, formatColIdx + 1).getValue();
      if (formatValue) return formatEventType_(String(formatValue));
    }
  } catch (e) {
    // Ignore errors reading format
  }

  return 'Unknown';
}

/**
 * Normalizes event type strings to display format.
 * @param {string} type - Raw event type
 * @return {string} Formatted type
 * @private
 */
function formatEventType_(type) {
  if (!type) return 'Unknown';

  const upper = String(type).toUpperCase();

  if (upper === 'CONSTRUCTED') return 'Constructed';
  if (upper === 'LIMITED') return 'Limited';
  if (upper === 'HYBRID') return 'Hybrid';
  if (upper === 'COMMANDER' || upper === 'CMDR') return 'Commander';
  if (upper === 'DRAFT') return 'Draft';
  if (upper === 'SEALED') return 'Sealed';
  if (upper === 'STANDARD') return 'Standard';
  if (upper === 'MODERN') return 'Modern';
  if (upper === 'PIONEER') return 'Pioneer';
  if (upper === 'LEGACY') return 'Legacy';
  if (upper === 'PAUPER') return 'Pauper';

  // Capitalize first letter for other types
  const s = String(type);
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

// ============================================================================
// PRIVATE HELPERS - PLAYER COUNT
// ============================================================================

/**
 * Counts players in an event sheet.
 *
 * @param {Sheet} sheet - Event sheet
 * @return {number|null} Player count or null if unable to determine
 * @private
 */
function getEventPlayerCount_(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 0; // Only header or empty

    // Look for PreferredName column
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameColIdx = headers.findIndex(h => {
      const normalized = String(h).toLowerCase().replace(/_/g, '');
      return normalized === 'preferredname' ||
             normalized === 'playername' ||
             normalized === 'name' ||
             normalized === 'player';
    });

    if (nameColIdx !== -1) {
      // Count non-empty cells in name column
      const nameData = sheet.getRange(2, nameColIdx + 1, lastRow - 1, 1).getValues();
      let count = 0;
      nameData.forEach(row => {
        if (row[0] && String(row[0]).trim()) count++;
      });
      return count;
    }

    // Fallback: count non-empty rows (excluding header), assume column B has player names
    const colBData = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    let count = 0;
    colBData.forEach(row => {
      if (row[0] && String(row[0]).trim()) count++;
    });

    return count > 0 ? count : null;
  } catch (e) {
    return null;
  }
}

// ============================================================================
// PRIVATE HELPERS - FINANCIALS
// ============================================================================

/**
 * Computes financial metrics for an event.
 *
 * @param {Sheet} sheet - Event sheet
 * @param {Object} props - Event properties
 * @param {number|null} playerCount - Number of players
 * @param {Object|null} spentData - Spent pool data for this event
 * @return {Object} {rlPercentUsed, revenue, prizeSpend, margin}
 * @private
 */
function getEventFinancials_(sheet, props, playerCount, spentData) {
  const result = {
    rlPercentUsed: null,
    revenue: null,
    prizeSpend: null,
    margin: null
  };

  try {
    // Calculate revenue from entry fee × player count
    const entry = parseFloat(props.entry) || 0;
    if (entry > 0 && playerCount && playerCount > 0) {
      result.revenue = entry * playerCount;
    }

    // Get prize spend from Spent_Pool data
    if (spentData && spentData.total > 0) {
      result.prizeSpend = spentData.total;
    }

    // Calculate margin if both revenue and prizeSpend exist
    if (result.revenue !== null && result.prizeSpend !== null) {
      result.margin = result.revenue - result.prizeSpend;
    }

    // Calculate RL% used (based on RL budget, not raw revenue)
    if (result.prizeSpend !== null && result.revenue !== null && result.revenue > 0) {
      const throttle = (typeof getThrottleKV === 'function')
        ? getThrottleKV()
        : { RL_Percentage: '0.95' };

      const rlPercent = parseFloat(throttle.RL_Percentage || 0.95);

      // Adjust for kit cost if LIMITED-style event
      const eventType = String(props.event_type || props.eventType || '').toUpperCase();
      let eligibleRevenue = result.revenue;

      if (eventType === 'LIMITED' || eventType === 'DRAFT' || eventType === 'SEALED' || eventType === 'PRERELEASE') {
        const kitCost = parseFloat(props.kit_cost_per_player || props.kitCostPerPlayer || 0) || 0;
        eligibleRevenue = (entry - kitCost) * playerCount;
      }

      const rlBudget = eligibleRevenue * rlPercent;
      if (rlBudget > 0) {
        result.rlPercentUsed = (result.prizeSpend / rlBudget) * 100;
      }
    }

    // Try to read RL_Used_Percent directly from sheet if available (optional override)
    try {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const rlColIdx = headers.findIndex(h =>
        String(h).toLowerCase().includes('rl_used') ||
        String(h).toLowerCase().includes('rlpercent')
      );

      if (rlColIdx !== -1 && sheet.getLastRow() > 1) {
        const rlValue = sheet.getRange(2, rlColIdx + 1).getValue();
        if (rlValue && !isNaN(parseFloat(rlValue))) {
          result.rlPercentUsed = parseFloat(rlValue);
        }
      }
    } catch (e) {
      // Ignore
    }

  } catch (e) {
    // Return defaults on error
  }

  return result;
}

/**
 * Loads Spent_Pool data grouped by event ID.
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Map} Map of eventId -> {total, count}
 * @private
 */
function loadSpentPoolByEvent_(ss) {
  const map = new Map();

  const sheet = ss.getSheetByName('Spent_Pool');
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return map;

  // Expected columns (typical): Event_ID, Item_Code, Item_Name, Level, Qty, COGS, Total, Timestamp, Batch_ID, Reverted, Event_Type
  for (let i = 1; i < data.length; i++) {
    const eventId = data[i][0];
    const total = parseFloat(data[i][6]) || 0; // Total column
    const reverted = data[i][9]; // Reverted column

    if (!eventId || reverted === true || String(reverted).toUpperCase() === 'TRUE') continue;

    if (!map.has(eventId)) {
      map.set(eventId, { total: 0, count: 0 });
    }

    const entry = map.get(eventId);
    entry.total += total;
    entry.count++;
  }

  return map;
}

// ============================================================================
// PRIVATE HELPERS - STATUS & LAST UPDATED
// ============================================================================

/**
 * Determines event status and last update time.
 *
 * @param {Sheet} sheet - Event sheet
 * @param {string} sheetName - Sheet name
 * @param {Object|null} integrityData - Integrity log data for this event
 * @return {Object} {status, lastUpdated}
 * @private
 */
function getEventStatus_(sheet, sheetName, integrityData) {
  const result = {
    status: 'Unknown',
    lastUpdated: null
  };

  try {
    // Check for LOCKED flag in developer metadata
    const props = getEventPropsForIndex_(sheet);
    const lockedVal = (
      props.locked ??
      props.Locked ??
      props.LOCKED ??
      props.status ??
      props.event_status
    );

    if (String(lockedVal).toUpperCase() === 'LOCKED' || String(lockedVal).toLowerCase() === 'true') {
      result.status = 'Locked';
    }

    // Check integrity log for COMMIT actions (indicates completed/locked)
    if (integrityData) {
      if (integrityData.hasCommit) {
        result.status = 'Locked';
      }
      if (integrityData.lastTimestamp) {
        result.lastUpdated = integrityData.lastTimestamp;
      }
    }

    // If no commit found and not explicitly locked, infer by recency
    if (result.status === 'Unknown') {
      const parsed = parseEventDateFromName_(sheetName);
      if (parsed) {
        const eventDate = parsed.dateObj;
        const now = new Date();
        const daysDiff = Math.floor((now - eventDate) / (1000 * 60 * 60 * 24));

        if (daysDiff <= 7) {
          result.status = 'Open';
        } else {
          result.status = 'Unknown';
        }
      }
    }

  } catch (e) {
    // Return defaults on error
  }

  return result;
}

/**
 * Loads Integrity_Log data grouped by event ID.
 *
 * @param {Spreadsheet} ss - Active spreadsheet
 * @return {Map} Map of eventId -> {hasCommit, lastTimestamp, actions}
 * @private
 */
function loadIntegrityByEvent_(ss) {
  const map = new Map();

  const sheet = ss.getSheetByName('Integrity_Log');
  if (!sheet) return map;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return map;

  // Typical columns: Timestamp, Event_ID, Action, ...
  for (let i = 1; i < data.length; i++) {
    const timestamp = data[i][0];
    const eventId = data[i][1];
    const action = data[i][2];

    if (!eventId) continue;

    if (!map.has(eventId)) {
      map.set(eventId, {
        hasCommit: false,
        lastTimestamp: null,
        actions: []
      });
    }

    const entry = map.get(eventId);
    entry.actions.push(action);

    // Track if any COMMIT-style action exists
    if (action === 'COMMIT' || action === 'ROUND_ALLOCATE') {
      entry.hasCommit = true;
    }

    // Track latest timestamp (string compare ok if ISO-ish; otherwise still decent for "last seen")
    const tsStr = timestamp ? String(timestamp) : null;
    if (tsStr && (!entry.lastTimestamp || tsStr > entry.lastTimestamp)) {
      entry.lastTimestamp = tsStr;
    }
  }

  return map;
}

// ============================================================================
// MENU HANDLER
// ============================================================================

/**
 * Opens the Event Index sidebar.
 * NOTE: Renamed to avoid colliding with the canonical onViewEventIndex in Main file.
 */
function onViewEventIndex_() {
  const html = HtmlService.createHtmlOutputFromFile('ui/event_index')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}
