/**
 * callLog.js — Absence log sheet management for COMET.
 * VERSION: 0.6.1
 *
 * This file owns all server-side logic for the Absence Log feature:
 *   - Creating and managing a single Call_Log sheet
 *   - Reading absence entries for a date, week, month, or search query
 *   - Writing new absence entries (JSON-serialized)
 *   - Triggering absence notification emails (construction/sending lives in notifier.js)
 *   - Purging old entries based on retention policy
 *
 * CALL LOG SHEET STRUCTURE (v0.6.0):
 *   Single persistent sheet named "Call_Log" with 4 columns:
 *   Col A — Date (JS Date object; used for retention and filtering)
 *   Col B — Employee ID (string; for dedup)
 *   Col C — Entry JSON (string; contains all absence entry fields)
 *   Col D — Sent (string; blank if unsent, "sent at HH:MM" or "auto sent at HH:MM" if processed)
 *
 * RETENTION POLICY:
 *   Entries older than CALL_LOG_RETENTION_DAYS are auto-purged by a time-trigger.
 *   Default: 365 days (one year rolling window).
 *   Purge function: purgeExpiredCallLogEntries_() runs monthly.
 */


// ---------------------------------------------------------------------------
// Read Entries
// ---------------------------------------------------------------------------

/**
 * Returns all absence entries for a specific date.
 * Filters readAllCallLogEntries_() by date.
 *
 * @param {string} dateString — ISO date string (YYYY-MM-DD).
 * @returns {Array<CallLogEntry>}
 */
function getCallLogEntriesForDate_(dateString) {
  const timeZone = Session.getScriptTimeZone();
  const targetDateKey = dateString; // already normalized in ISO format

  return readAllCallLogEntries_().filter(entry => {
    const entryDateKey = Utilities.formatDate(entry.date, timeZone, 'yyyy-MM-dd');
    return entryDateKey === targetDateKey;
  });
}

/**
 * Returns all absence entries for the week containing the given date.
 * Week runs Monday–Sunday.
 *
 * @param {string} dateString — ISO date string (YYYY-MM-DD).
 * @returns {{ weekStart: string, weekEnd: string, entries: Array<CallLogEntry> }}
 */
function getCallLogEntriesForWeek_(dateString) {
  const date = new Date(dateString + 'T00:00:00');
  const monday = getMondayOfWeek_(date);
  const sunday = new Date(monday);
  sunday.setDate(sunday.getDate() + 6);

  const timeZone = Session.getScriptTimeZone();
  const weekStart = Utilities.formatDate(monday, timeZone, 'yyyy-MM-dd');
  const weekEnd = Utilities.formatDate(sunday, timeZone, 'yyyy-MM-dd');

  const allEntries = readAllCallLogEntries_();
  const filtered = allEntries.filter(entry => {
    const entryDateKey = Utilities.formatDate(entry.date, timeZone, 'yyyy-MM-dd');
    return entryDateKey >= weekStart && entryDateKey <= weekEnd;
  });

  // Sort by date ascending (oldest first in the week)
  filtered.sort((a, b) => a.date - b.date);

  return { weekStart, weekEnd, entries: filtered };
}

/**
 * Returns all absence entries for the month containing the given date.
 *
 * @param {string} dateString — ISO date string (YYYY-MM-DD).
 * @returns {{ month: number, year: number, entries: Array<CallLogEntry> }}
 */
function getCallLogEntriesForMonth_(dateString) {
  const date = new Date(dateString + 'T00:00:00');
  const year = date.getFullYear();
  const month = date.getMonth(); // 0-indexed

  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);

  const timeZone = Session.getScriptTimeZone();
  const monthStartKey = Utilities.formatDate(firstDay, timeZone, 'yyyy-MM-dd');
  const monthEndKey = Utilities.formatDate(lastDay, timeZone, 'yyyy-MM-dd');

  const allEntries = readAllCallLogEntries_();
  const filtered = allEntries.filter(entry => {
    const entryDateKey = Utilities.formatDate(entry.date, timeZone, 'yyyy-MM-dd');
    return entryDateKey >= monthStartKey && entryDateKey <= monthEndKey;
  });

  // Sort by date ascending
  filtered.sort((a, b) => a.date - b.date);

  return { month: month + 1, year, entries: filtered };
}

/**
 * Returns all absence entries that match a search query (by employee name).
 * Case-insensitive partial match on the name field.
 * Results sorted by date descending (newest first).
 *
 * @param {string} query — Search query (e.g. "Smith" or "john").
 * @returns {{ query: string, entries: Array<CallLogEntry> }}
 */
function searchCallLogEntries_(query) {
  const queryLower = query.toLowerCase().trim();
  if (queryLower.length === 0) return { query, entries: [] };

  const allEntries = readAllCallLogEntries_();
  const filtered = allEntries.filter(entry =>
    entry.name.toLowerCase().includes(queryLower)
  );

  // Sort by date descending (newest first)
  filtered.sort((a, b) => b.date - a.date);

  return { query, entries: filtered };
}


// ---------------------------------------------------------------------------
// Write Entry
// ---------------------------------------------------------------------------

/**
 * Appends a new absence entry to the Call_Log sheet.
 * Creates the sheet if it does not exist yet.
 * Entry fields are serialized to JSON in column C.
 *
 * @param {{
 *   name:             string,
 *   employeeId:       string,
 *   department:       string,
 *   time:             string,
 *   manager:          string,
 *   scheduledShift:   string,
 *   isCallout:        boolean,
 *   isFmla:           boolean,
 *   isNoShow:         boolean,
 *   comment:          string,
 * }} data
 * @returns {{ sheetName: string, rowNumber: number }}
 */
function appendCallLogEntry_(data) {
  const sheet = getOrCreateCallLogSheet_();
  const today = new Date();

  // Build the entry JSON object
  const entryJson = {
    name:             data.name            || '',
    department:       data.department      || '',
    isCallout:        !!data.isCallout,
    isFmla:           !!data.isFmla,
    isNoShow:         !!data.isNoShow,
    isLate:           !!data.isLate,
    time:             data.time            || '',
    manager:          data.manager         || '',
    scheduledShift:   data.scheduledShift  || '',
    comment:          data.comment         || '',
  };

  // Build the 4-column row
  const row = [
    today,                           // DATE
    data.employeeId || '',           // EMPLOYEE_ID
    JSON.stringify(entryJson),       // ENTRY_JSON
    '',                              // SENT (empty until marked as sent)
  ];

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 4).setValues([row]);
  SpreadsheetApp.flush();

  return { sheetName: CALL_LOG_SHEET_NAME, rowNumber: newRow };
}


// ---------------------------------------------------------------------------
// Email Notification
// ---------------------------------------------------------------------------

/**
 * Sends an absence notification email for the entry at the given row.
 * Reads the row from the Call_Log sheet, parses the entry JSON, and
 * determines recipients from the department's mailing list.
 *
 * When sendEmails is false in COMET Config, logs instead of sending.
 *
 * @param {string} sheetName  — Expected to be CALL_LOG_SHEET_NAME.
 * @param {number} rowNumber  — 1-indexed row number to send notification for.
 * @returns {{ sent: boolean, recipients: string[], sentLabel?: string, alreadySent?: boolean, disabled?: boolean, noRecipients?: boolean }}
 */
function sendAbsenceNotification_(sheetName, rowNumber, isAutoSend = false) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(CALL_LOG_SHEET_NAME);

  if (!sheet) throw new Error(`Call_Log sheet not found.`);

  const row = sheet.getRange(rowNumber, 1, 1, 4).getValues()[0];

  // Check if already sent
  const sentCell = row[CALL_LOG_COLUMN.SENT];
  if (sentCell && String(sentCell).trim().length > 0) {
    const entry = parseCallLogRow_(row);
    console.log(`callLog: Email already sent for ${entry.name} — marked as "${sentCell}"`);
    return { sent: false, alreadySent: true, recipients: [] };
  }

  const entry = parseCallLogRow_(row);

  // Check if email sending is enabled
  const config = readCometConfig_(); // setup.js
  const shouldSendEmails = !!config.sendEmails;

  if (!shouldSendEmails) {
    console.log(`callLog: Email sending is disabled in settings for ${entry.name}`);
    return { sent: false, disabled: true, recipients: [] };
  }

  const recipients = resolveAbsenceRecipients_(entry.department); // notifier.js
  if (recipients.length === 0) {
    console.warn(`callLog: No email recipients configured for ${entry.department}`);
    return { sent: false, noRecipients: true, recipients: [] };
  }

  sendAbsenceEmail_(entry, recipients); // notifier.js

  // Mark as sent with timestamp
  const now = new Date();
  const hours = String(now.getHours()).padStart(2, '0');
  const mins = String(now.getMinutes()).padStart(2, '0');
  const sentLabel = isAutoSend ? `auto sent at ${hours}:${mins}` : `sent at ${hours}:${mins}`;
  sheet.getRange(rowNumber, CALL_LOG_COLUMN.SENT + 1).setValue(sentLabel);
  SpreadsheetApp.flush();

  console.log(`callLog: Sent email for ${entry.name} to ${recipients.join(', ')} — marked as "${sentLabel}"`);
  return { sent: true, recipients, sentLabel };
}


// ---------------------------------------------------------------------------
// Sheet Management
// ---------------------------------------------------------------------------

/**
 * Returns the Call_Log sheet, creating it (with headers and formatting) if it doesn't exist.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateCallLogSheet_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const existing = workbook.getSheetByName(CALL_LOG_SHEET_NAME);
  if (existing) return existing;

  const sheet = workbook.insertSheet(CALL_LOG_SHEET_NAME);
  writeCallLogHeader_(sheet);
  return sheet;
}

/**
 * Writes the header row and applies formatting to the Call_Log sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function writeCallLogHeader_(sheet) {
  const headers = [
    'Date',
    'Employee ID',
    'Entry',
    'Sent',
  ];

  sheet.getRange(1, 1, 1, 4)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground(COLORS.PT_SHIFT) // Green (#57BB8A) from config.js
    .setFontColor(COLORS.HEADER_TEXT); // config.js

  sheet.setFrozenRows(1);
  sheet.setTabColor(COLORS.PT_SHIFT);

  // Set column widths
  sheet.setColumnWidth(CALL_LOG_COLUMN.DATE + 1,        110);
  sheet.setColumnWidth(CALL_LOG_COLUMN.EMPLOYEE_ID + 1, 120);
  sheet.setColumnWidth(CALL_LOG_COLUMN.ENTRY_JSON + 1,  600);
  sheet.setColumnWidth(CALL_LOG_COLUMN.SENT + 1,        150);
}

/**
 * Purges absence log entries older than CALL_LOG_RETENTION_DAYS.
 * Called by a monthly time-based trigger.
 * Runs as a batched delete operation (deletes from bottom to avoid shifting rows mid-operation).
 */
function purgeExpiredCallLogEntries_() {
  if (CALL_LOG_RETENTION_DAYS <= 0) {
    console.log('callLog: Retention is disabled (CALL_LOG_RETENTION_DAYS <= 0). Skipping purge.');
    return;
  }

  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(CALL_LOG_SHEET_NAME);

  if (!sheet) {
    console.log('callLog: Call_Log sheet does not exist yet. Nothing to purge.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < CALL_LOG_DATA_START_ROW) {
    console.log('callLog: Call_Log sheet is empty. Nothing to purge.');
    return;
  }

  const now = new Date();
  const cutoffDate = new Date(now.getTime() - CALL_LOG_RETENTION_DAYS * 86400000);

  const numRows = lastRow - CALL_LOG_DATA_START_ROW + 1;
  const data = sheet.getRange(CALL_LOG_DATA_START_ROW, 1, numRows, 1).getValues();

  const rowsToDelete = [];
  data.forEach((row, index) => {
    const dateCell = row[0];
    const date = coerceCallLogDate_(dateCell);
    if (date && date < cutoffDate) {
      rowsToDelete.push(CALL_LOG_DATA_START_ROW + index);
    }
  });

  // Delete from bottom to top to avoid shifting issues
  rowsToDelete.reverse();
  rowsToDelete.forEach(rowNum => {
    sheet.deleteRow(rowNum);
  });

  if (rowsToDelete.length > 0) {
    console.log(`callLog: Purged ${rowsToDelete.length} expired entries from Call_Log.`);
  }
}


// ---------------------------------------------------------------------------
// Internal Helpers
// ---------------------------------------------------------------------------

/**
 * Reads all rows from the Call_Log sheet and returns parsed CallLogEntry objects.
 * Used by all query functions to avoid duplicating sheet reads.
 *
 * @returns {Array<CallLogEntry>}
 */
function readAllCallLogEntries_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(CALL_LOG_SHEET_NAME);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CALL_LOG_DATA_START_ROW) return [];

  const numRows = lastRow - CALL_LOG_DATA_START_ROW + 1;
  const data = sheet.getRange(CALL_LOG_DATA_START_ROW, 1, numRows, 4).getValues();

  return data
    .map((row, offset) => parseCallLogRow_(row, CALL_LOG_DATA_START_ROW + offset))
    .filter(entry => entry !== null);
}

/**
 * Parses a 4-column row from the Call_Log sheet into a CallLogEntry object.
 * Returns null if the row is invalid or missing date/entry data.
 *
 * @typedef {{
 *   name:             string,
 *   employeeId:       string,
 *   isCallout:        boolean,
 *   isFmla:           boolean,
 *   isNoShow:         boolean,
 *   department:       string,
 *   time:             string,
 *   manager:          string,
 *   scheduledShift:   string,
 *   comment:          string,
 *   date:             Date,
 *   sheetName:        string,
 *   rowNumber:        number,
 * }} CallLogEntry
 *
 * @param {any[]} row — 4-column array from Call_Log sheet
 * @param {number} rowNumber — 1-indexed row number
 * @returns {CallLogEntry|null}
 */
function parseCallLogRow_(row, rowNumber = 0) {
  const dateCell = row[CALL_LOG_COLUMN.DATE];
  const employeeId = String(row[CALL_LOG_COLUMN.EMPLOYEE_ID] || '').trim();
  const entryJsonStr = String(row[CALL_LOG_COLUMN.ENTRY_JSON] || '').trim();

  const date = coerceCallLogDate_(dateCell);
  if (!date) return null;

  let entryData = {};
  if (entryJsonStr) {
    try {
      entryData = JSON.parse(entryJsonStr);
    } catch (e) {
      console.warn(`callLog: Failed to parse entry JSON at row ${rowNumber}: ${e.message}`);
      return null;
    }
  }

  return {
    name:             entryData.name             || '',
    employeeId,
    isCallout:        !!entryData.isCallout,
    isFmla:           !!entryData.isFmla,
    isNoShow:         !!entryData.isNoShow,
    isLate:           !!entryData.isLate,
    department:       entryData.department      || '',
    time:             entryData.time            || '',
    manager:          entryData.manager         || '',
    scheduledShift:   entryData.scheduledShift  || '',
    comment:          entryData.comment         || '',
    date,
    sheetName:        CALL_LOG_SHEET_NAME,
    rowNumber,
  };
}

/**
 * Normalizes a Call Log date cell value to a JavaScript Date, or null.
 * Handles GAS Date objects, Sheets serial numbers, and ISO strings.
 *
 * @param {any} cellValue
 * @returns {Date|null}
 */
function coerceCallLogDate_(cellValue) {
  if (!cellValue) return null;
  if (cellValue instanceof Date) return cellValue;
  if (typeof cellValue === 'number') {
    // Sheets date serial → JS Date
    return new Date((cellValue - SHEETS_EPOCH_OFFSET) * 86400000);
  }
  const parsed = new Date(cellValue);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Returns a new Date set to midnight on the Monday of the week containing date.
 *
 * @param {Date} date
 * @returns {Date}
 */
function getMondayOfWeek_(date) {
  const d = new Date(date);
  const day = d.getDay(); // 0 = Sunday
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}


// ---------------------------------------------------------------------------
// Auto-Send Notifications (Rolling 30-Minute Window)
// ---------------------------------------------------------------------------

/**
 * Auto-sends absence notifications for entries logged in the last 30 minutes.
 * Runs on a time-based trigger every 30 minutes.
 * Marks sent entries with "auto sent at HH:MM" in the SENT column.
 */
function autoSendAbsenceNotifications_() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CALL_LOG_SHEET_NAME);
    if (!sheet) {
      console.log('callLog: Call_Log sheet not found. Skipping autosend.');
      return;
    }

    const now = new Date();
    const windowStart = new Date(now.getTime() - 30 * 60 * 1000); // 30 min ago
    const lastRow = sheet.getLastRow();

    if (lastRow < CALL_LOG_DATA_START_ROW) return;

    const numRows = lastRow - CALL_LOG_DATA_START_ROW + 1;
    const values = sheet.getRange(CALL_LOG_DATA_START_ROW, 1, numRows, 4).getValues();

    values.forEach((row, index) => {
      const rowNumber = CALL_LOG_DATA_START_ROW + index;
      const sentValue = String(row[CALL_LOG_COLUMN.SENT] || '').trim();

      // Skip if already sent
      if (sentValue) return;

      // Parse the entry
      const entry = parseCallLogRow_(row, rowNumber);
      if (!entry) return;

      // Skip if not an absence entry
      if (!entry.isCallout && !entry.isFmla && !entry.isNoShow) return;

      // Skip if no time recorded
      if (!entry.time) return;

      // Parse HH:MM and create a time for today
      const [hours, mins] = entry.time.split(':').map(Number);
      const entryTime = new Date(now);
      entryTime.setHours(hours, mins, 0, 0);

      // Check if entry is within the 30-minute window
      if (entryTime < windowStart || entryTime > now) return;

      // Send notification and mark as auto-sent
      try {
        sendAbsenceNotification_(CALL_LOG_SHEET_NAME, rowNumber, true);
        console.log(`callLog: Auto-sent notification for row ${rowNumber}`);
      } catch (error) {
        console.error(`callLog: Failed to auto-send row ${rowNumber} — ${error.message}`);
      }
    });
  } catch (error) {
    console.error(`callLog: autoSendAbsenceNotifications failed — ${error.message}`);
  }
}
