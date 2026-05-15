/**
 * callLog.js — Absence log sheet management for COMET.
 * VERSION: 0.3.3
 *
 * This file owns all server-side logic for the Absence Log feature:
 *   - Auto-creating a Call Log sheet for each new week
 *   - Reading absence entries for a given date
 *   - Writing new absence entries
 *   - Triggering absence notification emails (construction/sending lives in notifier.js)
 *
 * CALL LOG SHEET NAMING:
 *   One sheet per fiscal week, named "Call Log MM-DD-YYYY" where the date is
 *   the Monday of that week. Example: "Call Log 04-21-2026".
 *   Sheets are created automatically on the first entry of each new week.
 *   Tab color: green (#57BB8A) to distinguish from infraction and config sheets.
 *
 * CALL LOG COLUMN LAYOUT (from config.js CALL_LOG_COLUMN, 0-indexed):
 *   A (0)  — Employee Name
 *   B (1)  — Employee ID
 *   C (2)  — (reserved)
 *   D (3)  — Is Callout       (checkbox)
 *   E (4)  — (reserved)
 *   F (5)  — Is FMLA          (checkbox)
 *   G (6)  — Is No Show       (checkbox)
 *   H (7)  — Department
 *   I (8)  — Time Called      (HH:MM)
 *   J (9)  — Manager          (who took the call)
 *   K (10) — Scheduled Shift
 *   L-M    — (reserved)
 *   N (13) — Comment
 *   O (14) — Date          (date value)
 *
 * EMAIL NOTIFICATION:
 *   This file triggers absence notification emails via sendAbsenceEmail_() and
 *   resolveAbsenceRecipients_() in notifier.js. Recipients come from the
 *   department's "mailing" field in the Settings sheet. All email construction
 *   and sending lives in notifier.js.
 */


// ---------------------------------------------------------------------------
// Read Entries
// ---------------------------------------------------------------------------

/**
 * Returns all absence entries from the Call Log sheet for the given date.
 *
 * Reads every row in the sheet for the week containing dateString and filters
 * to only rows whose DATE column matches the target date. Returns an empty
 * array if no sheet exists for that week yet.
 *
 * @param {string} dateString — ISO date string (YYYY-MM-DD).
 * @returns {Array<CallLogEntry>}
 */
function getCallLogEntriesForDate_(dateString) {
  const date = new Date(dateString + 'T00:00:00');
  const sheetName = getCallLogSheetName_(date);
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(sheetName);

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < CALL_LOG_DATA_START_ROW) return []; // config.js

  const numRows = lastRow - CALL_LOG_DATA_START_ROW + 1;
  const totalCols = CALL_LOG_COLUMN.DATE + 1; // config.js — read through DATE column
  const data = sheet.getRange(CALL_LOG_DATA_START_ROW, 1, numRows, totalCols).getValues();

  const timeZone = Session.getScriptTimeZone();
  const targetDateKey = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');

  return data
    .map((row, offset) => rowToCallLogEntry_(row, sheetName, CALL_LOG_DATA_START_ROW + offset))
    .filter(entry => {
      if (!entry.dateRaw) return false;
      const entryDate = coerceCallLogDate_(entry.dateRaw);
      if (!entryDate) return false;
      return Utilities.formatDate(entryDate, timeZone, 'yyyy-MM-dd') === targetDateKey;
    });
}


// ---------------------------------------------------------------------------
// Write Entry
// ---------------------------------------------------------------------------

/**
 * Appends a new absence entry to the Call Log sheet for today's week.
 * Creates the sheet if it does not exist yet.
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
  const today = new Date();
  const sheet = getOrCreateCallLogSheet_(today);

  // Build the 15-column row matching the Call Log layout
  const row = new Array(CALL_LOG_COLUMN.DATE + 1).fill('');
  row[CALL_LOG_COLUMN.NAME]             = data.name            || '';
  row[CALL_LOG_COLUMN.EMPLOYEE_ID]      = data.employeeId      || '';
  row[CALL_LOG_COLUMN.IS_CALLOUT]       = !!data.isCallout;
  row[CALL_LOG_COLUMN.IS_FMLA]          = !!data.isFmla;
  row[CALL_LOG_COLUMN.IS_NOSHOW]        = !!data.isNoShow;
  row[CALL_LOG_COLUMN.DEPARTMENT]       = data.department      || '';
  row[CALL_LOG_COLUMN.TIME]             = data.time            || '';
  row[CALL_LOG_COLUMN.MANAGER]          = data.manager         || '';
  row[CALL_LOG_COLUMN.SCHEDULED_SHIFT]  = data.scheduledShift  || '';
  row[CALL_LOG_COLUMN.COMMENT]          = data.comment         || '';
  row[CALL_LOG_COLUMN.DATE]             = today;

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();

  return { sheetName: sheet.getName(), rowNumber: newRow };
}


// ---------------------------------------------------------------------------
// Email Notification
// ---------------------------------------------------------------------------

/**
 * Sends an absence notification email for the entry at the given row.
 *
 * Reads the row from the sheet, determines the correct recipients from
 * MAILING_LIST (config.js), and sends a plain-text email via GmailApp.
 *
 * When sendEmails is false in the COMET Config sheet, logs the email content
 * instead of sending.
 *
 * @param {string} sheetName  — The Call Log sheet tab name.
 * @param {number} rowNumber  — 1-indexed row number to send notification for.
 * @returns {{ sent: boolean, recipients: string[] }}
 */
function sendAbsenceNotification_(sheetName, rowNumber, isAutoSend = false) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(sheetName);

  if (!sheet) throw new Error(`Call Log sheet "${sheetName}" not found.`);

  const totalCols = CALL_LOG_COLUMN.SENT + 1;
  const row = sheet.getRange(rowNumber, 1, 1, totalCols).getValues()[0];
  const entry = rowToCallLogEntry_(row, sheetName, rowNumber);

  // Check if already sent
  const sentCell = row[CALL_LOG_COLUMN.SENT];
  if (sentCell && String(sentCell).trim().length > 0) {
    console.log(`callLog: Email already sent for ${entry.name} (${entry.department}) — marked as "${sentCell}"`);
    return { sent: false, alreadySent: true, recipients: [] };
  }

  // Check if email sending is enabled in UI settings
  const config = readCometConfig_(); // setup.js
  const shouldSendEmails = !!config.sendEmails;

  if (!shouldSendEmails) {
    console.log(`callLog: Email sending is disabled in settings for ${entry.name} (${entry.department})`);
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

  console.log(`callLog: Sent email for ${entry.name} (${entry.department}) to ${recipients.join(', ')} — marked as "${sentLabel}"`);
  return { sent: true, recipients, sentLabel };
}



// ---------------------------------------------------------------------------
// Sheet Management
// ---------------------------------------------------------------------------

/**
 * Returns the Call Log sheet for the week containing the given date,
 * creating it (with headers and formatting) if it does not exist yet.
 *
 * @param {Date} date — Any date within the target week.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateCallLogSheet_(date) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getCallLogSheetName_(date);
  const existing = workbook.getSheetByName(sheetName);
  if (existing) return existing;

  const sheet = workbook.insertSheet(sheetName);
  writeCallLogHeader_(sheet, date);
  return sheet;
}

/**
 * Writes the header row and applies formatting to a new Call Log sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Date} date — The Monday of the week (used in the title row).
 */
function writeCallLogHeader_(sheet, date) {
  const timeZone = Session.getScriptTimeZone();
  const monday = getMondayOfWeek_(date);
  const sunday = new Date(monday);
  sunday.setDate(sunday.getDate() + 6);

  const weekLabel =
    Utilities.formatDate(monday, timeZone, 'MMM d') + ' – ' +
    Utilities.formatDate(sunday, timeZone, 'MMM d, yyyy');

  // Row 1 — merged week title
  const totalCols = CALL_LOG_COLUMN.SENT + 1;
  sheet.getRange(1, 1, 1, totalCols)
    .merge()
    .setValue('Call Log — Week of ' + weekLabel)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#005DAA')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');

  // Row 2 — column headers
  const headers = new Array(totalCols).fill('');
  headers[CALL_LOG_COLUMN.NAME]             = 'Employee Name';
  headers[CALL_LOG_COLUMN.EMPLOYEE_ID]      = 'Employee ID';
  headers[CALL_LOG_COLUMN.IS_CALLOUT]       = 'Call-Out';
  headers[CALL_LOG_COLUMN.IS_FMLA]          = 'FMLA';
  headers[CALL_LOG_COLUMN.IS_NOSHOW]        = 'No Show';
  headers[CALL_LOG_COLUMN.DEPARTMENT]       = 'Department';
  headers[CALL_LOG_COLUMN.TIME]             = 'Time Called';
  headers[CALL_LOG_COLUMN.MANAGER]          = 'Manager';
  headers[CALL_LOG_COLUMN.SCHEDULED_SHIFT]  = 'Scheduled Shift';
  headers[CALL_LOG_COLUMN.COMMENT]          = 'Comment';
  headers[CALL_LOG_COLUMN.DATE]             = 'Date';
  headers[CALL_LOG_COLUMN.SENT]             = 'Sent';

  sheet.getRange(2, 1, 1, totalCols)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#57BB8A')
    .setFontColor('#FFFFFF');

  sheet.setFrozenRows(2);
  sheet.setTabColor('#57BB8A');

  // Column widths for data columns
  sheet.setColumnWidth(CALL_LOG_COLUMN.NAME + 1,             180);
  sheet.setColumnWidth(CALL_LOG_COLUMN.EMPLOYEE_ID + 1,      110);
  sheet.setColumnWidth(CALL_LOG_COLUMN.DEPARTMENT + 1,       150);
  sheet.setColumnWidth(CALL_LOG_COLUMN.TIME + 1,             100);
  sheet.setColumnWidth(CALL_LOG_COLUMN.MANAGER + 1,          150);
  sheet.setColumnWidth(CALL_LOG_COLUMN.SCHEDULED_SHIFT + 1,  120);
  sheet.setColumnWidth(CALL_LOG_COLUMN.COMMENT + 1,          250);
  sheet.setColumnWidth(CALL_LOG_COLUMN.DATE + 1,             110);

  // Compress reserved columns (hide visual gaps)
  sheet.setColumnWidth(3,  1); // Column C (2-indexed as 3)
  sheet.setColumnWidth(5,  1); // Column E (4-indexed as 5)
  sheet.setColumnWidth(10, 1); // Column J (9-indexed as 10)
  sheet.setColumnWidth(11, 1); // Column K (10-indexed as 11)
  sheet.setColumnWidth(12, 1); // Column L (11-indexed as 12)
  sheet.setColumnWidth(13, 1); // Column M (12-indexed as 13)
}

/**
 * Returns the canonical sheet name for the Call Log containing the given date.
 * Format: "Call Log MM-DD-YYYY" where the date is the Monday of the week.
 *
 * @param {Date} date
 * @returns {string} e.g. "Call Log 04-21-2026"
 */
function getCallLogSheetName_(date) {
  const monday = getMondayOfWeek_(date);
  const timeZone = Session.getScriptTimeZone();
  return 'Call Log ' + Utilities.formatDate(monday, timeZone, 'MM-dd-yyyy');
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
// Row Parsing
// ---------------------------------------------------------------------------

/**
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
 *   dateRaw:          any,
 *   sheetName:        string,
 *   rowNumber:        number,
 * }} CallLogEntry
 */

/**
 * Converts a raw getValues() row array into a CallLogEntry object.
 *
 * @param {any[]}  row
 * @param {string} sheetName
 * @param {number} rowNumber — 1-indexed sheet row number
 * @returns {CallLogEntry}
 */
function rowToCallLogEntry_(row, sheetName, rowNumber) {
  return {
    name:             String(row[CALL_LOG_COLUMN.NAME]             || '').trim(),
    employeeId:       String(row[CALL_LOG_COLUMN.EMPLOYEE_ID]      || '').trim(),
    isCallout:        coerceBool_(row[CALL_LOG_COLUMN.IS_CALLOUT]),
    isFmla:           coerceBool_(row[CALL_LOG_COLUMN.IS_FMLA]),
    isNoShow:         coerceBool_(row[CALL_LOG_COLUMN.IS_NOSHOW]),
    department:       String(row[CALL_LOG_COLUMN.DEPARTMENT]       || '').trim(),
    time:             formatCallLogTime_(row[CALL_LOG_COLUMN.TIME]),
    manager:          String(row[CALL_LOG_COLUMN.MANAGER]          || '').trim(),
    scheduledShift:   String(row[CALL_LOG_COLUMN.SCHEDULED_SHIFT]  || '').trim(),
    comment:          String(row[CALL_LOG_COLUMN.COMMENT]          || '').trim(),
    dateRaw:          row[CALL_LOG_COLUMN.DATE],
    sheetName,
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
    return new Date((cellValue - SHEETS_EPOCH_OFFSET) * 86400000); // config.js
  }
  const parsed = new Date(cellValue);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Normalizes a checkbox cell value to a boolean.
 *
 * @param {any} cellValue
 * @returns {boolean}
 */
function coerceBool_(cellValue) {
  if (typeof cellValue === 'boolean') return cellValue;
  if (typeof cellValue === 'string')  return cellValue.toUpperCase() === 'TRUE';
  if (typeof cellValue === 'number')  return cellValue !== 0;
  return false;
}

/**
 * Formats a GAS time value (decimal fraction of a day) to "HH:MM", or returns
 * the value as-is if it is already a string.
 *
 * @param {any} cellValue
 * @returns {string}
 */
function formatCallLogTime_(cellValue) {
  if (!cellValue && cellValue !== 0) return '';
  if (typeof cellValue === 'string') return cellValue;
  if (cellValue instanceof Date) {
    const h = String(cellValue.getHours()).padStart(2, '0');
    const m = String(cellValue.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  // GAS time serial (fraction of a day)
  if (typeof cellValue === 'number') {
    const totalMinutes = Math.round(cellValue * 24 * 60);
    const h = String(Math.floor(totalMinutes / 60) % 24).padStart(2, '0');
    const m = String(totalMinutes % 60).padStart(2, '0');
    return `${h}:${m}`;
  }
  return String(cellValue);
}


// ---------------------------------------------------------------------------
// Auto-Send Notifications (Rolling 30-Minute Window)
// ---------------------------------------------------------------------------

/**
 * Auto-sends absence notifications for entries logged in the last 30 minutes.
 * Runs on a time-based trigger every 30 minutes. Marks sent entries with
 * "auto sent at HH:MM" in the SENT column.
 *
 * Called by: Time-based trigger (via setup.js)
 * Behavior: Finds all Call Log sheets for the current week, searches for
 * unsent entries within the rolling 30-minute window, and sends notifications.
 */
function autoSendAbsenceNotifications_() {
  try {
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    const windowStart = new Date(now.getTime() - 30 * 60 * 1000); // 30 min ago
    const timeZone = Session.getScriptTimeZone();

    // Get all sheets and filter to Call Log sheets (named "Call Log MM-DD-YYYY")
    const sheets = workbook.getSheets().filter(s => s.getName().match(/^Call Log/));

    if (sheets.length === 0) {
      console.log('callLog: No Call Log sheets found for autosend.');
      return;
    }

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const lastRow = sheet.getLastRow();

      // Data starts at row 3 (row 1 = title, row 2 = headers)
      if (lastRow < 3) return;

      const totalCols = CALL_LOG_COLUMN.SENT + 1;
      const values = sheet.getRange(3, 1, lastRow - 2, totalCols).getValues();

      values.forEach((row, index) => {
        const rowNumber = 3 + index;
        const sentValue = String(row[CALL_LOG_COLUMN.SENT] || '').trim();

        // Skip if already sent
        if (sentValue) return;

        // Check if it's an absence entry
        const isCallout = coerceBool_(row[CALL_LOG_COLUMN.IS_CALLOUT]);
        const isFmla = coerceBool_(row[CALL_LOG_COLUMN.IS_FMLA]);
        const isNoShow = coerceBool_(row[CALL_LOG_COLUMN.IS_NOSHOW]);
        if (!isCallout && !isFmla && !isNoShow) return;

        // Parse the time called
        const timeValue = row[CALL_LOG_COLUMN.TIME];
        const timeStr = formatCallLogTime_(timeValue);
        if (!timeStr) return;

        // Parse HH:MM and create a time for today
        const [hours, mins] = timeStr.split(':').map(Number);
        const entryTime = new Date(now);
        entryTime.setHours(hours, mins, 0, 0);

        // Check if entry is within the 30-minute window
        if (entryTime < windowStart || entryTime > now) return;

        // Send notification and mark as auto-sent
        try {
          sendAbsenceNotification_(sheetName, rowNumber, true);
          console.log(`callLog: Auto-sent notification for row ${rowNumber} in "${sheetName}"`);
        } catch (error) {
          console.error(`callLog: Failed to auto-send row ${rowNumber} in "${sheetName}" — ${error.message}`);
        }
      });
    });
  } catch (error) {
    console.error(`callLog: autoSendAbsenceNotifications failed — ${error.message}`);
  }
}
