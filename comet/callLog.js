/**
 * callLog.js — Absence log sheet management and email notifications for COMET.
 * VERSION: 0.2.4
 *
 * This file owns all server-side logic for the Absence Log feature:
 *   - Auto-creating a Call Log sheet for each new week
 *   - Reading absence entries for a given date
 *   - Writing new absence entries
 *   - Sending notification emails to the correct department recipients
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
 *   Recipients are determined by the entry's department using MAILING_LIST
 *   from config.js. Departments not in MAILING_LIST fall back to FALLBACK_EMAIL.
 *   Emails are sent via GmailApp.sendEmail() and are always plain text.
 *   DRY_RUN = true in config.js suppresses sends and logs instead.
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
 * When DRY_RUN is true (config.js), logs the email content instead of sending.
 *
 * @param {string} sheetName  — The Call Log sheet tab name.
 * @param {number} rowNumber  — 1-indexed row number to send notification for.
 * @returns {{ sent: boolean, recipients: string[] }}
 */
function sendAbsenceNotification_(sheetName, rowNumber) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(sheetName);

  if (!sheet) throw new Error(`Call Log sheet "${sheetName}" not found.`);

  const totalCols = CALL_LOG_COLUMN.DATE + 1;
  const row = sheet.getRange(rowNumber, 1, 1, totalCols).getValues()[0];
  const entry = rowToCallLogEntry_(row, sheetName, rowNumber);

  const recipients = resolveRecipients_(entry.department);
  const { subject, body } = buildEmailContent_(entry);

  if (DRY_RUN) { // config.js
    console.log(`callLog: DRY RUN — would send to ${recipients.join(', ')}\nSubject: ${subject}\n${body}`);
    return { sent: false, recipients };
  }

  GmailApp.sendEmail(recipients.join(','), subject, body);
  console.log(`callLog: Sent absence notification for ${entry.name} to ${recipients.join(', ')}`);
  return { sent: true, recipients };
}

/**
 * Builds the subject and plain-text body for an absence notification email.
 *
 * @param {CallLogEntry} entry
 * @returns {{ subject: string, body: string }}
 */
function buildEmailContent_(entry) {
  const timeZone = Session.getScriptTimeZone();
  const dateStr = entry.dateRaw
    ? Utilities.formatDate(coerceCallLogDate_(entry.dateRaw) || new Date(), timeZone, 'MMMM d, yyyy')
    : 'Unknown date';

  const types = [];
  if (entry.isCallout) types.push('Call-Out');
  if (entry.isFmla)    types.push('FMLA');
  if (entry.isNoShow)  types.push('No Show');
  const typeLabel = types.length > 0 ? types.join(', ') : 'Absence';

  const subject = `[COMET] ${typeLabel} — ${entry.name} (${entry.department}) — ${dateStr}`;

  const lines = [
    'COMET Absence Notification',
    '──────────────────────────',
    `Employee:        ${entry.name}`,
    `ID:              ${entry.employeeId || '—'}`,
    `Department:      ${entry.department || '—'}`,
    `Date:            ${dateStr}`,
    `Time Called:     ${entry.time || '—'}`,
    `Type:            ${typeLabel}`,
  ];

  if (entry.manager) {
    lines.push(`Manager:         ${entry.manager}`);
  }

  if (entry.scheduledShift) {
    lines.push(`Scheduled Shift: ${entry.scheduledShift}`);
  }

  if (entry.comment) {
    lines.push(`Comment:         ${entry.comment}`);
  }

  lines.push('', '──────────────────────────');
  lines.push('This notification was generated automatically by COMET.');

  return { subject, body: lines.join('\n') };
}

/**
 * Resolves the list of email recipients for the given department.
 * Falls back to FALLBACK_EMAIL if the department is not in MAILING_LIST.
 *
 * @param {string} department
 * @returns {string[]}
 */
function resolveRecipients_(department) {
  const list = MAILING_LIST[department]; // config.js
  if (list && list.length > 0) return list;

  console.warn(`callLog: No mailing list entry for department "${department}" — using fallback.`);
  return FALLBACK_EMAIL; // config.js
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
  const totalCols = CALL_LOG_COLUMN.DATE + 1;
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
