/**
 * dataIngestion.js — Reads the call log sheet and produces a filtered list of absence records.
 * VERSION: 0.2.1
 *
 * This file is responsible for turning raw spreadsheet data into structured JavaScript
 * objects that the rest of the notifier can work with. It owns three sequential steps:
 *
 *   1. SHEET READ:   Fetch all data rows from the current week's call log sheet in a
 *                    single getValues() call. One sheet read per trigger run — no loops
 *                    that call getRange() or getValue() row by row.
 *
 *   2. ROW MAPPING:  Convert each raw array row into a named AbsenceRecord object.
 *                    This decouples all downstream code from raw column indices so that
 *                    a column rearrangement only requires updating CALL_LOG_COLUMNS in
 *                    config.js, not every function that touches the data.
 *
 *   3. FILTERING:    Discard any record that should not generate a notification:
 *                    rows with no absence flag, rows whose call time falls outside the
 *                    current window, and rows whose call date does not match today.
 *
 * DATA SHAPE — AbsenceRecord:
 *   {
 *     rowNumber:        number  — 1-based row in the sheet (for logging / debugging)
 *     employeeName:     string  — Employee full name (column B)
 *     employeeId:       string  — Employee ID (column C)
 *     isAbsence:        boolean — TRUE if at least one of IS_CALLOUT, IS_FMLA, IS_NOSHOW is checked
 *     absenceReason:    string  — "Call-Out" | "FMLA" | "No-Show"
 *     department:       string  — Department name from column D (used to route emails)
 *     calledAt:         Date    — The resolved Date object for when the employee called
 *     employeeComment:  string  — Optional comment from column L
 *   }
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Reads the active call log sheet and returns a filtered array of AbsenceRecord
 * objects for rows that fall within the given time window.
 *
 * This is the only function in this file that is called from outside (by
 * sendAbsenceDigest in absenceNotifier.js). All other functions are private helpers
 * that support this one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The call log sheet for this period/week.
 * @param {{ start: Date, end: Date }} window — The time window to filter rows by.
 * @param {string} timeZone — The script's time zone (from Session.getScriptTimeZone()).
 * @returns {AbsenceRecord[]} Filtered, structured records ready for email grouping.
 */
function loadAbsenceRecordsInWindow_(sheet, window, timeZone) {
  const rawRows = readRawRowsFromSheet_(sheet);

  if (rawRows.length === 0) {
    console.log('dataIngestion: No data rows found in sheet.');
    return [];
  }

  const allRecords = mapRowsToAbsenceRecords_(rawRows, CALL_LOG_DATA_START_ROW);
  const recordsInWindow = filterRecordsToWindow_(allRecords, window, timeZone);

  console.log(
    `dataIngestion: ${rawRows.length} rows read, ` +
    `${allRecords.filter(r => r.isAbsence).length} absences found, ` +
    `${recordsInWindow.length} in window.`
  );
  return recordsInWindow;
}


// ---------------------------------------------------------------------------
// Step 1: Sheet Read
// ---------------------------------------------------------------------------

/**
 * Performs a single getValues() call to read all data rows from the sheet.
 *
 * Reading all rows in one call is substantially faster than reading row by row
 * inside a loop. The GAS documentation recommends batching all reads into a
 * single range read, which this function enforces.
 *
 * The generated call log sheet has:
 *   Row 1 — Fiscal period header (merged cells, not data)
 *   Row 2 — Column headers (not data)
 *   Row 3+ — One entry per absence
 *
 * Returns an empty array if the sheet has no data rows below the header rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The call log sheet.
 * @returns {Array[]} A 2D array of raw cell values; each inner array is one row.
 */
function readRawRowsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < CALL_LOG_DATA_START_ROW) {
    return []; // Sheet has headers but no data rows yet
  }

  const numberOfDataRows = lastRow - CALL_LOG_DATA_START_ROW + 1;

  // Read all rows from the data start row through the last row, covering
  // all columns A–L as defined by CALL_LOG_COLUMNS.TOTAL_COLUMNS_TO_READ.
  return sheet
    .getRange(CALL_LOG_DATA_START_ROW, 1, numberOfDataRows, CALL_LOG_COLUMNS.TOTAL_COLUMNS_TO_READ)
    .getValues();
}


// ---------------------------------------------------------------------------
// Step 2: Row Mapping
// ---------------------------------------------------------------------------

/**
 * Converts a 2D array of raw row values into an array of AbsenceRecord objects.
 *
 * Each raw row is a flat array of cell values indexed by column position. This
 * function lifts those values into named properties using CALL_LOG_COLUMNS from
 * config.js. After this step, no other function in the codebase needs to know
 * which array index corresponds to which column — they operate on field names.
 *
 * Note: this function maps ALL rows, including rows that are not absences.
 * The filtering step (filterRecordsToWindow_) is responsible for discarding
 * non-absence rows. Keeping mapping and filtering separate makes each step
 * simpler to reason about and to test.
 *
 * @param {Array[]} rawRows      — 2D array from readRawRowsFromSheet_().
 * @param {number}  dataStartRow — The 1-based sheet row where data begins.
 *   Used to populate the rowNumber field for debugging and logging.
 * @returns {AbsenceRecord[]} One record per input row (not yet filtered).
 */
function mapRowsToAbsenceRecords_(rawRows, dataStartRow) {
  const columns = CALL_LOG_COLUMNS;

  return rawRows.map((row, arrayIndex) => {
    const isCallout = coerceToBool_(row[columns.IS_CALLOUT]);
    const isFmla    = coerceToBool_(row[columns.IS_FMLA]);
    const isNoShow  = coerceToBool_(row[columns.IS_NOSHOW]);

    // Determine the absence reason from the first flag that is true.
    // Priority order: Call-Out → FMLA → No-Show. If none are true, the reason
    // is irrelevant because the row will be filtered out, but we still set a
    // safe default so no field is ever undefined.
    let absenceReason = 'Unknown';
    if (isCallout)     absenceReason = 'Call-Out';
    else if (isFmla)   absenceReason = 'FMLA';
    else if (isNoShow) absenceReason = 'No-Show';

    return {
      rowNumber:        dataStartRow + arrayIndex, // 1-based sheet row number
      employeeName:     String(row[columns.NAME]             || 'Unknown').trim(),
      employeeId:       String(row[columns.EMPLOYEE_ID]      || 'Unknown').trim(),
      isAbsence:        isCallout || isFmla || isNoShow,
      absenceReason:    absenceReason,
      department:       String(row[columns.DEPT]             || '').trim(),
      employeeComment:  String(row[columns.COMMENT]          || '').trim(),
      scheduledShift:   String(row[columns.SCHEDULED_SHIFT]  || '').trim(),

      // The date of the absence lives in column A. It is used by the filter
      // step to verify the row belongs to the same calendar day as the window.
      dateRaw: row[columns.DATE],

      // The time called lives in column H. parseTimeToMilliseconds_() in
      // timeWindow.js handles all format variations (Date object, fractional
      // day, serial number, string).
      timeRaw: row[columns.TIME_CALLED],

      // calledAt is populated during filtering once time parsing succeeds.
      // It is undefined at this stage.
      calledAt: undefined,
    };
  });
}


// ---------------------------------------------------------------------------
// Step 3: Filtering
// ---------------------------------------------------------------------------

/**
 * Filters an array of AbsenceRecord objects down to only those that should
 * generate a notification in the current trigger run.
 *
 * A record passes the filter only if ALL of the following are true:
 *   1. The row has at least one absence flag checked (isAbsence === true).
 *   2. The call time can be parsed into a valid millisecond value.
 *   3. The call time falls within the current window: (start, end].
 *   4. The call date in column A resolves to a real calendar date.
 *   5. The call date's local calendar day matches the window's local calendar day.
 *      (This prevents time-only false positives from triggering on the wrong day.)
 *
 * Records that pass are returned with their calledAt field set to a resolved Date.
 * Records that fail any condition are silently discarded; skip reasons are logged
 * to the console for debugging.
 *
 * @param {AbsenceRecord[]} records  — All records from mapRowsToAbsenceRecords_().
 * @param {{ start: Date, end: Date }} window — The time window to match against.
 * @param {string} timeZone — The script's time zone string.
 * @returns {AbsenceRecord[]} Only the records that belong in this digest.
 */
function filterRecordsToWindow_(records, window, timeZone) {
  const windowStartMilliseconds = window.start.getTime();
  const windowEndMilliseconds   = window.end.getTime();
  const windowDayKey            = getLocalDateKey_(window.end, timeZone);

  return records.filter(record => {
    // Condition 1: Must be an absence row
    if (!record.isAbsence) return false;

    // Conditions 2 & 3: Time must be parseable and within the window
    const callTimeMilliseconds = parseTimeToMilliseconds_(record.timeRaw, window);
    if (callTimeMilliseconds == null) {
      console.log(`dataIngestion: Row ${record.rowNumber} skipped — time value could not be parsed.`);
      return false;
    }
    if (callTimeMilliseconds <= windowStartMilliseconds || callTimeMilliseconds > windowEndMilliseconds) {
      return false; // Row is real data but was logged in a different window; not an error
    }

    // Conditions 4 & 5: Date must be a real calendar date matching the window's day
    const callDate = coerceToCalendarDate_(record.dateRaw, timeZone);
    if (!callDate) {
      console.log(`dataIngestion: Row ${record.rowNumber} skipped — absence date is missing or a time-only value.`);
      return false;
    }
    const callDayKey = getLocalDateKey_(callDate, timeZone);
    if (callDayKey !== windowDayKey) {
      console.log(
        `dataIngestion: Row ${record.rowNumber} skipped — ` +
        `absence date ${callDayKey} does not match window day ${windowDayKey}.`
      );
      return false;
    }

    // All conditions met — populate calledAt and include this record
    record.calledAt = new Date(callTimeMilliseconds);
    return true;
  });
}


// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

/**
 * Coerces common cell value types into a boolean.
 *
 * Checkbox cells in Google Sheets deliver TRUE/FALSE as JavaScript booleans.
 * However, cells that were manually typed or pasted may arrive as the strings
 * "TRUE"/"FALSE" or the numbers 1/0. This function normalizes all of those
 * forms so that the mapping step does not need conditional type logic.
 *
 * Returns null for any value that does not map cleanly to true or false, which
 * allows callers to distinguish between "definitely false" and "unreadable cell".
 *
 * @param {boolean|string|number|*} cellValue — The raw checkbox cell value.
 * @returns {boolean|null} true, false, or null if the value is unrecognized.
 */
function coerceToBool_(cellValue) {
  if (cellValue === true  || cellValue === false) return cellValue;
  if (typeof cellValue === 'string') {
    const lowered = cellValue.trim().toLowerCase();
    if (lowered === 'true')  return true;
    if (lowered === 'false') return false;
  }
  if (typeof cellValue === 'number') {
    if (cellValue === 1) return true;
    if (cellValue === 0) return false;
  }
  return null;
}
