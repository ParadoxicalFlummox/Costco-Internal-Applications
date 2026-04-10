/*
VERSION: 0.2.1
LICENSE: MIT (Generic Logic Engine)
PURPOSE: A time-windowed digest notification system.
         Reads call log entries from a Google Sheets workbook, filters them
         to the most recently completed N-minute window, and emails one
         summary digest per department to that department's configured manager(s).
*/

/**
 * digestEngine.js — Main entry point and pipeline orchestrator.
 * VERSION: 0.2.1
 *
 * This file contains a single public function: sendAbsenceDigest(). Its only
 * job is to call the other files in the correct order and pass data between them.
 * It contains no time math, no sheet reads, no email formatting, and no
 * recipient lookups. Each of those concerns lives in its own file:
 *
 *   config.js          — All configurable values (columns, window length, mailing list)
 *   timeWindow.js      — Window boundary calculation and time value parsing
 *   sheetUtils.js      — Fiscal calendar math and sheet name resolution ("P# W#")
 *   dataIngestion.js   — Sheet reading, row mapping, and window filtering
 *   notifier.js        — Recipient resolution, email body building, and sending
 *   sheetGenerator.js  — Generates new call log sheets and the Absence Config sheet
 *   autofill.js        — onEdit autofill of Employee ID and Department from roster
 *   ui.js              — onOpen menu, menu handlers, and toast feedback
 *
 * HOW TO SET UP THE DIGEST TRIGGER:
 *   1. In the Apps Script editor, go to Triggers → Add Trigger.
 *   2. Choose function: sendAbsenceDigest
 *   3. Event source: Time-driven
 *   4. Type: Minutes timer — every 15 minutes (or match WINDOW_MINUTES in config.js)
 *
 * HOW TO SET UP THE AUTOFILL TRIGGER:
 *   1. In the Apps Script editor, go to Triggers → Add Trigger.
 *   2. Choose function: onEdit
 *   3. Event source: From spreadsheet → On edit
 *   (Must be an installable trigger — simple triggers cannot write across ranges.)
 *
 * HOW TO DRY-RUN (send no emails):
 *   In notifier.js, comment out the GmailApp.sendEmail() call inside
 *   sendDepartmentDigests_(). The function will still log what it would have sent.
 */


// ---------------------------------------------------------------------------
// Entry Point
// ---------------------------------------------------------------------------

/**
 * Orchestrates a single digest run for the most recently completed time window.
 *
 * Called automatically by the Apps Script time-driven trigger. The sequence is:
 *
 *   1. Determine the current call log sheet title (fiscal "P# W#" or fallback).
 *   2. Locate that sheet in the workbook.
 *   3. Compute the previous N-minute window (e.g., 09:00–09:15 if now is 09:17).
 *   4. Read and filter the sheet for rows that fall within that window.
 *   5. Group the filtered rows by department and send one email per department.
 *
 * Early exits (no email sent):
 *   - The current period/week sheet does not exist in the workbook.
 *   - The sheet has no data rows.
 *   - No rows in the sheet fall within the current window.
 *
 * Errors are caught and logged rather than thrown. This prevents a bad row from
 * halting the entire trigger run and causing GAS to retry or disable the trigger.
 */
function sendAbsenceDigest() {
  try {
    // Step 1: Determine the current sheet title and locate the sheet.
    // getActiveCallLogSheetName_() in sheetUtils.js calculates "P# W#" if a FY
    // start date is configured, or "Week Ending MM/DD/YY" as a fallback.
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTitle = getActiveCallLogSheetName_();
    const currentSheet = workbook.getSheetByName(sheetTitle);

    if (!currentSheet) {
      console.warn(
        `sendAbsenceDigest: Sheet "${sheetTitle}" not found. ` +
        `Use "Call Log Admin → Generate New Week Sheet" to create it.`
      );
      return;
    }

    // Step 2: Compute the previous window and fetch the script time zone.
    // Session.getScriptTimeZone() is called once here and passed through all
    // helpers to avoid redundant GAS API calls inside each function.
    const timeWindow = getPreviousWindow_(WINDOW_MINUTES);
    const timeZone = Session.getScriptTimeZone();

    console.log(
      `sendAbsenceDigest: Scanning window ` +
      `${Utilities.formatDate(timeWindow.start, timeZone, 'h:mm a')} – ` +
      `${Utilities.formatDate(timeWindow.end, timeZone, 'h:mm a')} ` +
      `on sheet "${sheetTitle}".`
    );

    // Step 3: Read the sheet and filter to absence records within the window
    const absenceRecords = loadAbsenceRecordsInWindow_(currentSheet, timeWindow, timeZone);

    if (absenceRecords.length === 0) {
      console.log('sendAbsenceDigest: No absences found in the current window. No emails sent.');
      return;
    }

    // Step 4: Group by department and send one digest email per department
    sendDepartmentDigests_(absenceRecords, timeWindow, timeZone);

  } catch (error) {
    console.error('sendAbsenceDigest: Unexpected error —', error);
  }
}
