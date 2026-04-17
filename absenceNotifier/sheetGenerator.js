/**
 * sheetGenerator.js — Generates new call log sheets and the Absence Config setup sheet.
 * VERSION: 1.0.1
 *
 * This file owns all sheet creation logic for the absence notifier workbook:
 *
 *   1. CALL LOG SHEET GENERATION: Creates a new, correctly-formatted call log
 *      sheet for the current fiscal period and week (or the current calendar
 *      week if no FY start date is configured). The generated sheet has the
 *      merged fiscal header in row 1, column headers in row 2, checkboxes
 *      pre-populated in the absence flag columns, and frozen header rows so
 *      managers can scroll without losing context.
 *
 *   2. CONFIG SHEET SETUP: Creates the "Absence Config" sheet with a labeled
 *      input cell for the FY start date. This is a one-time setup step that
 *      managers run when first deploying the notifier.
 *
 * The custom menu (onOpen) and all menu handler functions live in ui.js.
 *
 * HOW TO RUN SETUP:
 *   Use "Call Log Admin → Setup Config Sheet (First Run)" from the menu.
 *   Then enter the FY start date in the Absence Config sheet (cell B2).
 *   After that, use "Call Log Admin → Generate New Week Sheet" each period/week.
 */


// ---------------------------------------------------------------------------
// Call Log Sheet Generation
// ---------------------------------------------------------------------------

/**
 * Generates a new blank call log sheet for the current fiscal period and week.
 *
 * This is the primary action managers perform each week. It creates a sheet
 * with the correct title, the fiscal header in row 1, the column headers in
 * row 2, and checkboxes pre-inserted in the Call-Out, FMLA, and No-Show columns
 * for all anticipated data rows. Managers then fill in entries as calls come in.
 *
 * GUARD: If a sheet with the calculated title already exists, the function
 * shows a warning dialog and exits without modifying anything. This prevents
 * accidentally wiping a sheet that already contains entries.
 *
 * Callable from: "Call Log Admin" → "Generate New Week Sheet"
 */
function generateNewCallLogSheet() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTitle = getActiveCallLogSheetName_(); // defined in sheetUtils.js

  // Guard: do not overwrite an existing sheet
  if (workbook.getSheetByName(sheetTitle)) {
    SpreadsheetApp.getUi().alert(
      `Sheet "${sheetTitle}" already exists.\n\n` +
      `If you need to reset it, delete the sheet manually first.`
    );
    return;
  }

  const newSheet = workbook.insertSheet(sheetTitle);

  writeCallLogHeaderRows_(newSheet);
  applyCallLogFormatting_(newSheet);
  insertAbsenceCheckboxes_(newSheet);

  // Move the new sheet to the front so it is visible immediately after creation.
  workbook.setActiveSheet(newSheet);
  workbook.moveActiveSheet(1);

  console.log(`sheetGenerator: Created new call log sheet "${sheetTitle}".`);
}

/**
 * Writes the two header rows that appear at the top of every call log sheet.
 *
 * Row 1 — Fiscal header:
 *   - Cells A1:B1 merged, containing the fiscal year label (e.g. "FY'26")
 *   - Cell C1 left blank as a visual spacer
 *   - Cells D1:L1 merged, containing:
 *       - When FY is configured: "P3 W1  |  April 6 – 12, 2026"
 *       - Fallback (no FY date): "April 6 – 12, 2026"
 *
 * Row 2 — Column headers:
 *   One label per column A through L, matching CALL_LOG_COLUMNS in config.js.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The newly created call log sheet.
 */
function writeCallLogHeaderRows_(sheet) {
  const fyStartDate = readFiscalYearStartDate_(); // defined in sheetUtils.js
  const today = new Date();

  // --- Row 1: Fiscal period header ---

  // The week date range (e.g. "April 6 – 12, 2026") is always shown.
  // When a FY start date is configured, the period/week label is prepended
  // so the full header reads: "P3 W1  |  April 6 – 12, 2026"
  const weekRangeText = calculateWeekDateRangeLabel_(today); // defined in sheetUtils.js

  let headerText;
  if (fyStartDate) {
    const fiscal = calculateFiscalPeriodAndWeek_(fyStartDate, today); // defined in sheetUtils.js
    if (fiscal) {
      headerText = `P${fiscal.period} W${fiscal.week}  \u2502  ${weekRangeText}`;
    } else {
      headerText = weekRangeText; // reference date is before FY start — just show date range
    }
  } else {
    headerText = weekRangeText;
  }

  // Fiscal year label (A1:B1 merged)
  const fyLabel = getCurrentFiscalYearLabel_(); // defined in sheetUtils.js
  sheet.getRange('A1:B1').merge().setValue(fyLabel);

  // Period/week + date range label (D1:L1 merged) — C1 left blank as spacer
  sheet.getRange('D1:L1').merge().setValue(headerText);

  // --- Row 2: Column headers ---
  const columnHeaders = [
    'DATE',
    'EMPLOYEE NAME',
    'EMPLOYEE ID',
    'HOME DEPT',
    'CALL-OUT',
    'FMLA',
    'NO-SHOW',
    'TIME CALLED',
    'SCHEDULED SHIFT',
    'ARRIVAL TIME',
    'MANAGER APP',
    'COMMENT',
    'SEND',
  ];
  sheet.getRange(2, 1, 1, columnHeaders.length).setValues([columnHeaders]);
}

/**
 * Applies visual formatting to the header rows and sets column widths.
 *
 * Formatting applied:
 *   - Row 1: Bold, slightly larger font, light gray background
 *   - Row 2: Bold, white text on dark background (matches AutoScheduler header style)
 *   - Rows 1–2: Frozen so they stay visible while scrolling through entries
 *   - Column widths set so names and comments have enough space to read at a glance
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The newly created call log sheet.
 */
function applyCallLogFormatting_(sheet) {
  // Row 1: fiscal header styling
  sheet.getRange('A1:M1')
    .setBackground('#F5F5F5')
    .setFontWeight('bold')
    .setFontSize(11);

  // Row 2: column header styling — dark background, white text
  sheet.getRange('A2:M2')
    .setBackground('#263238')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Freeze the first two rows so headers stay visible while scrolling
  sheet.setFrozenRows(2);

  // Column widths (in pixels) — tuned for readability without excessive scrolling
  const columnWidths = {
    1: 90,  // A — DATE
    2: 160, // B — EMPLOYEE NAME
    3: 100, // C — EMPLOYEE ID
    4: 120, // D — HOME DEPT
    5: 80,  // E — CALL-OUT  (checkbox)
    6: 65,  // F — FMLA      (checkbox)
    7: 80,  // G — NO-SHOW   (checkbox)
    8: 100, // H — TIME CALLED
    9: 120, // I — SCHEDULED SHIFT
    10: 100, // J — ARRIVAL TIME
    11: 110, // K — MANAGER APP
    12: 240, // L — COMMENT   (wide; comments can be lengthy)
    13: 110, // M — SEND      (shows checkbox, then "Sent HH:MM AM" or "Auto sent HH:MM")
  };
  Object.entries(columnWidths).forEach(([columnNumber, widthPixels]) => {
    sheet.setColumnWidth(Number(columnNumber), widthPixels);
  });

  // Apply alternating row banding to the data rows so entries are easy to track
  // across wide rows. Applied to a large range so future dynamically-added rows
  // are covered automatically without any additional per-row formatting calls.
  //   Odd rows:  white         (#FFFFFF) — clean, unobtrusive base
  //   Even rows: light blue    (#EDF2FA) — subtle, complements the dark header
  const bandingRange = sheet.getRange(CALL_LOG_DATA_START_ROW, 1, 500, CALL_LOG_COLUMNS.TOTAL_COLUMNS_TO_READ);
  const banding = bandingRange.applyRowBanding();
  banding.setHeaderRowColor(null);       // don't override the header row we already styled
  banding.setFirstRowColor('#FFFFFF');   // odd data rows
  banding.setSecondRowColor('#EDF2FA'); // even data rows — very light blue-gray
  banding.setFooterRowColor(null);
}

/**
 * Inserts a single "log entry" checkbox into column A of the first data row.
 *
 * Rather than pre-populating 200 rows, the sheet starts with one ready row.
 * When the manager checks the box in column A, autofill.js stamps today's date
 * into that cell, adds the Call-Out/FMLA/No-Show checkboxes to the same row,
 * and appends a new trigger checkbox on the row below — so the sheet grows
 * one row at a time as entries are logged.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The newly created call log sheet.
 */
function insertAbsenceCheckboxes_(sheet) {
  insertEntryTriggerCheckbox_(sheet, CALL_LOG_DATA_START_ROW);
}

/**
 * Inserts a single checkbox into column A of the given row number.
 *
 * This is the "log entry" trigger cell. Checking it signals autofill.js to
 * stamp today's date into that cell, add absence-type checkboxes to the row,
 * and create the next trigger row.
 *
 * This function is also called by autofill.js when creating the next row after
 * a manager activates an entry, so it lives here as a shared utility.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet.
 * @param {number}                             rowNumber — 1-based row to insert the checkbox into.
 */
function insertEntryTriggerCheckbox_(sheet, rowNumber) {
  const checkboxValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  sheet.getRange(rowNumber, 1) // Column A
    .setDataValidation(checkboxValidation)
    .setValue(false);
}

/**
 * Inserts Call-Out, FMLA, and No-Show checkboxes into the absence-type columns
 * of a single entry row.
 *
 * Called by autofill.js when a manager activates a row (checks column A).
 * At that point the row transitions from "pending trigger" to "active entry"
 * and needs its absence-type checkboxes ready for the manager to fill in.
 *
 * Columns written (1-indexed):
 *   E (5) — Call-Out
 *   F (6) — FMLA
 *   G (7) — No-Show
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet.
 * @param {number}                             rowNumber — 1-based row to add checkboxes to.
 */
function insertAbsenceTypeCheckboxes_(sheet, rowNumber) {
  const checkboxValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  const absenceColumns = [
    CALL_LOG_COLUMNS.IS_CALLOUT + 1, // E (0-indexed 4 → 1-indexed 5)
    CALL_LOG_COLUMNS.IS_FMLA + 1, // F (0-indexed 5 → 1-indexed 6)
    CALL_LOG_COLUMNS.IS_NOSHOW + 1, // G (0-indexed 6 → 1-indexed 7)
  ];

  absenceColumns.forEach(columnNumber => {
    sheet.getRange(rowNumber, columnNumber)
      .setDataValidation(checkboxValidation)
      .setValue(false);
  });
}

/**
 * Inserts the SEND checkbox into column M of a newly activated entry row.
 *
 * The SEND checkbox is the manager's explicit signal to send the notification
 * for this row immediately. When checked, autofill.js fires the email and
 * replaces the checkbox with a "Sent HH:MM AM" timestamp so the row is
 * permanently marked and will not be re-sent by the time-driven trigger.
 *
 * This is called by activateEntryRow_() in autofill.js alongside
 * insertAbsenceTypeCheckboxes_() so all per-row UI controls appear at the
 * same moment the manager opens the row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet.
 * @param {number}                             rowNumber — 1-based row to insert the checkbox into.
 */
function insertNotifyCheckbox_(sheet, rowNumber) {
  const checkboxValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();

  sheet.getRange(rowNumber, CALL_LOG_NOTIFY_COLUMN_NUMBER) // Column M
    .setDataValidation(checkboxValidation)
    .setValue(false);
}


// ---------------------------------------------------------------------------
// Config Sheet Setup
// ---------------------------------------------------------------------------

/**
 * Creates (or re-initializes) the "Absence Config" sheet.
 *
 * This sheet has a single input cell (B2) where a manager enters the fiscal
 * year start date — the Monday of P1 W1 for the current fiscal year. Once
 * this date is set, the sheet name calculation in sheetUtils.js produces
 * "P# W#" labels instead of the "Week Ending" fallback.
 *
 * GUARD: If the config sheet already exists AND cell B2 contains a non-empty,
 * non-placeholder value, setup is skipped to avoid overwriting a configured date.
 * If B2 is empty or contains the instructional placeholder, setup runs and
 * clears/resets the sheet layout.
 *
 * Callable from: "Call Log Admin" → "Setup Config Sheet"
 */
function setupAbsenceConfigSheet() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = workbook.getSheetByName(CONFIG_SHEET_NAME);

  // Guard: if the sheet exists and is already configured, do not overwrite it.
  if (configSheet) {
    const existingValue = configSheet
      .getRange(FISCAL_YEAR_START_CELL)
      .getValue()
      .toString()
      .trim();
    const isAlreadyConfigured = existingValue !== '' &&
      existingValue !== '← enter date here';
    if (isAlreadyConfigured) {
      SpreadsheetApp.getUi().alert(
        `Config sheet "${CONFIG_SHEET_NAME}" is already set up.\n\n` +
        `Current FY start date: ${existingValue}\n\n` +
        `To change the date, edit cell B2 directly.`
      );
      return;
    }
  }

  if (!configSheet) {
    configSheet = workbook.insertSheet(CONFIG_SHEET_NAME);
  } else {
    configSheet.clearContents();
    configSheet.clearFormats();
  }

  writeConfigSheetLayout_(configSheet);

  workbook.setActiveSheet(configSheet);
  console.log(`sheetGenerator: Absence Config sheet created or reset.`);
}

/**
 * Writes the labels and input cell layout to the Absence Config sheet.
 *
 * Layout:
 *   A1 — Bold title: "Absence Notifier — Configuration"
 *   A2 — Label: "Fiscal Year Start Date:"
 *   B2 — Input cell (left blank with a placeholder note; manager fills this in)
 *   A4 — Instructional text explaining what to enter and why
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet — The config sheet to write to.
 */
function writeConfigSheetLayout_(configSheet) {
  // Title row
  configSheet.getRange('A1')
    .setValue('Absence Notifier — Configuration')
    .setFontWeight('bold')
    .setFontSize(13);

  // --- Row 2: Fiscal year start date ---
  configSheet.getRange('A2')
    .setValue('Fiscal Year Start Date:')
    .setFontWeight('bold');

  configSheet.getRange(FISCAL_YEAR_START_CELL) // "B2"
    .setValue('')
    .setNote(
      'Enter the first day of P1 W1 for this fiscal year (e.g. 9/1/2025).\n\n' +
      'This date must be the Monday that begins Period 1, Week 1.\n\n' +
      'Once set, new call log sheets will automatically be titled "P# W#".\n' +
      'Leave blank to use the "Week Ending MM/DD/YY" format instead.'
    );

  // --- Row 3: Master employee data spreadsheet ID ---
  configSheet.getRange('A3')
    .setValue('Employee Data Spreadsheet ID:')
    .setFontWeight('bold');

  configSheet.getRange(EMPLOYEE_DATA_SPREADSHEET_ID_CELL) // "B3"
    .setValue('')
    .setNote(
      'Paste the Google Spreadsheet ID of your master employee data source ' +
      '(the same spreadsheet used by the AutoScheduler for roster ingestion).\n\n' +
      'Find it in the spreadsheet URL:\n' +
      'docs.google.com/spreadsheets/d/ *** PASTE THIS PART *** /edit\n\n' +
      'Once set:\n' +
      '  \u2022 Typing an employee name in column B will autofill their ID and department.\n' +
      '  \u2022 The name becomes a clickable link that opens the master spreadsheet,\n' +
      '    so payroll clerks can navigate directly to the employee\'s record.\n\n' +
      'Leave blank to use the local "Employee Roster" sheet for lookups instead.'
    );

  // --- Row 5: Instructions block ---
  configSheet.getRange('A5')
    .setValue(
      'Row 2 — Enter the P1 W1 start date for the fiscal year. ' +
      'Row 3 — Paste the spreadsheet ID of the master employee data source. ' +
      'Both fields are optional but improve the experience significantly.'
    )
    .setFontStyle('italic')
    .setFontColor('#666666')
    .setWrap(true);

  configSheet.setColumnWidth(1, 230); // A — label column
  configSheet.setColumnWidth(2, 200); // B — input column
  configSheet.setColumnWidth(3, 400); // C — overflow for notes/instructions
}
