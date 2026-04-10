/**
 * config.js — Central configuration for the Costco Absence Notifier.
 * VERSION: 0.2.1
 *
 * This file is the single source of truth for every configurable value in the
 * notification system. Nothing in any other file should hard-code a value that
 * lives here. If a business rule changes — such as the notification window length,
 * a manager's email address, or the column layout of the call log sheet — this is
 * the only file that needs to be updated.
 *
 * HOW TO CUSTOMIZE:
 *   - To change how often the digest runs, edit WINDOW_MINUTES and update the
 *     Apps Script time-driven trigger cadence to match.
 *   - To update which spreadsheet columns hold which data, edit CALL_LOG_COLUMNS.
 *   - To add a department or update manager email addresses, edit MAILING_LIST.
 *   - To change who receives notifications for unrecognized departments, edit
 *     FALLBACK_RECIPIENTS.
 *   - To change which sheet holds the FY start date or the employee roster,
 *     edit CONFIG_SHEET_NAME, FISCAL_YEAR_START_CELL, or ROSTER_SHEET_NAME.
 */


// ---------------------------------------------------------------------------
// Supporting Sheet Names
// ---------------------------------------------------------------------------

/**
 * The name of the sheet tab that holds system-wide configuration for the
 * absence notifier (currently just the fiscal year start date).
 *
 * This sheet is created automatically by setupAbsenceConfigSheet() in
 * sheetGenerator.js the first time setup is run. Managers use it to enter
 * the FY start date that drives the P# W# sheet name calculation.
 */
const CONFIG_SHEET_NAME = "Absence Config";

/**
 * The A1-notation cell address on the CONFIG_SHEET_NAME sheet where the
 * fiscal year start date is entered by the manager.
 *
 * sheetUtils.js reads this cell to calculate the current period and week.
 * If the cell is empty or contains an invalid date, the sheet name falls
 * back to the "Week Ending MM/DD/YY" format automatically.
 */
const FISCAL_YEAR_START_CELL = "B2";

/**
 * The name of the sheet tab in the call log workbook that contains the
 * employee roster used for autofill.
 *
 * This sheet is maintained manually by managers (or in a future version,
 * synced from the AutoScheduler's Roster sheet). It must have:
 *   Column A — Employee Name
 *   Column B — Employee ID
 *   Column C — Home Department
 *
 * autofill.js reads this sheet when a manager types a name in column B
 * of a call log entry row.
 */
const ROSTER_SHEET_NAME = "Employee Roster";


// ---------------------------------------------------------------------------
// Call Log Column Positions (0-indexed for use with getValues() array reads)
// ---------------------------------------------------------------------------

/**
 * Maps each logical field name to its 0-based column index in the generated
 * call log sheet.
 *
 * The generated call log sheet layout is:
 *   A (0)  — Date of absence
 *   B (1)  — Employee Name      (autofill trigger — populates C and D on edit)
 *   C (2)  — Employee ID        (autofilled from Employee Roster)
 *   D (3)  — Home Department    (autofilled from Employee Roster)
 *   E (4)  — Call-Out checkbox
 *   F (5)  — FMLA checkbox
 *   G (6)  — No-Show checkbox
 *   H (7)  — Time Called
 *   I (8)  — Scheduled Shift
 *   J (9)  — Arrival Time (for tardy entries)
 *   K (10) — Manager Approval
 *   L (11) — Comment
 *
 * Using named constants instead of raw numbers means a column rearrangement
 * requires updating only this object rather than hunting through every file
 * that reads the sheet.
 */
const CALL_LOG_COLUMNS = {
  DATE:             0,  // A — Date of the absence
  NAME:             1,  // B — Employee full name (autofill trigger)
  EMPLOYEE_ID:      2,  // C — Employee ID (autofilled)
  DEPT:             3,  // D — Home department (autofilled)
  IS_CALLOUT:       4,  // E — Call-Out checkbox (TRUE/FALSE)
  IS_FMLA:          5,  // F — FMLA checkbox (TRUE/FALSE)
  IS_NOSHOW:        6,  // G — No-Show checkbox (TRUE/FALSE)
  TIME_CALLED:      7,  // H — Time the employee called in
  SCHEDULED_SHIFT:  8,  // I — The employee's scheduled shift for that day
  ARRIVAL_TIME:     9,  // J — Actual arrival time if the entry is a tardy
  MANAGER_APPROVAL: 10, // K — Manager initials or name confirming the entry
  COMMENT:          11, // L — Employee's voluntary comment

  // The total number of columns to read in a single getRange() call.
  // Must cover all columns above: index 11 (L) + 1 = 12 columns (A–L).
  TOTAL_COLUMNS_TO_READ: 12,
};

/**
 * The spreadsheet row where call log entry data begins.
 * Row 1 is the fiscal period header, row 2 is the column headers row.
 */
const CALL_LOG_DATA_START_ROW = 3;

/**
 * The 1-based column number for the Employee Name column.
 * Used by autofill.js to identify which column triggers a roster lookup.
 * Must stay in sync with CALL_LOG_COLUMNS.NAME (0-indexed) + 1.
 */
const CALL_LOG_NAME_COLUMN_NUMBER = 2; // Column B


// ---------------------------------------------------------------------------
// Notification Window Configuration
// ---------------------------------------------------------------------------

/**
 * The length of each digest window in minutes.
 *
 * The notifier fires on a time-driven trigger. Each run looks back at the most
 * recently completed window of this length to find rows that were logged during
 * that period. For example, if the trigger fires at 09:15 and WINDOW_MINUTES is
 * 15, the script scans for rows with a time in (09:00, 09:15].
 *
 * IMPORTANT: The Apps Script time-driven trigger cadence must match this value.
 * If you change WINDOW_MINUTES to 30, update the trigger to run every 30 minutes.
 */
const WINDOW_MINUTES = 15;


// ---------------------------------------------------------------------------
// Mailing List and Fallback Recipients
// ---------------------------------------------------------------------------

/**
 * Maps department names (as they appear in column D of the call log) to one
 * or more recipient email addresses for that department's absence digest.
 *
 * Keys are matched case-insensitively against the value in column D.
 * Each value may be a single email string, an array of email strings, or an
 * array containing a semicolon/comma-separated string — all formats are
 * normalized by resolveRecipientsForDepartment_() in notifier.js.
 *
 * To add a department: add a new key-value entry below.
 * To update a manager's address: change the string in the array.
 */
const MAILING_LIST = {
  "Administration":  ["w01119mgr@costco.com;w01119mgr3@costco.com;w01119adm@costco.com"],
  "Bakery":          ["bakery.manager@example.com"],
  "Center":          ["center.manager@example.com"],
  "Food Court":      ["foodcourt.manager@example.com"],
  "Foods":           ["foods.manager@example.com"],
  "Front End":       ["frontend.manager@example.com"],
  "Gasoline":        ["gasoline.manager@example.com"],
  "Hardlines":       ["hardlines.manager@example.com"],
  "Hearing Aids":    ["hearingaids.manager@example.com"],
  "Maintenance":     ["maintenance.manager@example.com"],
  "Meat":            ["meat.manager@example.com"],
  "Membership":      ["membership.manager@example.com"],
  "Merchandising":   ["merchandising.manager@example.com"],
  "Night Merch":     ["w01119nm01@costco.com"],
  "Optical":         ["optical.manager@example.com"],
  "Produce":         ["produce.manager@example.com"],
  "Receiving":       ["receiving.manager@example.com"],
  "Sales":           ["sales.manager@example.com"],
  "Security":        ["security.manager@example.com"],
  "Service Deli":    ["servicedeli.manager@example.com"],
  "Tire Shop":       ["tireshop.manager@example.com"],
  "Whse Loss Prev":  ["whselossprev.manager@example.com"],
};

/**
 * The email addresses that receive a digest when a row's department either does
 * not match any key in MAILING_LIST or is blank.
 *
 * This acts as a safety net so that no absence notification is silently dropped
 * due to a typo in column D or an unlisted department.
 */
const FALLBACK_RECIPIENTS = [
  "admin.management@example.com",
];
