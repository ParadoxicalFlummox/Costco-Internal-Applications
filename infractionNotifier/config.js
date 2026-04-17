/**
 * config.js — Central configuration for the Costco Infraction Notifier.
 * VERSION: 1.0.0
 *
 * This file is the single source of truth for every configurable value in the
 * infraction detection and notification system. Nothing in any other file should
 * hard-code a value that lives here.
 *
 * HOW TO CUSTOMIZE:
 *   - To change infraction thresholds or window lengths, edit CODE_RULES or the
 *     global WINDOW_DAYS / THRESHOLD_COUNT fallbacks.
 *   - To add or remove codes that count as infractions, edit INFRACTION_CODES.
 *   - To add or remove codes that are always ignored, edit IGNORE_CODES.
 *   - To update payroll recipient addresses, edit PAYROLL_RECIPIENTS.
 *   - To change how long a CN stays Active before it expires, edit EXPIRY_DAYS.
 *   - To point the log at a different spreadsheet, update the ID in the
 *     "Infraction Config" sheet cell B2 (see LOG_SPREADSHEET_ID_CELL).
 */


// ---------------------------------------------------------------------------
// Behavior Flags
// ---------------------------------------------------------------------------

/**
 * When true, the scanner logs what it would do but sends no emails and writes
 * nothing to the CN_Log. Set to false in production.
 *
 * This can also be overridden per-call by passing { dryRun: true/false } to
 * scanAndIssueCNs() in infractionEngine.js.
 */
const DRY_RUN = true;

/**
 * Number of days to look back when scanning for infraction windows.
 *
 * 60 days covers two full 30-day windows and ensures that events near the
 * boundary of last month are not missed. For SE/SZ which use a 365-day window,
 * the daily trigger accumulates evidence across many runs.
 */
const DAYS_BACK = 60;


// ---------------------------------------------------------------------------
// Global Infraction Window (fallback when no per-code rule is defined)
// ---------------------------------------------------------------------------

/**
 * The default rolling window length in days. Used for any infraction code that
 * does not have an entry in CODE_RULES.
 */
const WINDOW_DAYS = 30;

/**
 * The default number of infractions within WINDOW_DAYS that triggers a CN.
 * Used for any code without an entry in CODE_RULES.
 */
const THRESHOLD_COUNT = 3;


// ---------------------------------------------------------------------------
// Per-Code Rules
// ---------------------------------------------------------------------------

/**
 * Infraction-specific threshold and window overrides.
 *
 * Each key is an attendance code (uppercase). The value sets the minimum
 * number of that specific code within the given window (in days) that
 * triggers a Counseling Notice.
 *
 * Rules are derived from the Costco Employee Agreement (March 2025):
 *
 *   TD — Tardy (§11.4.2):
 *     Three separate occasions of 4+ minutes in any 30-day period.
 *     threshold: 3, windowDays: 30
 *
 *   NS — No Show (§11.4.3e):
 *     Any failure to report without notification is a disciplinary cause.
 *     A single no-show triggers a CN.
 *     threshold: 1, windowDays: 30
 *
 *   SE — Swiping Error (§11.4.12a):
 *     Three separate failures to swipe consistently in a 30-day period.
 *     threshold: 3, windowDays: 30
 *
 *   MP — Meal Period Occurrence (§11.4.12b):
 *     Three failures to begin meal period on time in a 30-day period.
 *     threshold: 3, windowDays: 30
 *
 *   SZ — Suspension (§11.3.11a):
 *     Third unpaid suspension in a 12-month period triggers termination review.
 *     Track approaching the limit; CN issued at 3 within 365 days.
 *     threshold: 3, windowDays: 365
 *
 * All thresholds and windows are subject to GM discretion and may be adjusted
 * here without touching any other file.
 */
const CODE_RULES = {
  TD: { threshold: 3, windowDays: 30 },
  NS: { threshold: 1, windowDays: 30 },
  SE: { threshold: 3, windowDays: 30 },
  MP: { threshold: 3, windowDays: 30 },
  SZ: { threshold: 3, windowDays: 365 },
};


// ---------------------------------------------------------------------------
// Infraction and Ignore Code Lists
// ---------------------------------------------------------------------------

/**
 * Codes that count as infractions when they appear in an employee's calendar.
 *
 * A code must appear in this list OR have an entry in CODE_RULES to be counted.
 * Codes in CODE_RULES are automatically treated as infractions regardless of
 * whether they also appear here.
 *
 * Current infraction codes and their meanings:
 *   TD  — Tardy
 *   NS  — No Show
 *   SE  — Swiping Error
 *   MP  — Meal Period Occurrence
 *   SZ  — Suspension (unpaid)
 */
const INFRACTION_CODES = ['TD', 'NS', 'SE', 'MP', 'SZ'];

/**
 * Codes that are explicitly ignored and never counted as infractions,
 * even if they somehow match an entry in INFRACTION_CODES.
 *
 * Current ignore codes and their meanings:
 *   BL  — Bereavement Leave
 *   CN  — Counseling Notice Given  (recorded on the sheet; not a new infraction)
 *   FH  — Floating Holiday
 *   H   — Holiday
 *   JD  — Jury Duty
 *   SPF — Sick Paid Full
 *   SUF — Sick Unpaid Full
 *   SPH — Sick Paid Half
 *   SUH — Sick Unpaid Half
 *   LP  — Personal Non-Medical Leave
 *   NY  — New Year (holiday variant, used in some controllers)
 *   FJ  — (holiday/admin code, used in some controllers)
 */
const IGNORE_CODES = ['BL', 'CN', 'FH', 'H', 'JD', 'SPF', 'SUF', 'SPH', 'SUH', 'LP', 'NY', 'FJ'];


// ---------------------------------------------------------------------------
// CN Expiry Policy
// ---------------------------------------------------------------------------

/**
 * Number of days after a CN is issued before it is automatically marked
 * Expired in the CN_Log and a notification is sent to payroll.
 *
 * 180 days (~6 months) matches the Costco Employee Agreement retention policy
 * for most disciplinary codes. SZ entries (§11.3.11a) reference a 12-month
 * window, but CN expiry is tracked separately from the detection window.
 */
const EXPIRY_DAYS = 180;


// ---------------------------------------------------------------------------
// Attendance Controller Spreadsheet Layout
// ---------------------------------------------------------------------------

/**
 * The row bands that divide the attendance controller into three calendar
 * sections (one per grid of ~4 months each).
 *
 * Each band describes:
 *   monthRow     — The row containing month name headers (e.g. "JANUARY")
 *   dayOfWeekRow — The row containing day-of-week headers (MON, TUE, etc.)
 *   firstGridRow — The first row of actual calendar data (day numbers and codes)
 *
 * These values are standardized across Costco warehouses. If a specific
 * controller uses a different layout, update these values here.
 */
const DATA_BANDS = [
  { monthRow: 5, dayOfWeekRow: 6, firstGridRow: 7, lastGridRow: 30 },
  { monthRow: 32, dayOfWeekRow: 33, firstGridRow: 34, lastGridRow: 57 },
  { monthRow: 59, dayOfWeekRow: 60, firstGridRow: 61, lastGridRow: 83 },
];

/**
 * The A1-notation starting column for each of the four month blocks that
 * appear side-by-side across the attendance controller sheet.
 *
 * Each block covers 7 day columns. The 8th column in each block is a visual
 * separator and is not read.
 */
const START_COLUMNS = ['A', 'I', 'Q', 'Y'];

/**
 * The number of day-data columns to read per month block.
 * The 8th column (separator) is intentionally excluded.
 */
const DAY_COLS_PER_BLOCK = 7;

/**
 * A1-notation cell addresses for employee metadata fields on each individual
 * employee tab in the attendance controller.
 *
 * These are standardized across Costco warehouses:
 *   yearTitle    — D1  e.g. "2026 Attendance Controller"
 *   employeeName — X1  e.g. "Le, Tony"
 *   department   — R3  e.g. "Night Merch"
 *   employeeId   — X3  e.g. "1234578"
 *   hireDate     — AD3 e.g. date value
 */
const EMPLOYEE_FIELDS = {
  yearTitle: 'D1',
  employeeName: 'X1',
  department: 'R3',
  employeeId: 'X3',
  hireDate: 'AD3',
};

/**
 * The regex pattern used to identify which sheet tabs in the attendance
 * controller are individual employee sheets (vs. instruction or summary tabs).
 *
 * Employee tabs follow the format "Last, First - EmployeeNumber"
 * e.g. "Le, Tony - 1234578"
 *
 * Tabs whose names do NOT match this pattern (e.g. "(Employee Details)",
 * "(A - Instructions Link)") are skipped during the scan.
 */
const EMPLOYEE_TAB_PATTERN = /^.+,\s*.+\s*-\s*\d+$/;


// ---------------------------------------------------------------------------
// CN Log Configuration
// ---------------------------------------------------------------------------

/**
 * The name of the sheet tab within the CN log spreadsheet where issued CNs
 * are recorded.
 */
const CN_LOG_SHEET_NAME = 'CN_Log';

/**
 * The ordered list of column headers for the CN_Log sheet.
 *
 * Column meanings:
 *   CN_Key       — Unique deduplication key: "EmployeeID|RULE:code|windowStart|windowEnd"
 *   EmployeeID   — Employee number from the attendance controller
 *   EmployeeName — Employee name as read from the controller
 *   Department   — Department from the controller
 *   WindowStart  — Start date of the infraction window (YYYY-MM-DD)
 *   WindowEnd    — End date of the infraction window (YYYY-MM-DD)
 *   Count        — Number of infraction events in the window
 *   EventsHash   — SHA-1 of the events list; used to detect if events changed
 *   IssuedAt     — Timestamp when the CN was first issued
 *   IssuedBy     — Email of the script runner (Session.getActiveUser())
 *   DryRun       — TRUE if the CN was generated in dry-run mode
 *   SheetName    — Name of the employee tab the events were parsed from
 *   Status       — "Active" or "Expired"
 *   ExpiredAt    — Timestamp when the CN was marked Expired (blank if still Active)
 *   Rule         — The code rule that triggered this CN (e.g. "TD", "NS", "GLOBAL")
 */
const CN_LOG_HEADERS = [
  'CN_Key', 'EmployeeID', 'EmployeeName', 'Department',
  'WindowStart', 'WindowEnd', 'Count', 'EventsHash',
  'IssuedAt', 'IssuedBy', 'DryRun', 'SheetName',
  'Status', 'ExpiredAt', 'Rule',
  'SourceSpreadsheetId', 'SourceSheetGid',
];

/**
 * The name of the manager-facing sheet that lists all currently Active CNs.
 * Each row has a clickable hyperlink in the Employee Name column that opens
 * the employee's tab in the attendance controller directly.
 *
 * This sheet lives in the same workbook as the CN_Log (either the external
 * log workbook or the active workbook as a fallback).
 */
const ACTIVE_CNS_SHEET_NAME = 'Active CNs';

/**
 * The name of the archive sheet that holds CNs that have passed their
 * expiry date. The parentheses prefix follows the attendance controller
 * convention for hidden/reference sheets that are not part of the main
 * workflow.
 *
 * When a CN expires, its row is moved from Active CNs to this sheet and
 * the sheet is hidden so it does not clutter the tab bar. It remains
 * accessible via right-click → Show Sheet if needed for audit purposes.
 */
const EXPIRED_CNS_SHEET_NAME = '(Expired CNs)';

/**
 * Column headers for the Active CNs sheet.
 *
 *   CN_Key        — Internal dedup key; used to locate the row on expiry
 *   Employee Name — HYPERLINK formula linking to the employee's tab
 *   Employee ID   — Employee number
 *   Department    — Home department
 *   Rule          — The code that triggered this CN (e.g. "TD", "NS")
 *   Count         — Number of occurrences in the window
 *   Window Start  — Start date of the infraction window
 *   Window End    — End date of the infraction window
 *   Issued At     — Timestamp when the CN was issued
 *   Sheet         — Name of the source employee tab
 */
const ACTIVE_CNS_HEADERS = [
  'CN_Key', 'Employee Name', 'Employee ID', 'Department',
  'Rule', 'Count', 'Window Start', 'Window End', 'Issued At', 'Sheet',
];

/**
 * Column headers for the (Expired CNs) sheet.
 * Same as ACTIVE_CNS_HEADERS with an Expired At column appended.
 * Rows are moved here intact from Active CNs when they age out.
 */
const EXPIRED_CNS_HEADERS = [
  'CN_Key', 'Employee Name', 'Employee ID', 'Department',
  'Rule', 'Count', 'Window Start', 'Window End', 'Issued At', 'Sheet',
  'Expired At',
];


// ---------------------------------------------------------------------------
// Config Sheet (CN Log Spreadsheet ID input)
// ---------------------------------------------------------------------------

/**
 * The name of the configuration sheet tab in the attendance controller workbook.
 * This sheet holds the CN Log spreadsheet ID so the script knows where to write
 * CN records.
 *
 * If this sheet does not exist, the script falls back to writing the CN_Log
 * into the active (attendance controller) workbook.
 */
const INFRACTION_CONFIG_SHEET_NAME = 'Infraction Config';

/**
 * The A1-notation cell on the INFRACTION_CONFIG_SHEET_NAME sheet where the
 * CN Log spreadsheet ID is entered.
 *
 * The ID is the alphanumeric string in the log spreadsheet's URL:
 *   docs.google.com/spreadsheets/d/ *** THIS PART *** /edit
 *
 * Leaving this cell blank causes the script to write the CN_Log into the
 * active workbook instead.
 */
const LOG_SPREADSHEET_ID_CELL = 'B2';


// ---------------------------------------------------------------------------
// Notification Recipients
// ---------------------------------------------------------------------------

/**
 * The email addresses that receive all CN notifications (both new CNs and
 * expiry notifications).
 *
 * All departments currently route to payroll. Per-department routing may be
 * added in a future version by introducing a MAILING_LIST map keyed by
 * department name, following the same pattern as the Absence Notifier.
 */
const PAYROLL_RECIPIENTS = [
  'payroll.clerk@example.com',
];
