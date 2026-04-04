/**
 * config.js — Central configuration constants for the Auto Schedule Generator.
 * VERSION: 0.3.3
 *
 * This file is the single source of truth for every magic number, color, column
 * position, and rule in the system. Nothing in any other file should hard-code a
 * value that lives here. If a business rule changes (e.g., the minimum hours for
 * a part-time employee), this is the only file that needs to be updated.
 *
 * HOW TO CUSTOMIZE:
 *   - To change hour rules, edit HOUR_RULES.
 *   - To change colors on the generated schedule, edit COLORS.
 *   - To add a new sheet or rename an existing one, edit SHEET_NAMES.
 *   - To change which column holds which data on the Roster sheet, edit ROSTER_COLUMN.
 *   - Do NOT reorder Roster sheet columns without updating ROSTER_COLUMN to match.
 */


// ---------------------------------------------------------------------------
// Sheet Names
// ---------------------------------------------------------------------------

/**
 * The canonical names for every sheet tab in the workbook.
 *
 * Using a central map instead of string literals scattered across files means
 * a sheet rename requires changing exactly one value here rather than hunting
 * through every file for the old string.
 */
const SHEET_NAMES = {
  INGESTION: "Ingestion",
  ROSTER: "Roster",
  SETTINGS: "Settings",
};


// ---------------------------------------------------------------------------
// Roster Sheet Column Positions (1-indexed for use with GAS range operations)
// ---------------------------------------------------------------------------

/**
 * Maps each logical field to its column number (1 = column A) in the Roster sheet.
 *
 * All reads and writes to the Roster sheet use these constants so that moving a
 * column only requires updating the number here, not touching every function that
 * accesses that column.
 */
const ROSTER_COLUMN = {
  NAME: 1,  // A — Employee full name, synced from the source spreadsheet
  EMPLOYEE_ID: 2,  // B — Unique employee identifier, used as the deduplication key during sync
  HIRE_DATE: 3,  // C — Date the employee was hired; drives the seniority rank calculation
  STATUS: 4,  // D — Employment status: "FT" (full-time) or "PT" (part-time)
  DAY_OFF_PREF_ONE: 5,  // E — First preferred day off (e.g., "Monday")
  DAY_OFF_PREF_TWO: 6,  // F — Second preferred day off (e.g., "Tuesday")
  PREFERRED_SHIFT: 7,  // G — The shift name the employee prefers (must match a Settings shift name)
  QUALIFIED_SHIFTS: 8,  // H — Comma-separated list of shift names this employee is trained to work
  VACATION_DATES: 9,  // I — Comma-separated vacation dates (YYYY-MM-DD or MM/DD format)
  SENIORITY_RANK: 10, // J — Calculated by the script; do not edit manually
};

/**
 * The row number where Roster data begins (row 1 is the header).
 */
const ROSTER_DATA_START_ROW = 2;


// ---------------------------------------------------------------------------
// Settings Sheet Range Addresses
// ---------------------------------------------------------------------------

/**
 * The A1-notation ranges for each table on the Settings sheet.
 *
 * The Settings sheet contains two separate tables:
 *   1. The staffing requirements table (how many employees are needed each day).
 *   2. The shift definitions table (shift names, times, and paid hours).
 *
 * If you need to add more shift rows, increase the end row of SHIFT_DEFINITIONS_TABLE.
 */
const SETTINGS_RANGE = {
  STAFFING_REQUIREMENTS_TABLE: "A2:B8",   // 7 rows, one per day of the week
  SHIFT_DEFINITIONS_TABLE: "D2:I50",  // Up to 49 shift rows; expand if needed
};

/**
 * Column offsets (0-indexed) within the shift definitions table read from SETTINGS_RANGE.SHIFT_DEFINITIONS_TABLE.
 *
 * The raw data array returned by getValues() uses 0-based column offsets,
 * so these constants convert from logical name to array position.
 */
const SHIFT_TABLE_COLUMN = {
  NAME: 0, // D — The display name of the shift (e.g., "Morning", "Closing")
  STATUS: 1, // E — "FT" or "PT" — this row applies only to employees of this status
  START_TIME: 2, // F — The shift start time as a GAS time value (decimal fraction of a day)
  END_TIME: 3, // G — The shift end time as a GAS time value (includes unpaid lunch block for FT)
  PAID_HOURS: 4, // H — Hours counted toward the employee's weekly minimum/maximum
  HAS_LUNCH: 5, // I — TRUE if this shift includes an unpaid 30-minute lunch break
};


// ---------------------------------------------------------------------------
// Generated Week Sheet Layout
// ---------------------------------------------------------------------------

/**
 * Row and column positions for the generated Week_MM_DD_YY schedule sheets.
 *
 * Each employee occupies a block of three consecutive rows:
 *   Row 1 of block: VAC — vacation checkboxes (checked = employee is on vacation that day)
 *   Row 2 of block: RDO — requested day off checkboxes (checked = employee requested the day off)
 *   Row 3 of block: SHIFT — the assigned shift text (e.g., "08:00 - 16:30") or "OFF"
 *
 * The block structure makes it visually clear which three rows belong to one employee,
 * and it allows the onEdit handler to determine which type of checkbox was changed
 * by computing (rowNumber - DATA_START_ROW) % 3.
 */
const WEEK_SHEET = {
  HEADER_ROW: 1,  // Row containing the merged week label (e.g., "Week of April 7 – 13, 2026")
  TIMESTAMP_ROW: 2,  // Row showing when this draft was last generated
  DEPARTMENT_ROW: 3,  // Row showing the department name
  COLUMN_HEADER_ROW: 5, // Row containing "Label | Employee | Mon | Tue | ... | Sun | Total Hrs"
  DATA_START_ROW: 6,  // First row of employee data blocks

  // Column positions (1-indexed)
  COL_ROW_LABEL: 1,  // A — "VAC", "RDO", or "SHIFT" label for each row
  COL_EMPLOYEE_NAME: 2, // B — Employee name (merged across all 3 rows of the block)
  COL_MONDAY: 3,  // C
  COL_TUESDAY: 4,  // D
  COL_WEDNESDAY: 5,  // E
  COL_THURSDAY: 6,  // F
  COL_FRIDAY: 7,  // G
  COL_SATURDAY: 8,  // H
  COL_SUNDAY: 9,  // I
  COL_TOTAL_HOURS: 10, // J — Weekly paid hours total, written by the script (not a formula)

  ROWS_PER_EMPLOYEE: 3,  // VAC + RDO + SHIFT
  ROW_OFFSET_VAC: 0,  // Offset from the employee block start row for the VAC row
  ROW_OFFSET_RDO: 1,  // Offset from the employee block start row for the RDO row
  ROW_OFFSET_SHIFT: 2,  // Offset from the employee block start row for the SHIFT row

  DAYS_IN_WEEK: 7,   // Monday through Sunday
};

/**
 * The day names in the order they appear as column headers (Monday first).
 * Used when matching staffing requirements by day name and when computing
 * the date for each column during schedule generation.
 */
const DAY_NAMES_IN_ORDER = [
  "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"
];


// ---------------------------------------------------------------------------
// Hour Rules
// ---------------------------------------------------------------------------

/**
 * Weekly paid hour minimums and maximums for full-time and part-time employees.
 *
 * These rules are enforced by Phase 2 of the generation algorithm (minimum
 * enforcement) and by Phase 1 and Phase 3 as a cap guard (maximum enforcement).
 *
 * FT employees must work exactly 40 hours — the minimum equals the maximum,
 * which means a full-time employee will always be scheduled to exactly 40 hours
 * unless they have too many vacation days in the week.
 *
 * PT employees must work between 24 and 35 hours. The engine first tries to
 * honor their preferred days off (which may leave them below 35), and then
 * Phase 2 adds shifts until they reach 24 if they fell short.
 *
 * The keys intentionally use the short forms MIN and MAX because these are
 * well-understood universal abbreviations, unlike domain-specific ones.
 */
const HOUR_RULES = {
  FT_MIN: 40,
  FT_MAX: 40,
  PT_MIN: 24,
  PT_MAX: 40,
  // PT+ (lunch-qualified) shifts may have between 5 and 8 paid hours per shift.
  // The engine and Settings validation use these bounds when evaluating + shifts.
  PT_PLUS_MIN_HOURS: 5,
  PT_PLUS_MAX_HOURS: 8,
};


// ---------------------------------------------------------------------------
// Schedule Rules
// ---------------------------------------------------------------------------

/**
 * High-level scheduling constraints that apply across all phases of the engine.
 *
 * These are distinct from HOUR_RULES (which govern weekly paid hour targets) because
 * they control structural properties of the schedule — how many days off each employee
 * must receive regardless of their hour budget.
 */
const SCHEDULE_RULES = {
  // Every employee must have at least this many non-working days each week.
  // "Non-working" includes OFF, RDO, and VAC — any cell that is not a SHIFT.
  // Phase 1 enforces this after assigning preferred shifts (choosing the best-covered
  // days to force off), Phase 2 respects it as a cap on additional shift assignments,
  // and Phase 3 Cascade B skips employees who are already at this floor.
  MIN_DAYS_OFF: 2,
};


// ---------------------------------------------------------------------------
// Seniority Rank Formula Constants
// ---------------------------------------------------------------------------

/**
 * Constants used by the seniority rank calculation.
 *
 * The rank formula produces a single integer that encodes both employment status
 * and length of service. Full-time employees receive a base of 200,000,000 and
 * part-time employees receive a base of 100,000,000. The days-from-hire value
 * (computed relative to a future reference date so that older hire dates produce
 * larger numbers) is added to the base.
 *
 * This means:
 *   - Any FT employee will always outrank any PT employee with the same hire date.
 *   - Within the same status, employees hired earlier receive a higher rank.
 *   - The 100,000,000 gap between FT_BASE and PT_BASE is large enough that no
 *     realistic days-from-hire value can close it (100M days ≈ 273,000 years).
 */
const SENIORITY = {
  FT_BASE: 200000000,
  PT_BASE: 100000000,
  // An arbitrary future date used as the anchor for the days-from-hire calculation.
  // Subtracting the hire date from this future date produces a larger number for
  // employees hired earlier, giving them a higher seniority rank without any
  // conditional logic. The specific date of 2050-01-01 was chosen because it is
  // far enough in the future that all current employees will have been hired
  // before it, and close enough that the resulting integers remain manageable.
  REFERENCE_DATE_STRING: "2050-01-01",
};


// ---------------------------------------------------------------------------
// Coverage Slot Map Constants
// ---------------------------------------------------------------------------

/**
 * Defines the 30-minute coverage slot array used to track shift coverage.
 *
 * The coverage map represents the operating day as an array of 30-minute windows.
 * Each element counts how many employees are scheduled to be physically present
 * during that window. The engine uses this map to detect gaps (zero-count slots)
 * and to score shifts by how well they fill uncovered periods.
 *
 * The day is modeled from 04:00 to 23:30 (slot indices 0 through 38), covering
 * the earliest possible opening shift through the latest possible closing shift.
 * If your operation has shifts that start before 04:00 or end after 23:30, you
 * must adjust COVERAGE_START_MINUTE and SLOT_COUNT accordingly.
 *
 * Slot index calculation: Math.floor((minutesSinceMidnight - COVERAGE_START_MINUTE) / SLOT_DURATION_MINUTES)
 */
const COVERAGE = {
  SLOT_COUNT: 39,  // Number of 30-minute windows in the coverage day (04:00–23:30)
  COVERAGE_START_MINUTE: 240, // 04:00 expressed as minutes since midnight (4 * 60 = 240)
  SLOT_DURATION_MINUTES: 30,  // Each slot represents 30 minutes of clock time
};

// ---------------------------------------------------------------------------
// Coverage Windows (required staffing hours per day)
// ---------------------------------------------------------------------------

/**
 * Defines the start and end of the required coverage window for each day of the week
 * 
 * The coverage gap detection in Phase 3 only checks for uncovered slots within these
 * windows. Slots outside of the window are not treated as gaps, the engine will not
 * pull in employees or reassign shifts to cover time outside these bounds.
 * 
 * All times are expressed as minutes since midnight:
 *    240  = 4:00 AM
 *    1320 = 10:00 PM
 *    1260 = 9:00 PM
 *    1410 = 11:30 PM
 * 
 * To change a day's requirements, edit the endMinute value for that day.
 * The startMinute should generally stay at 240 to match COVERAGE_START_MINUTE
 */
const COVERAGE_WINDOW = {
  Monday: { startMinute: 240, endMinute: 1410 },
  Tuesday: { startMinute: 240, endMinute: 1410 },
  Wednesday: { startMinute: 240, endMinute: 1410 },
  Thursday: { startMinute: 240, endMinute: 1410 },
  Friday: { startMinute: 240, endMinute: 1410 },
  Saturday: { startMinute: 240, endMinute: 1320 },
  Sunday: { startMinute: 240, endMinute: 1260 }
};


// ---------------------------------------------------------------------------
// Colors
// ---------------------------------------------------------------------------

/**
 * Hex color codes applied to cells in generated schedule sheets.
 *
 * The color scheme communicates shift type at a glance:
 *   - Blue cells are full-time shifts (stable, full-day coverage).
 *   - Green cells are part-time shifts (shorter coverage window).
 *   - Yellow cells are vacation days (no coverage; employee is absent).
 *   - Gray cells are days off (no coverage; employee is not scheduled).
 *   - Red on the employee name signals the employee is below their weekly minimum.
 *
 * To change the color scheme, update the hex values here. All formatting
 * code in formatter.js reads from this object, so no other file needs to change.
 */
const COLORS = {
  FT_SHIFT: "#4A90D9", // Blue — full-time shift cell background
  PT_SHIFT: "#57BB8A", // Green — part-time shift cell background
  VACATION: "#FFD966", // Yellow — vacation day cell background
  DAY_OFF: "#B7B7B7", // Gray — day off cell background
  UNDER_HOURS: "#E06666", // Red — employee name cell background when below weekly minimum
  OVER_HOURS_FT: "#FF9900", // Orange — FT employee name cell background when above 40 weekly hours
  HEADER_BG: "#263238", // Dark slate — column header row background
  HEADER_TEXT: "#FFFFFF", // White — column header row text
  SUMMARY_OK: "#B7E1CD", // Light green — STATUS row cell when coverage is met
  SUMMARY_UNDER: "#F4C7C3", // Light red — STATUS row cell when coverage is short
  ROW_LABEL_BG: "#F5F5F5", // Light gray — VAC/RDO/SHIFT label column background
};


// ---------------------------------------------------------------------------
// Ingestion Sheet Cell Addresses
// ---------------------------------------------------------------------------

/**
 * Cell addresses on the Ingestion sheet where the script reads inputs and writes results.
 *
 * Using named constants here means that if the Ingestion sheet layout changes,
 * only this object needs to be updated rather than searching through rosterIngestion.js
 * for hard-coded cell addresses.
 */
const INGESTION_CELL = {
  SOURCE_SPREADSHEET_ID: "B3", // Where the manager types the source spreadsheet ID
  DEPARTMENT: "B4", // Dropdown for the selected department
  SYNC_STATUS: "B8", // Script writes the sync result summary here
  EMPLOYEES_ADDED: "B9", // Script writes the count of newly added employees
  EMPLOYEES_SKIPPED: "B10", // Script writes the count of skipped duplicates
};
