/**
 * config.js — Unified configuration constants for COMET.
 * VERSION: 0.5.7
 *
 * This file is the single source of truth for every magic number, color, column
 * position, and rule across all COMET modules. Nothing in any other file should
 * hard-code a value that lives here.
 *
 * ORGANIZATION:
 *   1. App Identity
 *   2. Employees Sheet
 *   3. Schedule Engine (shift settings, week sheet layout, hour rules, coverage)
 *   4. Schedule Rules & Supervisor Scheduling
 *   5. Split-Shift (Multi-Department) Scheduling
 *   6. Seniority Rank Formula Constants
 *   7. Coverage Slot Map Constants
 *   8. Colors and Role Colors
 *   9. Infraction / CN System
 *   10. Absence Log (Call Log)
 *   11. Notification Recipients
 *   12. Performance Tuning
 */


// ---------------------------------------------------------------------------
// App Identity
// ---------------------------------------------------------------------------

/** Title shown in the browser tab and the top-level HTML template. */
const COMET_APP_TITLE = 'COMET — Warehouse Management';

/** Name of the COMET runtime-settings sheet (FY start date, window minutes, etc.). */
const COMET_CONFIG_SHEET_NAME = 'COMET Config';


// ---------------------------------------------------------------------------
// Employees Sheet
// ---------------------------------------------------------------------------

/** Tab name for the master employee roster. */
const EMPLOYEES_SHEET_NAME = 'Employees';

/** Row number where employee data begins (row 1 is the header). */
const EMPLOYEES_DATA_START_ROW = 2;

/** Number of days before the Employees sheet is considered stale. */
const EMPLOYEE_SHEET_STALENESS_THRESHOLD_DAYS = 7;

/**
 * Column positions (1-indexed) for every field in the Employees sheet.
 *
 * Columns A–E are populated by UKG import.
 * Columns F–N are schedule-specific fields, editable via the Admin employee modal.
 *
 *   A — Name (Last, First)
 *   B — Employee ID
 *   C — Hire Date
 *   D — Department
 *   E — Status ("Active", "LOA", or "Archived")
 *         Active  = employed and available for scheduling
 *         LOA     = on Leave of Absence; excluded from schedule generation
 *         Archived = no longer employed (terminated / resigned)
 *   F — FT/PT             ("FT", "PT", or "LPT")
 *         FT  = full-time, 40 h/week
 *         PT  = part-time, 24–40 h/week
 *         LPT = limited part-time (students / weekenders), 0–12 h/week
 *   G — Day Off Pref 1    (day name, e.g. "Monday")
 *   H — Day Off Pref 2    (day name)
 *   I — Preferred Shift   (shift name from dept settings)
 *   J — Qualified Shifts  (comma-separated shift names)
 *   K — Vacation Dates    (comma-separated dates, YYYY-MM-DD or MM/DD)
 *   L — Role              (e.g. "Cashier", "Maintenance Associate")
 *   M — Seniority Rank    (calculated; do not edit manually)
 *   N — Secondary Depts   (comma-separated dept names; cross-dept scheduling)
 */
const EMPLOYEE_COLUMN = {
  NAME:                  1,  // A
  ID:                    2,  // B
  HIRE_DATE:             3,  // C
  DEPARTMENT:            4,  // D
  STATUS:                5,  // E
  FTPT:                  6,  // F
  DAY_OFF_PREF_ONE:      7,  // G
  DAY_OFF_PREF_TWO:      8,  // H
  PREFERRED_SHIFT:       9,  // I
  QUALIFIED_SHIFTS:      10, // J
  VACATION_DATES:        11, // K
  ROLE:                  12, // L
  SENIORITY_RANK:        13, // M
  SECONDARY_DEPARTMENTS: 14, // N
};


// ---------------------------------------------------------------------------
// Department Settings Sheets
// ---------------------------------------------------------------------------

/**
 * Prefix for per-department settings sheet names.
 * e.g. "Settings_Maintenance", "Settings_Night Merch"
 */
const DEPT_SETTINGS_PREFIX = 'Settings_';

/** Default staffing head-count written to new Settings sheets. */
const DEFAULT_STAFFING_COUNT = 6;

/** Staffing mode values used in the Settings sheet Mode column. */
const STAFFING_MODE = {
  COUNT: 'Count',  // Minimum employee head count per day
  HOURS: 'Hours',  // Minimum total paid hours per day
};

/**
 * A1-notation range addresses for each table on a Settings_[Dept] sheet.
 *
 *   STAFFING_REQUIREMENTS_TABLE — A2:C8  (Day | Count | Mode, 7 rows for Mon–Sun)
 *   SHIFT_DEFINITIONS_TABLE     — E2:N50 (Name | FT/PT | WkdyStart | SatStart | SunStart |
 *                                          PaidHours | HasLunch | FlexEnabled | FlexEarliest | FlexLatest)
 *                                 Column D is a visual spacer; not read.
 *
 * New sections added below the staffing requirements block (rows 9–onward, cols A–D):
 *   ENGINE_OPTIONS_HEADER_ROW   — Row 10: section title
 *   ENGINE_OPTIONS_START_ROW    — Row 11: first option (Enforce Role Minimums)
 *   ENGINE_OPTIONS_COUNT        — 2 options (enforceRoleMinimums, gapFillEnabled)
 *   ROLE_MINIMUMS_HEADER_ROW    — Row 14: section title
 *   ROLE_MINIMUMS_LABEL_ROW     — Row 15: column headers (Role | Low | Moderate | High)
 *   ROLE_MINIMUMS_START_ROW     — Row 16: first data row
 */
const SETTINGS_RANGE = {
  STAFFING_REQUIREMENTS_TABLE: 'A2:C8',
  SHIFT_DEFINITIONS_TABLE:     'E2:N50',
};

const SETTINGS_ROWS = {
  ENGINE_OPTIONS_HEADER: 10,  // merged section label
  ENGINE_OPTIONS_START:  11,  // row 11 = Enforce Role Minimums, row 12 = Enable Gap Fill
  ENGINE_OPTIONS_COUNT:  2,
  ROLE_MINIMUMS_HEADER:  14,  // merged section label
  ROLE_MINIMUMS_LABELS:  15,  // Role | Low | Moderate | High column headers
  ROLE_MINIMUMS_START:   16,  // first role data row
};

/**
 * Column offsets (0-indexed) within the array returned by reading SHIFT_DEFINITIONS_TABLE (E2:N50).
 *
 * Shift start times are defined as per-day anchors instead of a single start/end pair:
 *   WEEKDAY_START — Mon–Fri anchor (required)
 *   SAT_START     — Saturday override (optional; falls back to WEEKDAY_START if blank)
 *   SUN_START     — Sunday override   (optional; falls back to WEEKDAY_START if blank)
 *
 * End time is computed: anchorStart + paidHours × 60 + (hasLunch ? 30 : 0)
 */
const SHIFT_TABLE_COLUMN = {
  NAME:          0,  // E — Shift display name (e.g. "Morning", "Night Merch")
  STATUS:        1,  // F — "FT" or "PT" — this shift applies only to employees of this status
  WEEKDAY_START: 2,  // G — Mon–Fri anchor start time (stored as "HH:MM" string)
  SAT_START:     3,  // H — Saturday anchor start time (blank = same as WEEKDAY_START)
  SUN_START:     4,  // I — Sunday anchor start time   (blank = same as WEEKDAY_START)
  PAID_HOURS:    5,  // J — Hours counted toward the employee's weekly minimum/maximum
  HAS_LUNCH:     6,  // K — TRUE if this shift includes an unpaid 30-minute lunch break
};


// ---------------------------------------------------------------------------
// Generated Week Sheet Layout
// ---------------------------------------------------------------------------

/**
 * Row and column positions for generated Week_MM_DD_YY_[Dept] schedule sheets.
 *
 * Sheets are pure data storage — one row per employee, JSON payload in col C.
 * All display and interaction happens in the web UI; managers never edit these sheets directly.
 *
 *   Row 1: Week label ("Week of April 28 – May 4, 2026")
 *   Row 2: "Generated: [timestamp]"
 *   Row 3: "Department: [deptName]"
 *   Row 4: (spacer)
 *   Row 5: Column headers — Name | Employee ID | Schedule | Total Hours | Stored At
 *   Row 6+: One row per employee
 *     Col A (1): Employee name
 *     Col B (2): Employee ID
 *     Col C (3): JSON string — { Monday: {type,shiftName,displayText,paidHours,role,locked}, ... }
 *     Col D (4): Total paid hours (number)
 *     Col E (5): ISO timestamp of last write
 */
const WEEK_SHEET = {
  HEADER_ROW:        1,
  TIMESTAMP_ROW:     2,
  DEPARTMENT_ROW:    3,
  COLUMN_HEADER_ROW: 5,
  DATA_START_ROW:    6,

  COL_NAME:          1,  // A — employee name
  COL_EMPLOYEE_ID:   2,  // B — employee ID
  COL_SCHEDULE_JSON: 3,  // C — JSON string of the 7-day schedule
  COL_TOTAL_HOURS:   4,  // D — total paid hours for the week
  COL_STORED_AT:     5,  // E — ISO timestamp of last write

  DAYS_IN_WEEK:      7,
  SCHEMA_VERSION:    2,
};

/**
 * Day names in column order (Monday first).
 * Used when matching staffing requirements and when computing dates per column.
 */
const DAY_NAMES_IN_ORDER = [
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday',
];


// ---------------------------------------------------------------------------
// Hour Rules
// ---------------------------------------------------------------------------

/**
 * Weekly paid hour minimums and maximums per employee scheduling type.
 *
 * FT:  40 hours exactly (minimum = maximum).
 * PT:  24–40 hours. Phase 2 adds shifts until the minimum is reached.
 * LPT: 0–12 hours. No minimum enforced; intended for students and
 *      weekend-only workers who cannot commit to standard PT hours.
 */
const HOUR_RULES = {
  FT_MIN:   40,
  FT_MAX:   40,
  PT_MIN:   24,
  PT_MAX:   40,
  LPT_MIN:   0,  // No weekly minimum — schedule as needed
  LPT_MAX:  12,  // Hard cap: never exceed 12 paid hours per week
};


// ---------------------------------------------------------------------------
// Schedule Rules
// ---------------------------------------------------------------------------

/**
 * Structural scheduling constraints applied across all phases.
 *
 * MIN_DAYS_OFF: every employee must have at least this many non-working days
 * (OFF, RDO, or VAC). Phase 1 trims excess SHIFT assignments; Phase 2 and
 * Phase 3 both treat this floor as a hard cap.
 */
const SCHEDULE_RULES = {
  MIN_DAYS_OFF: 2,
};


// ---------------------------------------------------------------------------
// Supervisor Scheduling
// ---------------------------------------------------------------------------

/**
 * Configuration for supervisor peak traffic scheduling.
 *
 * supervisorRole: Employee role identifier for supervisors (column L)
 * enabled: Whether supervisor scheduling is enabled (default false)
 * membersPerSupervisor: Ratio used to calculate required supervisors (e.g., 75 = 1 supervisor per 75 members)
 * maxDoorCount: Maximum door count on the Y-axis of the visualizer (soft ceiling; values can exceed but display clips)
 * defaultPeakProfile: Default daily peak traffic profile (8 elements, one per 3 hours: 0, 3, 6, 9, 12, 15, 18, 21)
 *                      Each element is a door count (members per hour), used as fallback if no per-dept config exists.
 *
 * Peak profiles are stored dynamically in the COMET Config sheet as:
 *   Key: SUPERVISOR_PEAK_WINDOWS_{DEPARTMENT}
 *   Value: JSON { department, peakProfile, enabled, membersPerSupervisor, maxDoorCount, lastUpdated }
 *
 * Each element in peakProfile[dayName] is an integer representing expected member foot traffic at that time slot.
 */
/**
 * DEPRECATED: Supervisor rules are now integrated into the Traffic Heatmap system.
 * This constant is kept for backwards compatibility during migration but should be
 * removed once all Phase 5 code is consolidated.
 */
const SUPERVISOR_RULES = {
  supervisorRole: 'Supervisor',
  enabled: false,
  membersPerSupervisor: 75,
  maxDoorCount: 500,
  defaultPeakProfile: {
    'Monday':    [0, 0, 0, 50, 150, 200, 175, 100],
    'Tuesday':   [0, 0, 0, 50, 150, 200, 175, 100],
    'Wednesday': [0, 0, 0, 50, 150, 200, 175, 100],
    'Thursday':  [0, 0, 0, 50, 150, 200, 175, 100],
    'Friday':    [0, 0, 0, 75, 200, 350, 400, 250],
    'Saturday':  [0, 0, 0, 75, 250, 400, 450, 300],
    'Sunday':    [0, 0, 0, 50, 200, 375, 400, 250],
  },
};


// ---------------------------------------------------------------------------
// Traffic Heatmap & Peak Coverage Scheduling
// ---------------------------------------------------------------------------

/**
 * Store closing times by day of week (used to enforce hard constraints on closer shifts).
 * Format: 24-hour time as "HH:MM" string.
 *
 * Closers' shift end times must not exceed these times:
 * - Weekday (Mon–Fri): 11:30 PM (23:30)
 * - Saturday: 10:00 PM (22:00)
 * - Sunday: 9:00 PM (21:00)
 *
 * Used in buildStaggeredStartMap_() to cap flexWindowLatest for closer shifts.
 */
const STORE_CLOSING_TIMES = {
  weekday: '23:30',  // Mon–Fri
  saturday: '22:00',
  sunday: '21:00',
};

/**
 * Default traffic heatmap configuration for a new department.
 * Managers can override these via the UI.
 */
const HEATMAP_DEFAULTS = {
  enabled: false,
  thresholds: {
    low: 100,
    high: 300,
  },
  staggerIncrement: 15,  // Minutes between staggered start times (15 or 30)
  levelMultipliers: {
    'Low': 0.75,
    'Moderate': 1.0,
    'High': 1.25,
  },
  poolSchedulingCounts: {
    'Low': 1,
    'Moderate': 3,
    'High': 5,
  },
  // Default traffic curves for new departments (can be overridden per week).
  // 24 values — one per hour, index 0 = midnight, index 9 = 9am, etc.
  // Based on observed member traffic at warehouse 1338:
  //   Weekdays open 9am–9pm; weekday peaks at 12pm, 3pm, 5pm (max), 7–8pm.
  //   Friday slightly busier. Saturday closes 7pm, Sunday closes 6pm;
  //   both weekends run consistent high traffic (~330) from 11am–5pm.
  defaultTrafficCurves: {
    //                       12a 1  2  3  4  5  6  7  8   9am 10  11  12p  1   2   3   4   5pm  6   7pm  8   9pm 10  11
    'Monday':    [0, 0, 0, 0, 0, 0, 0, 0, 0,  75, 100, 120, 180, 140, 120, 180, 160, 225, 170, 180, 150,   0, 0, 0],
    'Tuesday':   [0, 0, 0, 0, 0, 0, 0, 0, 0,  75, 100, 120, 180, 140, 120, 180, 160, 225, 170, 180, 150,   0, 0, 0],
    'Wednesday': [0, 0, 0, 0, 0, 0, 0, 0, 0,  75, 100, 120, 180, 140, 120, 180, 160, 225, 170, 180, 150,   0, 0, 0],
    'Thursday':  [0, 0, 0, 0, 0, 0, 0, 0, 0,  75, 100, 120, 180, 140, 120, 180, 160, 225, 170, 180, 150,   0, 0, 0],
    'Friday':    [0, 0, 0, 0, 0, 0, 0, 0, 0,  80, 110, 140, 200, 155, 140, 200, 180, 250, 195, 205, 170,   0, 0, 0],
    'Saturday':  [0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 200, 320, 330, 330, 320, 330, 320, 310, 200,   0,   0,   0, 0, 0],
    'Sunday':    [0, 0, 0, 0, 0, 0, 0, 0, 0, 100, 200, 310, 330, 325, 320, 320, 300, 250,   0,   0,   0,   0, 0, 0],
  },
};

/**
 * Extended shift definition column offsets for flex windows.
 * These columns follow SHIFT_TABLE_COLUMN in the E2:N50 range.
 *
 *   L (7) — flexEnabled         (boolean; if false, shift is fixed and stagger does not apply)
 *   M (8) — flexWindowEarliest  (string "HH:MM"; earliest valid start time)
 *   N (9) — flexWindowLatest    (string "HH:MM"; latest valid start time)
 *
 * Extended offsets (0-indexed, relative to the E2:N50 range start):
 */
const SHIFT_TABLE_FLEX_COLUMNS = {
  FLEX_ENABLED:         7,  // L
  FLEX_WINDOW_EARLIEST: 8,  // M
  FLEX_WINDOW_LATEST:   9,  // N
};


// ---------------------------------------------------------------------------
// Split-Shift (Multi-Department) Scheduling
// ---------------------------------------------------------------------------

/**
 * Configuration for cross-department employee scheduling.
 *
 * STORAGE_FORMAT: "simple_list"
 *   Secondary departments are stored as comma-separated strings in EMPLOYEE_COLUMN.SECONDARY_DEPARTMENTS.
 *   Format: "Front End,Receiving" (no spaces around commas; normalized department names).
 *   No target hours are specified; allocation is dynamic.
 *
 * ALLOCATION_STRATEGY: "dynamic"
 *   When generating a schedule for a department, the system queries all other departments
 *   (those listed in the employee's secondary departments) to find hours already assigned.
 *   Available budget = HOUR_RULES[status] - crossDeptHoursAlreadyScheduled
 *   This is enforced in Phase 2 (Minimum Hour Enforcement).
 *
 * VACATION_COORDINATION: "shared"
 *   Vacation dates are stored once per employee (EMPLOYEE_COLUMN.VACATION_DATES).
 *   All department schedules see the same vacation dates, so an employee's VAC day
 *   blocks work in all departments simultaneously.
 *   When a manager updates a VAC/RDO cell via updateCellOverride, the change is
 *   persisted to the Employees sheet (column K) so it's visible cross-dept.
 */
const SPLIT_SHIFT_CONFIG = {
  STORAGE_FORMAT: 'simple_list',
  ALLOCATION_STRATEGY: 'dynamic',
  VACATION_COORDINATION: 'shared',
};


// ---------------------------------------------------------------------------
// Seniority Rank Formula Constants
// ---------------------------------------------------------------------------

/**
 * Seniority rank formula encodes both employment status and tenure in one integer.
 *
 * FT employees receive FT_BASE; PT receive PT_BASE.
 * Days-from-hire (relative to REFERENCE_DATE_STRING) are added to the base,
 * so earlier hire dates produce higher ranks. The 100M gap ensures no PT
 * employee can ever out-rank an FT employee by tenure alone.
 */
const SENIORITY = {
  FT_BASE:                200000000,
  PT_BASE:                100000000,
  REFERENCE_DATE_STRING:  '2050-01-01',
};


// ---------------------------------------------------------------------------
// Coverage Slot Map Constants
// ---------------------------------------------------------------------------

/**
 * Defines the 30-minute coverage slot array used by Phases 3 of the engine.
 *
 * Models the day from 04:00 to 23:30 (39 half-hour slots).
 * A slot's index = Math.floor((minutesSinceMidnight - COVERAGE_START_MINUTE) / SLOT_DURATION_MINUTES)
 */
const COVERAGE = {
  SLOT_COUNT:             39,   // 04:00–23:30 in 30-minute windows
  COVERAGE_START_MINUTE:  240,  // 04:00 = 4 × 60
  SLOT_DURATION_MINUTES:  30,
};

/**
 * Required coverage windows per day of the week (minutes since midnight).
 * Phase 3 only fills gaps within these windows; slots outside are not penalized.
 */
const COVERAGE_WINDOW = {
  Monday:    { startMinute: 240, endMinute: 1410 },
  Tuesday:   { startMinute: 240, endMinute: 1410 },
  Wednesday: { startMinute: 240, endMinute: 1410 },
  Thursday:  { startMinute: 240, endMinute: 1410 },
  Friday:    { startMinute: 240, endMinute: 1410 },
  Saturday:  { startMinute: 240, endMinute: 1320 },
  Sunday:    { startMinute: 240, endMinute: 1260 },
};


// ---------------------------------------------------------------------------
// Colors
// ---------------------------------------------------------------------------

/**
 * Hex color codes applied to cells in generated schedule sheets.
 *
 * Blue  = FT shift  |  Green = PT shift  |  Yellow = Vacation
 * Gray  = Day off   |  Red   = Under hours  |  Orange = FT over-hours / combo shift
 * Lavender = Role row background
 */
const COLORS = {
  FT_SHIFT:      '#4A90D9',  // Blue — full-time shift background
  PT_SHIFT:      '#57BB8A',  // Green — part-time shift background
  COMBO_SHIFT:   '#FF7043',  // Deep orange — hybrid/combo shift (cross-dept handoff)
  VACATION:      '#FFD966',  // Yellow — vacation day background
  DAY_OFF:       '#B7B7B7',  // Gray — RDO/OFF background
  UNDER_HOURS:   '#E06666',  // Red — name cell when employee is below weekly minimum
  OVER_HOURS_FT: '#FF9900',  // Orange — total-hours cell when FT is above 40 hours
  HEADER_BG:     '#263238',  // Dark slate — column header row background
  HEADER_TEXT:   '#FFFFFF',  // White — column header text
  SUMMARY_OK:    '#B7E1CD',  // Light green — STATUS row when coverage is met
  SUMMARY_UNDER: '#F4C7C3',  // Light red — STATUS row when coverage is short
  ROW_LABEL_BG:  '#F5F5F5',  // Light gray — VAC/RDO/SHIFT label column background
  ROLE_ROW_BG:   '#EDE7F6',  // Lavender — ROLE row background (non-working days and label)
  POOL_SECTION_BG: '#F3E5F5',  // Light purple — pool member section background (traffic heatmap v0.5.0)
};

/**
 * Per-role background colors for the ROLE row of the generated schedule.
 * Keys must match the role strings stored in the Employees sheet column L.
 * Roles not found in this map receive the generic COLORS.ROLE_ROW_BG (lavender).
 */
const ROLE_COLORS = {
  // Front End — blue / cyan / purple family
  'Cashier':               '#DBEAFE',  // Blue
  'Assistant':             '#BFDBFE',  // Medium blue  (cashier assistant)
  'SCO':                   '#BAE6FD',  // Sky blue
  'Liquor':                '#A5F3FC',  // Cyan
  'PreScan':               '#EDE9FE',  // Purple
  'Go Backs':              '#F5D0FE',  // Fuchsia
  'Carts':                 '#CBD5E1',  // Slate
  'Floater':               '#FEF08A',  // Yellow

  // Merchandising — green / indigo family
  'Morning':               '#D1FAE5',  // Green
  'Night':                 '#C7D2FE',  // Indigo
  'Produce':               '#BBF7D0',  // Bright green
  'Stocker':               '#ECFDF5',  // Mint
  'Driver':                '#FDBA74',  // Orange

  // Bakery / Food — rose and warm yellow
  // Bakery uses rose so it reads distinctly alongside Maintenance amber in
  // the Bakery dept schedule when cross-trained Merch employees appear there.
  'Bakery':                '#FECACA',  // Rose
  'Deli':                  '#FDE68A',  // Warm yellow
  'Food Service':          '#FDE68A',  // Warm yellow (Food Court / Service Deli)
  'Food Court':            '#FDE68A',  // Warm yellow

  // Maintenance — amber
  'Maintenance Associate': '#FEF3C7',  // Amber
  'Maintenance':           '#FEF3C7',  // Amber (alias)

  // Other departments
  'Receiving':             '#A7F3D0',  // Teal green
  'Membership Service':    '#FFFBEB',  // Light warm
  'Security':              '#E2E8F0',  // Neutral gray-blue
  'Tire Tech':             '#F0FDF4',  // Light mint
  'Gas Attendant':         '#FFEDD5',  // Light orange
  'Pharmacy Tech':         '#F3E8FF',  // Light purple
  'Optician':              '#F0F9FF',  // Very light blue
  'Hearing Specialist':    '#FFF0F5',  // Very light pink

  // Leadership
  'Lead':                  '#FEE2E2',  // Light red
  'Supervisor':            '#FCE7F3',  // Pink
};

/**
 * Optional role ratio rules. Set to undefined to disable.
 *
 * Shape: { triggerRole: { requiresRole, ratio } }
 * Example: { 'Cashier': { requiresRole: 'SCO', ratio: 0.5 } }
 *   → for every 2 Cashiers working the same day, 1 SCO is required.
 *
 * Phase 4 of the engine checks this object. If undefined it is skipped.
 */
const ROLE_RULES = undefined;


// ---------------------------------------------------------------------------
// Infraction / CN System
// ---------------------------------------------------------------------------

/**
 * When true, the scanner logs proposals but sends no emails and writes nothing
 * to the CN_Log. Set to false to go live.
 */
const DRY_RUN = true;

/**
 * Number of days to look back when scanning for infraction windows.
 * 60 days covers two full 30-day windows and catches events near month boundaries.
 */
const DAYS_BACK = 60;

/** Default rolling window length (days) for codes without a CODE_RULES entry. */
const WINDOW_DAYS = 30;

/** Default trigger count for codes without a CODE_RULES entry. */
const THRESHOLD_COUNT = 3;

/**
 * Per-code infraction threshold and window overrides.
 * Derived from the Costco Employee Agreement (March 2025).
 */
const CODE_RULES = {
  TD: { threshold: 3, windowDays: 30  },  // Tardy (§11.4.2)
  NS: { threshold: 1, windowDays: 30  },  // No Show (§11.4.3e)
  SE: { threshold: 3, windowDays: 30  },  // Swiping Error (§11.4.12a)
  MP: { threshold: 3, windowDays: 30  },  // Meal Period Occurrence (§11.4.12b)
  SZ: { threshold: 3, windowDays: 365 },  // Suspension (§11.3.11a)
};

/** Codes that count as infractions when present in an employee's calendar. */
const INFRACTION_CODES = ['TD', 'NS', 'SE', 'MP', 'SZ'];

/**
 * Codes that are always ignored, even if they appear in INFRACTION_CODES.
 * Includes protected leave types and administrative codes.
 */
const IGNORE_CODES = ['BL', 'CN', 'FH', 'H', 'JD', 'SPF', 'SUF', 'SPH', 'SUH', 'LP', 'NY', 'FJ'];

/** Days after issuance before a CN is automatically marked Expired. */
const EXPIRY_DAYS = 180;

/**
 * Attendance controller calendar grid layout.
 * Three horizontal bands, each containing four side-by-side month blocks.
 */
const DATA_BANDS = [
  { monthRow: 5,  dayOfWeekRow: 6,  firstGridRow: 7,  lastGridRow: 30 },
  { monthRow: 32, dayOfWeekRow: 33, firstGridRow: 34, lastGridRow: 57 },
  { monthRow: 59, dayOfWeekRow: 60, firstGridRow: 61, lastGridRow: 83 },
];

/** Starting columns (A1-notation) for each of the four month blocks per band. */
const START_COLUMNS = ['A', 'I', 'Q', 'Y'];

/** Number of day-data columns per month block (the 8th is a visual separator). */
const DAY_COLS_PER_BLOCK = 7;

/**
 * Cell addresses for employee metadata on each individual attendance controller tab.
 * Standardized across Costco warehouses.
 */
const EMPLOYEE_FIELDS = {
  yearTitle:    'D1',
  employeeName: 'X1',
  department:   'R3',
  employeeId:   'X3',
  hireDate:     'AD3',
};

/**
 * Hidden template sheet used by the attendance controller tab generator.
 * Built once per year with all formatting; each employee tab is a copy of this sheet.
 * The leading underscore sorts it before employee tabs and signals it is system-managed.
 */
const ATTENDANCE_TEMPLATE_SHEET_NAME = '_AC_TEMPLATE_';

/**
 * Regex that identifies employee tabs in the attendance controller.
 * Format: "Last, First - EmployeeNumber"  e.g. "Le, Tony - 1234578"
 */
const EMPLOYEE_TAB_PATTERN = /^.+,\s*.+\s*-\s*\d+$/;

/** CN_Log sheet tab name. */
const CN_LOG_SHEET_NAME = 'CN_Log';

/** Active counseling notices sheet tab name. */
const ACTIVE_CNS_SHEET_NAME = 'Active CNs';

/** Archived (expired) CNs sheet tab name. */
const EXPIRED_CNS_SHEET_NAME = '(Expired CNs)';

/** Column headers for the CN_Log sheet. */
const CN_LOG_HEADERS = [
  'CN_Key', 'EmployeeID', 'EmployeeName', 'Department',
  'WindowStart', 'WindowEnd', 'Count', 'EventsHash',
  'IssuedAt', 'IssuedBy', 'DryRun', 'SheetName',
  'Status', 'ExpiredAt', 'Rule',
  'SourceSpreadsheetId', 'SourceSheetGid',
];

/** Column headers for the Active CNs sheet. */
const ACTIVE_CNS_HEADERS = [
  'CN_Key', 'Employee Name', 'Employee ID', 'Department',
  'Rule', 'Count', 'Window Start', 'Window End', 'Issued At', 'Sheet',
];

/** Column headers for the (Expired CNs) sheet. */
const EXPIRED_CNS_HEADERS = [
  'CN_Key', 'Employee Name', 'Employee ID', 'Department',
  'Rule', 'Count', 'Window Start', 'Window End', 'Issued At', 'Sheet',
  'Expired At',
];

/** Name of the configuration sheet that holds the CN Log spreadsheet ID. */
const INFRACTION_CONFIG_SHEET_NAME = 'Infraction Config';

/** Cell on INFRACTION_CONFIG_SHEET_NAME where the CN Log spreadsheet ID is entered. */
const LOG_SPREADSHEET_ID_CELL = 'B2';


// ---------------------------------------------------------------------------
// Absence Log (Call Log)
// ---------------------------------------------------------------------------

/**
 * Column positions (0-indexed) for the Call Log sheet.
 *
 *   A (0)  — Employee Name
 *   B (1)  — Employee ID
 *   C (2)  — (reserved)
 *   D (3)  — Is Callout       (checkbox)
 *   E (4)  — (reserved)
 *   F (5)  — Is FMLA          (checkbox)
 *   G (6)  — Is No Show       (checkbox)
 *   H (7)  — Department
 *   I (8)  — Time Called
 *   J (9)  — Manager (who took the call)
 *   K (10) — Scheduled Shift
 *   L-M (11-12) — (reserved)
 *   N (13) — Comment
 *   O (14) — Date
 */
const CALL_LOG_COLUMN = {
  NAME:             0,
  EMPLOYEE_ID:      1,
  IS_CALLOUT:       3,
  IS_FMLA:          5,
  IS_NOSHOW:        6,
  DEPARTMENT:       7,
  TIME:             8,
  MANAGER:          9,
  SCHEDULED_SHIFT:  10,
  COMMENT:          13,
  DATE:             14,
};

/** Row number where Call Log data begins (row 1 = title, row 2 = headers, row 3 = data). */
const CALL_LOG_DATA_START_ROW = 3;

/**
 * Google Sheets uses a serial date where Dec 30, 1899 = 0.
 * Unix epoch (Jan 1, 1970) is 25569 days after that start.
 * Used when converting a raw numeric cell value to a JS Date.
 */
const SHEETS_EPOCH_OFFSET = 25569;


// ---------------------------------------------------------------------------
// Notification Recipients
// ---------------------------------------------------------------------------

/**
 * Department-to-recipient email mapping for absence notifications.
 * Keys should match department names as they appear in the Employees sheet.
 * Update to real email addresses before go-live.
 */
const MAILING_LIST = {
  'Maintenance':  ['maintenance.manager@example.com'],
  'Night Merch':  ['nightmerch.manager@example.com'],
  'Front End':    ['frontend.manager@example.com'],
};

/** Fallback recipient used when a department has no MAILING_LIST entry. */
const FALLBACK_EMAIL = 'gm@example.com';

/**
 * Recipients for all CN notifications (new CNs and expiry notices).
 * Update to real payroll email(s) before go-live.
 */
const PAYROLL_RECIPIENTS = [
  'payroll.clerk@example.com',
];


// ---------------------------------------------------------------------------
// Performance Tuning
// ---------------------------------------------------------------------------

/**
 * Enable verbose profiling output in the console. Set to true to measure
 * execution times and debug performance bottlenecks. Should be false in production.
 */
const PROFILING_ENABLED = true;

/**
 * Batch size for sheet operations. When writing large blocks of data,
 * batching in groups of this size reduces API overhead.
 */
const BATCH_SIZE_SHEET_WRITE = 50;

/**
 * Cache time-to-live in minutes. Cached values (e.g., shift definitions,
 * attendance grids) expire after this interval.
 */
const CACHE_TTL_MINUTES = 30;

/**
 * Maximum execution time in milliseconds before the API layer should warn
 * or abort an operation. Default 4.5 minutes (270,000 ms), leaving a
 * 1.5-minute buffer before the GAS 6-minute hard timeout.
 */
const MAX_SAFE_EXECUTION_MS = 270000;
