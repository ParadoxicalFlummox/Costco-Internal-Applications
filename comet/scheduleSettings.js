/**
 * scheduleSettings.js — Per-department schedule settings read/write for COMET.
 * VERSION: 0.5.0
 *
 * Each department has its own Settings sheet tab named "Settings_[DeptName]"
 * (e.g., "Settings_Maintenance"). This file owns all reads and writes to those sheets.
 *
 * SHEET LAYOUT (Settings_[Dept]):
 *   A2:C8   — Staffing requirements: Day | Count | Mode ("Count" or "Hours")
 *   E2:M50  — Shift definitions:
 *             E: Name | F: FT/PT | G: StartTime | H: EndTime | I: PaidHours | J: HasLunch |
 *             K: FlexEnabled | L: FlexWindowEarliest | M: FlexWindowLatest
 *
 * RETURN SHAPES:
 *   getDeptSettings_() returns:
 *   {
 *     staffingReqs: [{ day, count, mode }],     // 7 entries, Mon–Sun
 *     shifts:       [{ name, ftpt, startTime, endTime, paidHours, hasLunch,
 *                      flexEnabled, flexWindowEarliest, flexWindowLatest }],
 *   }
 *
 *   Time values in the shifts array are stored as "HH:MM" strings (24-hour)
 *   for safe serialization over google.script.run. The engine's settingsManager.js
 *   converts them back to minutes-since-midnight.
 */


// ---------------------------------------------------------------------------
// Public API (called from api.js)
// ---------------------------------------------------------------------------

/**
 * Returns the settings for a single department.
 * Creates the Settings sheet with defaults if it doesn't exist yet.
 *
 * @param {string} deptName
 * @returns {{ staffingReqs: Array, shifts: Array }}
 */
function getDeptSettings_(deptName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDeptSettingsSheet_(workbook, deptName);
  return readDeptSettingsFromSheet_(sheet);
}

/**
 * Writes shift definitions and staffing requirements for a department back to its Settings sheet.
 *
 * @param {string} deptName
 * @param {{
 *   staffingReqs: Array<{ day: string, count: number, mode: string }>,
 *   shifts:       Array<{ name: string, ftpt: string, startTime: string, endTime: string,
 *                          paidHours: number, hasLunch: boolean }>,
 * }} data
 * @returns {{ saved: boolean }}
 */
function saveDeptSettings_(deptName, data) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDeptSettingsSheet_(workbook, deptName);

  writeStaffingRequirements_(sheet, data.staffingReqs || []);
  writeShiftDefinitions_(sheet, data.shifts || []);

  SpreadsheetApp.flush();
  return { saved: true };
}


// ---------------------------------------------------------------------------
// Sheet Bootstrap
// ---------------------------------------------------------------------------

/**
 * Returns the Settings sheet for the given department, creating it with defaults
 * if it does not yet exist.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @param {string} deptName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateDeptSettingsSheet_(workbook, deptName) {
  const sheetName = DEPT_SETTINGS_PREFIX + deptName; // config.js
  let sheet = workbook.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = workbook.insertSheet(sheetName);
  writeDefaultDeptSettings_(sheet, deptName);
  return sheet;
}

/**
 * Writes column headers and default data to a freshly created Settings sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} deptName
 */
function writeDefaultDeptSettings_(sheet, deptName) {
  // --- Header row (row 1) ---
  const headerRange = sheet.getRange(1, 1, 1, 9);
  headerRange.setValues([[
    'Day', 'Count', 'Mode', '',
    'Shift Name', 'FT/PT', 'Start Time', 'End Time', 'Paid Hours',
  ]]);
  headerRange.setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');

  // Add shift-related headers in columns J–M
  const shiftHeadersRange = sheet.getRange(1, 10, 1, 4);  // J1:M1
  shiftHeadersRange.setValues([[
    'Has Lunch', 'Flex Enabled', 'Flex Earliest', 'Flex Latest',
  ]]).setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');

  // --- Default staffing requirements (A2:C8) ---
  const defaultStaffing = DAY_NAMES_IN_ORDER.map(day => [day, DEFAULT_STAFFING_COUNT, STAFFING_MODE.COUNT]); // config.js
  sheet.getRange(2, 1, defaultStaffing.length, 3).setValues(defaultStaffing);

  // --- Default shift definitions (E2:M3) — two example rows ---
  // Include flex window fields: [Name, FT/PT, StartTime, EndTime, PaidHours, HasLunch, FlexEnabled, FlexEarliest, FlexLatest]
  const defaultShifts = [
    ['Morning', 'FT', '08:00', '16:30', 8, true, true, '07:30', '09:00'],
    ['Morning', 'PT', '08:00', '13:00', 5, false, true, '07:30', '09:00'],
  ];
  sheet.getRange(2, 5, defaultShifts.length, 9).setValues(defaultShifts);

  sheet.setTabColor('#005DAA');
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 110);  // Day
  sheet.setColumnWidth(2, 80);   // Count
  sheet.setColumnWidth(3, 80);   // Mode
  sheet.setColumnWidth(4, 20);   // spacer
  sheet.setColumnWidth(5, 140);  // Shift Name
  sheet.setColumnWidth(6, 70);   // FT/PT
  sheet.setColumnWidth(7, 90);   // Start Time
  sheet.setColumnWidth(8, 90);   // End Time
  sheet.setColumnWidth(9, 90);   // Paid Hours
  sheet.setColumnWidth(10, 90);  // Has Lunch
  sheet.setColumnWidth(11, 100); // Flex Enabled
  sheet.setColumnWidth(12, 110); // Flex Earliest
  sheet.setColumnWidth(13, 110); // Flex Latest
}


// ---------------------------------------------------------------------------
// Sheet Readers
// ---------------------------------------------------------------------------

/**
 * Reads a Settings sheet and returns the settings as a plain serializable object.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{ staffingReqs: Array, shifts: Array }}
 */
function readDeptSettingsFromSheet_(sheet) {
  return {
    staffingReqs: readStaffingRequirements_(sheet),
    shifts:       readShiftDefinitions_(sheet),
  };
}

/**
 * Reads the staffing requirements table from A2:C8.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{ day: string, count: number, mode: string }>}
 */
function readStaffingRequirements_(sheet) {
  const rows = sheet.getRange(SETTINGS_RANGE.STAFFING_REQUIREMENTS_TABLE).getValues(); // config.js
  return rows
    .filter(row => row[0])
    .map(row => ({
      day:   String(row[0] || '').trim(),
      count: Number(row[1] || 0),
      mode:  String(row[2] || STAFFING_MODE.COUNT).trim(),
    }));
}

/**
 * Reads the shift definitions table from E2:M50 (extended to include flex window fields).
 * Times are returned as "HH:MM" strings (safe for google.script.run serialization).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{ name, ftpt, startTime, endTime, paidHours, hasLunch, flexEnabled, flexWindowEarliest, flexWindowLatest }>}
 */
function readShiftDefinitions_(sheet) {
  // Note: We need to read M50 to capture all flex columns. Update the range.
  const rows = sheet.getRange('E2:M50').getValues();
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  return rows
    .filter(row => row[SHIFT_TABLE_COLUMN.NAME] && row[SHIFT_TABLE_COLUMN.STATUS])
    .map(row => {
      const startRaw = row[SHIFT_TABLE_COLUMN.START_TIME];
      const endRaw   = row[SHIFT_TABLE_COLUMN.END_TIME];
      const flexEarliestRaw = row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_WINDOW_EARLIEST];
      const flexLatestRaw = row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_WINDOW_LATEST];

      return {
        name:               String(row[SHIFT_TABLE_COLUMN.NAME]      || '').trim(),
        ftpt:               String(row[SHIFT_TABLE_COLUMN.STATUS]     || '').trim(),
        startTime:          formatGasTimeToString_(startRaw, timeZone),
        endTime:            formatGasTimeToString_(endRaw,   timeZone),
        paidHours:          Number(row[SHIFT_TABLE_COLUMN.PAID_HOURS] || 0),
        hasLunch:           row[SHIFT_TABLE_COLUMN.HAS_LUNCH] === true,
        flexEnabled:        row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_ENABLED] !== false,  // Default true if not explicitly false
        flexWindowEarliest: formatGasTimeToString_(flexEarliestRaw, timeZone) || '',
        flexWindowLatest:   formatGasTimeToString_(flexLatestRaw, timeZone) || '',
      };
    });
}


// ---------------------------------------------------------------------------
// Sheet Writers
// ---------------------------------------------------------------------------

/**
 * Writes the staffing requirements rows back to A2:C8.
 * Fills all 7 day rows; any days not in the input get count=0.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<{ day, count, mode }>} staffingReqs
 */
function writeStaffingRequirements_(sheet, staffingReqs) {
  const map = {};
  staffingReqs.forEach(r => { map[r.day] = r; });

  const rows = DAY_NAMES_IN_ORDER.map(day => { // config.js
    const entry = map[day] || {};
    return [day, Number(entry.count || 0), entry.mode || STAFFING_MODE.COUNT];
  });

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // Clear any leftover rows below (in case someone had more than 7)
  const lastRow = sheet.getLastRow();
  if (lastRow > 8) {
    sheet.getRange(9, 1, lastRow - 8, 3).clearContent();
  }
}

/**
 * Writes the shift definitions table starting at E2 through M50 (includes flex fields).
 * Clears old rows below the new data to avoid stale entries.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<{ name, ftpt, startTime, endTime, paidHours, hasLunch, flexEnabled, flexWindowEarliest, flexWindowLatest }>} shifts
 */
function writeShiftDefinitions_(sheet, shifts) {
  // Clear the entire shift table area first (E2:M50)
  sheet.getRange('E2:M50').clearContent();

  if (shifts.length === 0) return;

  const rows = shifts.map(s => [
    s.name,
    s.ftpt,
    s.startTime,          // written as "HH:MM" string; sheet displays it as text
    s.endTime,
    Number(s.paidHours || 0),
    s.hasLunch === true,
    s.flexEnabled !== false,  // Default true if not explicitly false
    s.flexWindowEarliest || '',
    s.flexWindowLatest   || '',
  ]);

  sheet.getRange(2, 5, rows.length, 9).setValues(rows);  // E2 onwards, 9 columns
}


// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Converts a GAS time cell value (Date object or numeric fraction) to "HH:MM" string.
 * Returns the value as-is if it is already a string.
 *
 * @param {any} cellValue
 * @param {string} timeZone
 * @returns {string}
 */
function formatGasTimeToString_(cellValue, timeZone) {
  if (!cellValue && cellValue !== 0) return '';
  if (typeof cellValue === 'string') return cellValue.trim();
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, timeZone, 'HH:mm');
  }
  if (typeof cellValue === 'number') {
    const totalMinutes = Math.round(cellValue * 1440);
    const hours   = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const minutes = String(totalMinutes % 60).padStart(2, '0');
    return hours + ':' + minutes;
  }
  return String(cellValue);
}

/**
 * Converts a "HH:MM" string (as stored in the settings sheet) to minutes since midnight.
 * Used by scheduleEngine.js when loading settings for generation.
 *
 * @param {string} timeString — e.g. "08:00" or "16:30"
 * @returns {number} minutes since midnight, or 0 if unparseable
 */
function timeStringToMinutes_(timeString) {
  if (!timeString) return 0;
  const parts = String(timeString).split(':');
  if (parts.length < 2) return 0;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}
