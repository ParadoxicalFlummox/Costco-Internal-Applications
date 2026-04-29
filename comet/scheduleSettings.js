/**
 * scheduleSettings.js — Per-department schedule settings read/write for COMET.
 * VERSION: 0.6.0
 *
 * Each department has its own Settings sheet tab named "Settings_[DeptName]"
 * (e.g., "Settings_Maintenance"). This file owns all reads and writes to those sheets.
 *
 * SHEET LAYOUT (Settings_[Dept]):
 *   A1:D1   — Section header row (Day | Count | Mode | spacer)
 *   A2:C8   — Staffing requirements: Day | Count | Mode ("Count" or "Hours")
 *   A10:B10 — ENGINE OPTIONS section header (merged)
 *   A11:B12 — Engine option rows: Label | TRUE/FALSE
 *               Row 11: Enforce Role Minimums
 *               Row 12: Enable Gap Fill
 *   A14:D14 — ROLE MINIMUMS section header (merged)
 *   A15:D15 — Column labels: Role | Low | Moderate | High
 *   A16:D*  — Role minimum rows: one per role
 *   E1:N1   — Shift definitions header
 *   E2:N50  — Shift definitions:
 *             E: Name | F: FT/PT | G: WkdyStart | H: SatStart | I: SunStart |
 *             J: PaidHours | K: HasLunch | L: FlexEnabled | M: FlexWindowEarliest | N: FlexWindowLatest
 *
 * RETURN SHAPES:
 *   getDeptSettings_() returns:
 *   {
 *     staffingReqs:  [{ day, count, mode }],        // 7 entries, Mon–Sun
 *     shifts:        [{ name, ftpt, weekdayStart, satStart, sunStart, paidHours, hasLunch,
 *                       flexEnabled, flexWindowEarliest, flexWindowLatest }],
 *     engineOptions: { enforceRoleMinimums: bool, gapFillEnabled: bool },
 *     roleMinimums:  { RoleName: { Low: n, Moderate: n, High: n }, ... },
 *   }
 *
 *   Time values in the shifts array are stored as "HH:MM" strings (24-hour)
 *   for safe serialization over google.script.run. The engine's settingsManager.js
 *   converts them back to minutes-since-midnight via getStartMinutesForDay_().
 */


// ---------------------------------------------------------------------------
// Public API (called from api.js)
// ---------------------------------------------------------------------------

/**
 * Returns the settings for a single department.
 * Creates the Settings sheet with defaults if it doesn't exist yet.
 *
 * @param {string} deptName
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object }}
 */
function getDeptSettings_(deptName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDeptSettingsSheet_(workbook, deptName);
  return readDeptSettingsFromSheet_(sheet);
}

/**
 * Writes shift definitions, staffing requirements, engine options, and role minimums
 * for a department back to its Settings sheet.
 *
 * @param {string} deptName
 * @param {{
 *   staffingReqs:  Array<{ day: string, count: number, mode: string }>,
 *   shifts:        Array<{ name: string, ftpt: string, weekdayStart: string,
 *                           paidHours: number, hasLunch: boolean }>,
 *   engineOptions: { enforceRoleMinimums: boolean, gapFillEnabled: boolean },
 *   roleMinimums:  { [roleName: string]: { Low: number, Moderate: number, High: number } },
 * }} data
 * @returns {{ saved: boolean }}
 */
function saveDeptSettings_(deptName, data) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDeptSettingsSheet_(workbook, deptName);

  writeStaffingRequirements_(sheet, data.staffingReqs || []);
  // Only touch shift definitions when the caller explicitly provides them.
  // Omitting shifts (e.g. from the pre-gen modal) must NOT clear the shift table.
  if (data.shifts !== undefined && data.shifts !== null) {
    writeShiftDefinitions_(sheet, data.shifts);
  }
  if (data.engineOptions !== undefined && data.engineOptions !== null) {
    writeEngineOptions_(sheet, data.engineOptions);
  }
  if (data.roleMinimums !== undefined && data.roleMinimums !== null) {
    writeRoleMinimums_(sheet, data.roleMinimums);
  }

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
 * @param {string} _deptName — unused; reserved for future dept-specific defaults
 */
function writeDefaultDeptSettings_(sheet, _deptName) {
  // --- Header row (row 1): staffing columns A–C + spacer D + shift columns E–N ---
  const staffingHeaderRange = sheet.getRange(1, 1, 1, 4);
  staffingHeaderRange.setValues([['Day', 'Count', 'Mode', '']]);
  staffingHeaderRange.setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');

  const shiftHeaderRange = sheet.getRange(1, 5, 1, 10);  // E1:N1
  shiftHeaderRange.setValues([[
    'Shift Name', 'FT/PT', 'Wkdy Start', 'Sat Start', 'Sun Start',
    'Paid Hours', 'Has Lunch', 'Flex Enabled', 'Flex Earliest', 'Flex Latest',
  ]]);
  shiftHeaderRange.setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');

  // --- Default staffing requirements (A2:C8) ---
  const defaultStaffing = DAY_NAMES_IN_ORDER.map(day => [day, DEFAULT_STAFFING_COUNT, STAFFING_MODE.COUNT]); // config.js
  sheet.getRange(2, 1, defaultStaffing.length, 3).setValues(defaultStaffing);

  // --- Default shift definitions (E2:N3) — two example rows ---
  // [Name, FT/PT, WkdyStart, SatStart, SunStart, PaidHours, HasLunch, FlexEnabled, FlexEarliest, FlexLatest]
  const defaultShifts = [
    ['Morning', 'FT', '08:00', '', '', 8, true,  true, '07:30', '09:00'],
    ['Morning', 'PT', '08:00', '', '', 5, false, true, '07:30', '09:00'],
  ];
  sheet.getRange(2, 5, defaultShifts.length, 10).setValues(defaultShifts);

  // --- Engine Options section (rows 10–12, cols A–B) ---
  writeEngineOptionsSectionHeader_(sheet);
  writeEngineOptions_(sheet, { enforceRoleMinimums: true, gapFillEnabled: true });

  // --- Role Minimums section (rows 14+, cols A–D) ---
  writeRoleMinimumsSectionHeader_(sheet);
  // No default role rows — managers add these via the UI.

  sheet.setTabColor('#005DAA');
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 140);  // Day / Role
  sheet.setColumnWidth(2, 80);   // Count / Low
  sheet.setColumnWidth(3, 90);   // Mode / Moderate
  sheet.setColumnWidth(4, 70);   // spacer / High
  sheet.setColumnWidth(5, 140);  // Shift Name
  sheet.setColumnWidth(6, 70);   // FT/PT
  sheet.setColumnWidth(7, 90);   // Wkdy Start
  sheet.setColumnWidth(8, 90);   // Sat Start
  sheet.setColumnWidth(9, 90);   // Sun Start
  sheet.setColumnWidth(10, 90);  // Paid Hours
  sheet.setColumnWidth(11, 90);  // Has Lunch
  sheet.setColumnWidth(12, 100); // Flex Enabled
  sheet.setColumnWidth(13, 110); // Flex Earliest
  sheet.setColumnWidth(14, 110); // Flex Latest
}


// ---------------------------------------------------------------------------
// Sheet Readers
// ---------------------------------------------------------------------------

/**
 * Reads a Settings sheet and returns the settings as a plain serializable object.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object }}
 */
function readDeptSettingsFromSheet_(sheet) {
  console.log('readDeptSettingsFromSheet_: reading staffing requirements...');
  const staffingReqs = readStaffingRequirements_(sheet);
  console.log('readDeptSettingsFromSheet_: staffingReqs = ' + staffingReqs.length + ' rows');

  console.log('readDeptSettingsFromSheet_: reading shift definitions...');
  const shifts = readShiftDefinitions_(sheet);
  console.log('readDeptSettingsFromSheet_: shifts = ' + shifts.length + ' rows');

  const engineOptions = readEngineOptions_(sheet);
  const roleMinimums  = readRoleMinimums_(sheet);

  return {
    staffingReqs:  staffingReqs,
    shifts:        shifts,
    engineOptions: engineOptions,
    roleMinimums:  roleMinimums,
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
 * Reads the shift definitions table from E2:N50.
 * Times are returned as "HH:MM" strings (safe for google.script.run serialization).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Array<{ name, ftpt, weekdayStart, satStart, sunStart, paidHours, hasLunch,
 *                   flexEnabled, flexWindowEarliest, flexWindowLatest }>}
 */
function readShiftDefinitions_(sheet) {
  try {
    console.log('readShiftDefinitions_: attempting to read E2:N50...');
    const rows = sheet.getRange('E2:N50').getValues();
    console.log('readShiftDefinitions_: got ' + rows.length + ' rows from range');

    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    const filtered = rows.filter(row => row[SHIFT_TABLE_COLUMN.NAME] && row[SHIFT_TABLE_COLUMN.STATUS]);
    console.log('readShiftDefinitions_: ' + filtered.length + ' rows have name + ftpt');

    const mapped = filtered.map(row => {
      const weekdayStartRaw = row[SHIFT_TABLE_COLUMN.WEEKDAY_START];
      const satStartRaw     = row[SHIFT_TABLE_COLUMN.SAT_START];
      const sunStartRaw     = row[SHIFT_TABLE_COLUMN.SUN_START];
      const flexEnabledRaw  = row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_ENABLED];

      return {
        name:               String(row[SHIFT_TABLE_COLUMN.NAME]      || '').trim(),
        ftpt:               String(row[SHIFT_TABLE_COLUMN.STATUS]     || '').trim(),
        weekdayStart:       formatGasTimeToString_(weekdayStartRaw, timeZone) || '',
        satStart:           formatGasTimeToString_(satStartRaw, timeZone)     || '',
        sunStart:           formatGasTimeToString_(sunStartRaw, timeZone)     || '',
        paidHours:          Number(row[SHIFT_TABLE_COLUMN.PAID_HOURS] || 0),
        hasLunch:           row[SHIFT_TABLE_COLUMN.HAS_LUNCH] === true,
        flexEnabled:        flexEnabledRaw !== false,  // Default true if not explicitly false (handles undefined)
        flexWindowEarliest: formatGasTimeToString_(row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_WINDOW_EARLIEST], timeZone) || '',
        flexWindowLatest:   formatGasTimeToString_(row[SHIFT_TABLE_FLEX_COLUMNS.FLEX_WINDOW_LATEST],   timeZone) || '',
      };
    });

    console.log('readShiftDefinitions_: mapped to ' + mapped.length + ' shift objects');
    return mapped;
  } catch (error) {
    console.log('ERROR in readShiftDefinitions_: ' + error.toString());
    if (error.stack) console.log('Stack: ' + error.stack);
    return [];
  }
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

  // Staffing requirements are always exactly 7 rows (Mon–Sun). Write them directly.
  // Do NOT clear below row 8 — the engine options and role minimums sections live there.
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * Writes the shift definitions table starting at E2 through N50 (includes per-day anchors + flex fields).
 * Clears old rows below the new data to avoid stale entries.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<{ name, ftpt, weekdayStart, satStart, sunStart, paidHours, hasLunch,
 *                 flexEnabled, flexWindowEarliest, flexWindowLatest }>} shifts
 */
function writeShiftDefinitions_(sheet, shifts) {
  // Clear the entire shift table area first (E2:N50)
  sheet.getRange('E2:N50').clearContent();

  if (shifts.length === 0) return;

  const rows = shifts.map(s => [
    s.name,
    s.ftpt,
    s.weekdayStart       || '',  // "HH:MM" string; Mon–Fri anchor
    s.satStart           || '',  // Saturday override (blank = use weekdayStart)
    s.sunStart           || '',  // Sunday override   (blank = use weekdayStart)
    Number(s.paidHours   || 0),
    s.hasLunch           === true,
    s.flexEnabled        !== false,
    s.flexWindowEarliest || '',
    s.flexWindowLatest   || '',
  ]);

  sheet.getRange(2, 5, rows.length, 10).setValues(rows);  // E2 onwards, 10 columns
}


// ---------------------------------------------------------------------------
// Engine Options — Read / Write
// ---------------------------------------------------------------------------

/**
 * Reads the engine options block from rows 11–12, cols A–B.
 * Defaults to all-enabled if rows are missing or values are not boolean.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{ enforceRoleMinimums: boolean, gapFillEnabled: boolean }}
 */
function readEngineOptions_(sheet) {
  try {
    const startRow = SETTINGS_ROWS.ENGINE_OPTIONS_START; // config.js
    const values   = sheet.getRange(startRow, 1, SETTINGS_ROWS.ENGINE_OPTIONS_COUNT, 2).getValues();
    // Row 0 = enforceRoleMinimums, Row 1 = gapFillEnabled
    // Column B (index 1) holds the boolean value.
    const enforceRoleMinimums = values[0] ? values[0][1] !== false : true;
    const gapFillEnabled      = values[1] ? values[1][1] !== false : true;
    return { enforceRoleMinimums: enforceRoleMinimums, gapFillEnabled: gapFillEnabled };
  } catch (_error) {
    return { enforceRoleMinimums: true, gapFillEnabled: true };
  }
}

/**
 * Writes the engine options block to rows 11–12, col B (preserves labels in col A).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{ enforceRoleMinimums: boolean, gapFillEnabled: boolean }} options
 */
function writeEngineOptions_(sheet, options) {
  const startRow = SETTINGS_ROWS.ENGINE_OPTIONS_START;
  sheet.getRange(startRow, 1, 2, 2).setValues([
    ['Enforce Role Minimums', options.enforceRoleMinimums !== false],
    ['Enable Gap Fill',       options.gapFillEnabled      !== false],
  ]);
}

/**
 * Writes the bold blue section header for the engine options block (row 10, cols A–B merged).
 * Only called during initial sheet creation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function writeEngineOptionsSectionHeader_(sheet) {
  const headerRow = SETTINGS_ROWS.ENGINE_OPTIONS_HEADER;
  const headerRange = sheet.getRange(headerRow, 1, 1, 2);
  headerRange.merge();
  headerRange.setValue('ENGINE OPTIONS');
  headerRange.setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');
  // Blank spacer above (row 9)
  sheet.getRange(9, 1, 1, 4).clearContent();
  // Blank spacer between options and role minimums (row 13)
  sheet.getRange(13, 1, 1, 4).clearContent();
}


// ---------------------------------------------------------------------------
// Role Minimums — Read / Write
// ---------------------------------------------------------------------------

/**
 * Reads the role minimums table from row ROLE_MINIMUMS_START downward.
 * Returns an empty object if no roles are configured.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{ [roleName: string]: { Low: number, Moderate: number, High: number } }}
 */
function readRoleMinimums_(sheet) {
  try {
    const startRow = SETTINGS_ROWS.ROLE_MINIMUMS_START;
    const lastRow  = sheet.getLastRow();
    if (lastRow < startRow) return {};

    const numRows = lastRow - startRow + 1;
    const values  = sheet.getRange(startRow, 1, numRows, 4).getValues();
    const result  = {};

    values.forEach(function(row) {
      const roleName = String(row[0] || '').trim();
      if (!roleName) return;
      result[roleName] = {
        Low:      Number(row[1] || 0),
        Moderate: Number(row[2] || 0),
        High:     Number(row[3] || 0),
      };
    });

    return result;
  } catch (_error) {
    return {};
  }
}

/**
 * Writes the role minimums table starting at ROLE_MINIMUMS_START, clearing stale rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {{ [roleName: string]: { Low: number, Moderate: number, High: number } }} roleMinimums
 */
function writeRoleMinimums_(sheet, roleMinimums) {
  const startRow = SETTINGS_ROWS.ROLE_MINIMUMS_START;
  const lastRow  = sheet.getLastRow();

  // Clear everything from startRow downward in cols A–D (stale role rows).
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, 4).clearContent();
  }

  const roleNames = Object.keys(roleMinimums);
  if (roleNames.length === 0) return;

  const rows = roleNames.map(function(roleName) {
    const entry = roleMinimums[roleName] || {};
    return [roleName, Number(entry.Low || 0), Number(entry.Moderate || 0), Number(entry.High || 0)];
  });

  sheet.getRange(startRow, 1, rows.length, 4).setValues(rows);
}

/**
 * Writes the bold blue section header and column labels for the role minimums block.
 * Only called during initial sheet creation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function writeRoleMinimumsSectionHeader_(sheet) {
  const headerRow = SETTINGS_ROWS.ROLE_MINIMUMS_HEADER;
  const labelRow  = SETTINGS_ROWS.ROLE_MINIMUMS_LABELS;

  const sectionHeaderRange = sheet.getRange(headerRow, 1, 1, 4);
  sectionHeaderRange.merge();
  sectionHeaderRange.setValue('ROLE MINIMUMS');
  sectionHeaderRange.setFontWeight('bold').setBackground('#005DAA').setFontColor('#FFFFFF');

  const columnLabelRange = sheet.getRange(labelRow, 1, 1, 4);
  columnLabelRange.setValues([['Role', 'Low', 'Moderate', 'High']]);
  columnLabelRange.setFontWeight('bold').setBackground('#D9E8F5').setFontColor('#003366');
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
