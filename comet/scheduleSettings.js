/**
 * scheduleSettings.js — Department settings read/write for COMET.
 * VERSION: 0.9.1
 *
 * STORAGE ARCHITECTURE:
 *   All department settings are stored in a single consolidated "Settings" sheet.
 *   Each row maps a department name (key) to its complete settings JSON (value).
 *
 * SHEET LAYOUT (Settings sheet):
 *   A1      — "DEPARTMENT" (header)
 *   B1      — "SETTINGS_JSON" (header) + warning to not modify this column
 *   A2+     — Department names (e.g., "Maintenance", "Cashiers", "Receiving")
 *   B2+     — Complete settings JSON for that department
 *
 * JSON SCHEMA (value in B2, B3, etc.):
 *   {
 *     staffingReqs:  [{ day: string, count: number, mode: string }, ...],     // 7 entries
 *     shifts:        [{ name: string, ftpt: string, weekdayStart: string,      // unlimited rows
 *                       satStart?: string, sunStart?: string, paidHours: number,
 *                       hasLunch: boolean, flexEnabled: boolean,
 *                       flexWindowEarliest?: string, flexWindowLatest?: string }, ...],
 *     engineOptions: { enforceRoleMinimums: boolean, gapFillEnabled: boolean },
 *     roleMinimums:  { [roleName: string]: { Low: number, Moderate: number, High: number } },
 *     roles:         [{ name: string, isPoolRole: boolean }, ...],             // unlimited
 *     employeeRoles: { [employeeId: string]: string },                        // maps emp ID to role
 *   }
 *
 * BENEFITS:
 *   - Single sheet, easy to back up and manage
 *   - Unlimited scalability (100+ departments, 100+ shifts per dept, etc.)
 *   - Clear key-value structure
 *   - No row/column conflicts
 *   - Time values in shifts are "HH:MM" strings (24-hour format)
 */


// ---------------------------------------------------------------------------
// Public API (called from api.js)
// ---------------------------------------------------------------------------

/**
 * Returns the settings for a single department from the consolidated Settings sheet.
 * Creates the Settings sheet with defaults if it doesn't exist yet.
 *
 * @param {string} deptName
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object, roles: Array, employeeRoles: object }}
 */
function getDeptSettings_(deptName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSettingsSheet_(workbook);
  return readDeptSettingsFromSheet_(sheet, deptName);
}

/**
 * Writes all settings for a department to the consolidated Settings sheet.
 * Supports partial updates — omitted fields are preserved from the existing settings.
 *
 * @param {string} deptName
 * @param {{
 *   staffingReqs?: Array<{ day: string, count: number, mode: string }>,
 *   shifts?: Array<{ name: string, ftpt: string, weekdayStart: string, ... }>,
 *   engineOptions?: { enforceRoleMinimums: boolean, gapFillEnabled: boolean },
 *   roleMinimums?: { [roleName: string]: { Low: number, Moderate: number, High: number } },
 *   roles?: Array<{ name: string, isPoolRole: boolean }>,
 *   employeeRoles?: { [employeeId: string]: string },
 * }} data (partial or complete)
 * @returns {{ saved: boolean }}
 */
function saveDeptSettings_(deptName, data) {
  // Ensure base structure exists and is valid
  let fullSettings = ensureDeptSettingsBaseStructure_(deptName);

  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSettingsSheet_(workbook);

  // Merge in the new data (partial updates supported)
  if (data.staffingReqs !== undefined && data.staffingReqs !== null) {
    fullSettings.staffingReqs = data.staffingReqs;
  }
  if (data.shifts !== undefined && data.shifts !== null) {
    fullSettings.shifts = data.shifts;
  }
  if (data.engineOptions !== undefined && data.engineOptions !== null) {
    fullSettings.engineOptions = data.engineOptions;
  }
  if (data.roleMinimums !== undefined && data.roleMinimums !== null) {
    fullSettings.roleMinimums = data.roleMinimums;
  }
  if (data.roles !== undefined && data.roles !== null) {
    fullSettings.roles = data.roles;
  }
  if (data.employeeRoles !== undefined && data.employeeRoles !== null) {
    fullSettings.employeeRoles = data.employeeRoles;
  }

  // Find or create the row for this department
  const rowNumber = findOrCreateDeptRow_(sheet, deptName);

  // Write merged settings as JSON to column B
  const jsonString = JSON.stringify(fullSettings);
  sheet.getRange(rowNumber, 2).setValue(jsonString); // Column B = settings JSON

  SpreadsheetApp.flush();
  return { saved: true };
}

/**
 * Ensures the base JSON structure exists for a department before any write operation.
 * Validates and repairs corrupted settings; initializes with defaults if missing.
 * Every function that writes settings should call this first.
 *
 * @param {string} deptName
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object, roles: Array, employeeRoles: object }}
 */
function ensureDeptSettingsBaseStructure_(deptName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSettingsSheet_(workbook);

  // Find or create the department row (initializes with defaults if missing)
  const rowNumber = findOrCreateDeptRow_(sheet, deptName);

  // Flush writes so the row is committed before we read it back
  SpreadsheetApp.flush();

  // Read back what we now have
  const data = sheet.getDataRange().getValues();
  const jsonCell = data[rowNumber - 1][1]; // rowNumber is 1-indexed, data array is 0-indexed

  let settings;
  try {
    // Try to parse existing JSON
    if (jsonCell && typeof jsonCell === 'string') {
      settings = JSON.parse(jsonCell);
    } else {
      settings = null;
    }
  } catch (e) {
    console.warn('ensureDeptSettingsBaseStructure_: JSON parse failed for ' + deptName + ', reinitializing with defaults');
    settings = null;
  }

  // If JSON is missing or corrupted, rebuild from defaults
  if (!settings || typeof settings !== 'object') {
    settings = getDefaultDeptSettings_();
    const jsonString = JSON.stringify(settings);
    sheet.getRange(rowNumber, 2).setValue(jsonString);
    console.log('ensureDeptSettingsBaseStructure_: reinitialized ' + deptName + ' with default structure');
  }

  // Validate and repair missing required fields
  let needsRepair = false;

  if (!Array.isArray(settings.staffingReqs)) {
    settings.staffingReqs = [];
    needsRepair = true;
  }
  if (!Array.isArray(settings.shifts)) {
    settings.shifts = [];
    needsRepair = true;
  }
  if (!settings.engineOptions || typeof settings.engineOptions !== 'object') {
    settings.engineOptions = { enforceRoleMinimums: true, gapFillEnabled: true };
    needsRepair = true;
  }
  if (!settings.roleMinimums || typeof settings.roleMinimums !== 'object') {
    settings.roleMinimums = {};
    needsRepair = true;
  }
  if (!Array.isArray(settings.roles)) {
    settings.roles = [];
    needsRepair = true;
  }
  if (!settings.employeeRoles || typeof settings.employeeRoles !== 'object') {
    settings.employeeRoles = {};
    needsRepair = true;
  }

  // Write repaired settings back if any field was missing
  if (needsRepair) {
    const jsonString = JSON.stringify(settings);
    sheet.getRange(rowNumber, 2).setValue(jsonString);
    SpreadsheetApp.flush();
    console.log('ensureDeptSettingsBaseStructure_: repaired missing fields for ' + deptName);
  }

  return settings;
}

/**
 * Returns the default department settings structure.
 * Used during initialization and repair.
 *
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object, roles: Array, employeeRoles: object }}
 */
function getDefaultDeptSettings_() {
  return {
    staffingReqs: DAY_NAMES_IN_ORDER.map(day => ({ day, count: DEFAULT_STAFFING_COUNT, mode: STAFFING_MODE.COUNT })),
    shifts: DEFAULT_SHIFTS, // config.js
    engineOptions: { enforceRoleMinimums: true, gapFillEnabled: true },
    roleMinimums: {},
    roles: [],
    employeeRoles: {}
  };
}


// ---------------------------------------------------------------------------
// Sheet Bootstrap
// ---------------------------------------------------------------------------

/**
 * Returns the consolidated Settings sheet, creating it with defaults if it doesn't exist.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSettingsSheet_(workbook) {
  const sheetName = SETTINGS_SHEET_NAME; // config.js
  let sheet = workbook.getSheetByName(sheetName);
  if (sheet) return sheet;

  sheet = workbook.insertSheet(sheetName);
  writeDefaultSettingsSheet_(sheet);
  return sheet;
}

/**
 * Finds the row number for a given department in the Settings sheet.
 * If the department doesn't exist, creates a new row for it with default settings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} deptName
 * @returns {number} rowNumber (1-indexed)
 */
function findOrCreateDeptRow_(sheet, deptName) {
  const data = sheet.getDataRange().getValues();

  // Search for existing department (starting from row 2, skip header row 1)
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim().toLowerCase() === deptName.toLowerCase()) {
      return i + 1; // Convert 0-indexed array to 1-indexed sheet row
    }
  }

  // Department not found, create a new row
  const newRow = data.length + 1;
  sheet.getRange(newRow, 1).setValue(deptName);

  // Initialize with default settings JSON
  const defaultSettings = {
    staffingReqs: DAY_NAMES_IN_ORDER.map(day => ({ day, count: DEFAULT_STAFFING_COUNT, mode: STAFFING_MODE.COUNT })),
    shifts: DEFAULT_SHIFTS, // config.js
    engineOptions: { enforceRoleMinimums: true, gapFillEnabled: true },
    roleMinimums: {},
    roles: [],
    employeeRoles: {}
  };
  sheet.getRange(newRow, 2).setValue(JSON.stringify(defaultSettings));

  return newRow;
}

/**
 * Writes headers and warning to a freshly created consolidated Settings sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function writeDefaultSettingsSheet_(sheet) {
  // --- Header row ---
  const headerRange = sheet.getRange(1, 1, 1, 2);
  headerRange.setValues([['Department', 'Settings JSON (Do not edit manually)']]);
  headerRange.setFontWeight('bold').setBackground(SHEET_TAB_COLORS.EMPLOYEES).setFontColor(COLORS.HEADER_TEXT); // config.js

  // Set column widths
  sheet.setColumnWidth(1, 150);  // Department name
  sheet.setColumnWidth(2, 2000); // Settings JSON (wide for readability)

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set tab color
  sheet.setTabColor(SHEET_TAB_COLORS.EMPLOYEES); // config.js
}


// ---------------------------------------------------------------------------
// Sheet Readers
// ---------------------------------------------------------------------------

/**
 * Reads settings for a specific department from the consolidated Settings sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} deptName
 * @returns {{ staffingReqs: Array, shifts: Array, engineOptions: object, roleMinimums: object, roles: Array, employeeRoles: object }}
 */
function readDeptSettingsFromSheet_(sheet, deptName) {
  try {
    const data = sheet.getDataRange().getValues();

    // Search for the department (starting from row 2, skip header row 1)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === deptName.toLowerCase()) {
        const jsonCell = data[i][1]; // Column B = settings JSON
        if (jsonCell && typeof jsonCell === 'string') {
          const parsed = JSON.parse(jsonCell);
          console.log('readDeptSettingsFromSheet_: loaded ' + deptName + ' from Settings sheet');
          return parsed;
        }
      }
    }

    console.warn('readDeptSettingsFromSheet_: department ' + deptName + ' not found, returning empty structure');
  } catch (e) {
    console.error('readDeptSettingsFromSheet_: JSON parse failed:', e.message);
  }

  // Return empty structure if department not found or JSON is corrupted
  return {
    staffingReqs:  [],
    shifts:        [],
    engineOptions: { enforceRoleMinimums: true, gapFillEnabled: true },
    roleMinimums:  {},
    roles:         [],
    employeeRoles: {}
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
