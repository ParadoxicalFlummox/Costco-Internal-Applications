/**
 * settingsManager.js — Reads shift definitions and staffing requirements from the Settings sheet.
 * VERSION: 1.2.0
 *
 * This file is the only place in the codebase that reads from the Settings sheet.
 * Every other file that needs shift or staffing data calls a function from this file
 * and receives a clean JavaScript object — no other file ever reads the Settings sheet directly.
 *
 * WHY THIS MATTERS:
 * Google Apps Script sheet reads are slow (each .getValues() call takes ~100–300ms).
 * By reading the Settings sheet once and returning a structured object, callers can
 * pass that object around in memory for the rest of the generation run without
 * triggering additional sheet reads.
 *
 * WHAT THE SETTINGS SHEET MUST CONTAIN:
 *   Table 1 (A2:B8)  — Staffing requirements: one row per day of the week.
 *   Table 2 (D2:I50) — Shift definitions: one row per shift variant (FT and PT versions
 *                       of each shift are separate rows with the same shift name).
 *
 * See README.md for the full Settings sheet setup guide.
 */


/**
 * Reads the Settings sheet and returns a map of shift definitions keyed by "ShiftName|Status".
 *
 * The compound key format (e.g., "Morning|FT", "Morning|PT") is used because the same
 * shift name can appear for both FT and PT employees with different start/end times and
 * paid hours. Using a compound key lets callers look up the exact variant they need with
 * a single map lookup instead of filtering an array.
 *
 * VALIDATION: This function logs a warning (but does not abort) if a shift row has
 * paid hours that violate the expected FT=8hr or PT=5hr rule. This makes misconfiguration
 * visible in the Execution Log without breaking a generation run over a minor discrepancy.
 *
 * @returns {Object} A map of "ShiftName|Status" → ShiftDefinition objects.
 *   Each ShiftDefinition has the shape:
 *   {
 *     name:         {string}  — e.g., "Morning"
 *     status:       {string}  — "FT" or "PT"
 *     startMinutes: {number}  — shift start expressed as minutes since midnight (e.g., 480 for 08:00)
 *     endMinutes:   {number}  — shift end expressed as minutes since midnight, including unpaid lunch block
 *     paidHours:    {number}  — hours that count toward the employee's weekly minimum/maximum
 *     blockHours:   {number}  — wall-clock hours the employee is physically present (may differ from paidHours)
 *     displayText:  {string}  — formatted time range shown in schedule cells, e.g., "08:00 - 16:30"
 *     hasLunch:     {boolean} — true if this shift includes an unpaid 30-minute lunch break
 *   }
 */
function buildShiftTimingMap(settingsSheet) {
  if (!settingsSheet) {
    settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  }

  if (!settingsSheet) {
    throw new Error(
      "Settings sheet not found. Run \"Setup Sheets\" from the Schedule Admin menu before generating."
    );
  }

  // Read all shift definition rows in a single call to minimize sheet read operations.
  const rawShiftRows = settingsSheet.getRange(SETTINGS_RANGE.SHIFT_DEFINITIONS_TABLE).getValues();

  if (!rawShiftRows || rawShiftRows.length === 0) {
    throw new Error(
      "No shift definitions found in the Settings sheet. " +
      "Add at least one shift row before generating a schedule."
    );
  }

  const shiftTimingMap = {};
  let validRowCount = 0;

  rawShiftRows.forEach(function (row, rowIndex) {
    const shiftName = row[SHIFT_TABLE_COLUMN.NAME];
    const status = row[SHIFT_TABLE_COLUMN.STATUS];
    const startValue = row[SHIFT_TABLE_COLUMN.START_TIME];
    const endValue = row[SHIFT_TABLE_COLUMN.END_TIME];
    const paidHours = row[SHIFT_TABLE_COLUMN.PAID_HOURS];
    const hasLunch = row[SHIFT_TABLE_COLUMN.HAS_LUNCH];

    // Skip blank rows — Google Sheets always returns the full range even if only
    // a few rows contain data, so empty rows at the bottom are expected.
    if (!shiftName || !status) {
      return;
    }

    // Convert GAS time values to minutes since midnight.
    // In Google Sheets, a time value stored as a fraction of a day (e.g., 0.333 for 08:00).
    // Multiplying by 1440 (minutes in a day) and rounding gives minutes since midnight.
    const startMinutes = convertGasTimeValueToMinutes(startValue, "start time", shiftName, rowIndex);
    const endMinutes = convertGasTimeValueToMinutes(endValue, "end time", shiftName, rowIndex);

    if (startMinutes === null || endMinutes === null) {
      // convertGasTimeValueToMinutes already logged the problem; skip this row.
      return;
    }

    if (endMinutes <= startMinutes) {
      // Midnight-crossing shifts are not supported because the coverage slot array
      // ends at 23:30. A shift ending at or before its start time is either a
      // data entry error or a midnight-crosser — either way, skip and warn.
      Logger.log(
        "WARNING: Shift \"" + shiftName + "\" (" + status + ") has an end time at or before " +
        "its start time. Midnight-crossing shifts are not supported. This row will be skipped."
      );
      return;
    }

    // The block hours represent how long the employee is physically present,
    // including any unpaid lunch. This is what drives the coverage slot map.
    // Paid hours (used for weekly hour enforcement) may be less if there is a lunch.
    const blockHours = (endMinutes - startMinutes) / 60;

    // Validate that the paid hours match the expected rules for this status.
    // This is a warning only — incorrect values in Settings are the manager's
    // responsibility to fix, and a warning in the log is more helpful than crashing.
    validateShiftPaidHours(shiftName, status, paidHours, blockHours);

    const displayText = formatMinutesAsTimeRange(startMinutes, endMinutes);

    // Use a compound key so both "Morning|FT" and "Morning|PT" can coexist in the map.
    const mapKey = shiftName + "|" + status;

    shiftTimingMap[mapKey] = {
      name: shiftName,
      status: status,
      startMinutes: startMinutes,
      endMinutes: endMinutes,
      paidHours: Number(paidHours),
      blockHours: blockHours,
      displayText: displayText,
      hasLunch: hasLunch === true,
    };

    validRowCount++;
  });

  if (validRowCount === 0) {
    throw new Error(
      "The Settings sheet has rows but none contained valid shift data. " +
      "Check that Shift Name and Status columns are filled in."
    );
  }

  return shiftTimingMap;
}


/**
 * Reads the Settings sheet staffing requirements table and returns a map of day name to minimum staff count.
 *
 * The staffing requirements table tells the engine how many employees must be
 * scheduled on each day of the week. The engine uses this number in Phase 1 to
 * determine how many RDO requests can be granted (it will not drop below this
 * floor), and in Phase 3 to detect coverage gaps.
 *
 * @returns {Object} A map of day name → minimum staff count, e.g.:
 *   { "Monday": 6, "Tuesday": 6, "Wednesday": 5, ..., "Sunday": 4 }
 */
function loadStaffingRequirements(settingsSheet) {
  if (!settingsSheet) {
    settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  }

  if (!settingsSheet) {
    throw new Error("Settings sheet not found. Run \"Setup Sheets\" from the Schedule Admin menu.");
  }

  const rawRequirementsRows = settingsSheet
    .getRange(SETTINGS_RANGE.STAFFING_REQUIREMENTS_TABLE)
    .getValues();

  const staffingRequirements = {};

  rawRequirementsRows.forEach(function (row) {
    const dayName     = row[0];
    const targetValue = row[1];
    const modeRaw     = (row[2] || "").toString().trim();

    // Skip blank rows in case the table range includes empty cells at the bottom.
    if (!dayName || targetValue === "" || targetValue === null) {
      return;
    }

    // Normalise mode: accept "Hours"/"hours" → HOURS, anything else → COUNT (default).
    const mode = modeRaw.toLowerCase() === "hours" ? STAFFING_MODE.HOURS : STAFFING_MODE.COUNT;

    staffingRequirements[dayName.toString().trim()] = {
      value: Number(targetValue),
      mode:  mode,
    };
  });

  // Warn if any of the seven days is missing from the staffing requirements.
  // A missing day means the engine will treat that day as requiring zero staff,
  // which could result in everyone getting that day off — clearly unintended.
  DAY_NAMES_IN_ORDER.forEach(function (dayName) {
    if (staffingRequirements[dayName] === undefined) {
      Logger.log(
        "WARNING: No staffing requirement found for \"" + dayName + "\" in the Settings sheet. " +
        "The engine will treat this day as requiring 0 staff. Add a row for this day."
      );
      staffingRequirements[dayName] = { value: 0, mode: STAFFING_MODE.COUNT };
    }
  });

  return staffingRequirements;
}


/**
 * Returns a deduplicated list of shift names from the Settings sheet.
 *
 * This list is used to populate the "Preferred Shift" and "Qualified Shifts"
 * dropdowns on the Roster sheet. Because the same shift name appears twice
 * (once for FT, once for PT), we deduplicate by collecting names into a Set.
 *
 * @returns {Array<string>} Unique shift names, in the order they first appear in Settings.
 */
function readShiftNamesFromSettings(settingsSheet) {
  if (!settingsSheet) {
    settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTINGS);
  }

  if (!settingsSheet) {
    return [];
  }

  const rawShiftRows = settingsSheet.getRange(SETTINGS_RANGE.SHIFT_DEFINITIONS_TABLE).getValues();

  // Use a Set to automatically discard duplicate shift names that arise from
  // having separate FT and PT rows with the same name.
  const uniqueShiftNames = new Set();

  rawShiftRows.forEach(function (row) {
    const shiftName = row[SHIFT_TABLE_COLUMN.NAME];
    if (shiftName && shiftName.toString().trim() !== "") {
      uniqueShiftNames.add(shiftName.toString().trim());
    }
  });

  return Array.from(uniqueShiftNames);
}


// ---------------------------------------------------------------------------
// Multi-Department Settings Loaders
// ---------------------------------------------------------------------------

/**
 * Reads the Departments tab and returns an array of active department entries.
 *
 * The Departments tab layout (one row per department, row 1 = headers):
 *   A — Department name (must match the Department column on the Roster sheet)
 *   B — Settings tab name (e.g., "Settings_Morning")
 *   C — Active flag (TRUE to include in generation, FALSE to skip)
 *   D — Header accent color (hex, e.g., "#4A90D9") — optional, falls back to COLORS.HEADER_BG
 *
 * @returns {Array<{ name, settingsTabName, active, accentColor }>}
 *   Empty array if the Departments tab does not exist (single-dept mode).
 */
function readDepartmentList_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const deptsSheet = workbook.getSheetByName(SHEET_NAMES.DEPARTMENTS);

  if (!deptsSheet) {
    return []; // Departments tab missing — caller falls back to single-dept mode
  }

  const lastRow = deptsSheet.getLastRow();
  if (lastRow < 2) return []; // header-only or empty

  const rows = deptsSheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const departments = [];

  rows.forEach(function (row) {
    const name          = (row[0] || '').toString().trim();
    const settingsTab   = (row[1] || '').toString().trim();
    const activeFlag    = row[2];
    const accentColor   = (row[3] || '').toString().trim() || COLORS.HEADER_BG;

    if (!name || !settingsTab) return; // blank row

    const isActive = (activeFlag === true || activeFlag === 'TRUE' || activeFlag === 'true');

    departments.push({
      name:            normalizeDeptName_(name), // canonical key used for all internal matching
      displayName:     name,                     // original text — used for headers and sheet names
      settingsTabName: settingsTab,
      active:          isActive,
      accentColor:     accentColor,
    });
  });

  return departments;
}


/**
 * Loads shift timing map and staffing requirements for a single department entry.
 *
 * @param {{ name, settingsTabName }} departmentEntry — One entry from readDepartmentList_()
 * @returns {{ shiftTimingMap, staffingRequirements }}
 * @throws {Error} if the Settings tab for this department cannot be found
 */
function loadSettingsForDepartment_(departmentEntry) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(departmentEntry.settingsTabName);

  if (!sheet) {
    throw new Error(
      "Settings tab \"" + departmentEntry.settingsTabName + "\" not found for department \"" +
      departmentEntry.name + "\". Create this tab or run \"Setup Department Settings\" from the menu."
    );
  }

  return {
    shiftTimingMap:        buildShiftTimingMap(sheet),
    staffingRequirements:  loadStaffingRequirements(sheet),
  };
}


/**
 * Loads settings for all active departments listed in the Departments tab.
 *
 * Returns a Map so callers can look up settings by department name in O(1).
 * If the Departments tab does not exist, returns null — the caller must fall
 * back to single-department mode using the existing Settings tab.
 *
 * @returns {Map<string, { shiftTimingMap, staffingRequirements, accentColor }> | null}
 */
function loadAllDepartmentSettings() {
  const departments = readDepartmentList_();

  if (departments.length === 0) {
    return null; // No Departments tab or empty — single-dept fallback
  }

  const settingsMap = new Map();

  departments.forEach(function (dept) {
    if (!dept.active) {
      console.log('settingsManager: Skipping inactive department "' + dept.name + '".');
      return;
    }

    try {
      const settings = loadSettingsForDepartment_(dept);
      settingsMap.set(dept.name, {
        shiftTimingMap:       settings.shiftTimingMap,
        staffingRequirements: settings.staffingRequirements,
        accentColor:          dept.accentColor,
        settingsTabName:      dept.settingsTabName,
        displayName:          dept.displayName, // original name for headers/sheet names
      });
      console.log('settingsManager: Loaded settings for "' + dept.name + '" (' + dept.settingsTabName + ').');
    } catch (error) {
      // Log and skip rather than aborting the entire multi-dept run over one bad Settings tab.
      console.error('settingsManager: Failed to load settings for "' + dept.name + '" — ' + error.message);
    }
  });

  if (settingsMap.size === 0) {
    throw new Error(
      "No valid department settings could be loaded. Check that at least one department " +
      "in the Departments tab is Active and has a valid Settings tab."
    );
  }

  return settingsMap;
}


// ---------------------------------------------------------------------------
// Private helper functions
// ---------------------------------------------------------------------------

/**
 * Converts a Google Apps Script time value to minutes since midnight.
 *
 * In Google Sheets, time values are stored as a decimal fraction of a 24-hour day.
 * For example, 08:00 is stored as 0.3333... (8/24), and 16:30 is stored as 0.6875 (16.5/24).
 * Multiplying by 1440 (the number of minutes in a day) converts this to minutes since midnight.
 *
 * WHY Utilities.formatDate() IS used for Date objects:
 * When getValues() returns a Date object for a time cell, calling getHours() or getMinutes()
 * on it returns values in the script's execution timezone — which is set in the Apps Script
 * project properties and may differ from the spreadsheet's timezone. Even a small mismatch
 * (e.g., a non-standard or incorrectly configured timezone) produces a consistent offset on
 * every shift time. Utilities.formatDate() with the spreadsheet's own timezone is the
 * GAS-idiomatic solution: it always returns the time exactly as it appears in the cell,
 * regardless of where the script is executing.
 *
 * @param {*}      timeValue  — The raw cell value from getValues(), either a number or Date.
 * @param {string} fieldLabel — "start time" or "end time", used in warning messages.
 * @param {string} shiftName  — The shift name, used in warning messages.
 * @param {number} rowIndex   — The 0-based row index within the shift table, used in warnings.
 * @returns {number|null} Minutes since midnight, or null if the value could not be parsed.
 */
function convertGasTimeValueToMinutes(timeValue, fieldLabel, shiftName, rowIndex) {
  if (timeValue instanceof Date) {
    // Use Utilities.formatDate() with the spreadsheets timezone to extract the
    // time as it appears in the cell which avoids the script timezone and sheet timezone
    // mismatch creating an offset when the two differ
    const spreadsheetTimeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const formattedTime = Utilities.formatDate(timeValue, spreadsheetTimeZone, "HH:mm");
    const timeParts = formattedTime.split(":");
    return parseInt(timeParts[0], 10) * 60 + parseInt(timeParts[1], 10);
  }

  if (typeof timeValue === "number" && timeValue >= 0 && timeValue < 1) {
    // GAS returns a raw decimal fraction (0 to <1) when the cell is formatted as
    // a plain number but contains a time. Multiplying by 1440 converts to minutes.
    return Math.round(timeValue * 1440);
  }

  Logger.log(
    "WARNING: Could not parse " + fieldLabel + " for shift \"" + shiftName + "\" " +
    "(Settings row " + (rowIndex + 2) + "). Expected a time value but got: " +
    JSON.stringify(timeValue) + ". This shift row will be skipped."
  );
  return null;
}


/**
 * Logs a warning if a shift's paid hours fall outside the expected range for its status and type.
 *
 * Rules:
 *   FT shifts:    exactly 8.0 paid hours.
 *   PT shifts:    exactly 5.0 paid hours.
 *   PT+ shifts:   between PT_PLUS_MIN_HOURS (5) and PT_PLUS_MAX_HOURS (8) paid hours.
 *                 The "+" suffix marks lunch-qualified shifts that can be extended for coverage.
 *
 * A deviation from the expected range usually means a data entry error in the Settings sheet.
 * This function makes it visible in the Execution Log without aborting generation.
 *
 * @param {string}  shiftName  — The name of the shift, for the warning message.
 * @param {string}  status     — "FT" or "PT".
 * @param {number}  paidHours  — The paid hours value read from the Settings row.
 * @param {number}  blockHours — The computed wall-clock block hours (end - start in hours).
 */
function validateShiftPaidHours(shiftName, status, paidHours, blockHours) {
  const numPaidHours = Number(paidHours);
  const isPlusShift = shiftName.toString().endsWith("+");

  if (status === "PT" && isPlusShift) {
    // PT+ shifts are lunch-qualified and can be scheduled between PT_PLUS_MIN_HOURS and
    // PT_PLUS_MAX_HOURS paid hours per shift. The block hours (including the 30-min lunch)
    // should equal paidHours + 0.5.
    const expectedBlockHours = numPaidHours + 0.5;
    if (
      numPaidHours < HOUR_RULES.PT_PLUS_MIN_HOURS ||
      numPaidHours > HOUR_RULES.PT_PLUS_MAX_HOURS ||
      Math.abs(blockHours - expectedBlockHours) > 0.01
    ) {
      Logger.log(
        "WARNING: Shift \"" + shiftName + "\" (" + status + ") has " + paidHours +
        " paid hours and a " + blockHours.toFixed(2) + "-hour clock block. PT+ shifts should " +
        "have " + HOUR_RULES.PT_PLUS_MIN_HOURS + "–" + HOUR_RULES.PT_PLUS_MAX_HOURS +
        " paid hours with the end time set to paidHours + 0:30 (unpaid lunch). " +
        "If this is intentional, ignore this warning. Otherwise, correct the Settings sheet."
      );
    }
  } else {
    const expectedPaidHours = status === "FT" ? 8.0 : 5.0;

    if (numPaidHours !== expectedPaidHours) {
      Logger.log(
        "WARNING: Shift \"" + shiftName + "\" (" + status + ") has " + paidHours +
        " paid hours. Expected " + expectedPaidHours + " for a " + status + " shift. " +
        "If this is intentional, ignore this warning. Otherwise, correct the Paid Hours " +
        "column in the Settings sheet."
      );
    }
  }
}


/**
 * Formats two minute-since-midnight values as a human-readable time range string.
 *
 * The output format "HH:mm - HH:mm" is written into SHIFT row cells of the generated
 * schedule. The COUNTIF formula in the ACTUAL summary row uses "*:*" as its pattern,
 * which matches any string containing a colon — so this format is also what makes
 * the summary formulas work correctly.
 *
 * @param {number} startMinutes — Start time in minutes since midnight.
 * @param {number} endMinutes   — End time in minutes since midnight.
 * @returns {string} A formatted time range, e.g., "8:00 AM - 4:30 PM".
 */
function formatMinutesAsTimeRange(startMinutes, endMinutes) {
  return formatMinutesAsTimeString(startMinutes) + " - " + formatMinutesAsTimeString(endMinutes);
}


/**
 * Converts a minutes-since-midnight value to a 12-hour "h:mm AM/PM" string.
 *
 * 12-hour format is used because it is what managers and employees naturally read
 * at a glance — "8:00 AM - 4:30 PM" is immediately understood without mental
 * conversion from 24-hour notation.
 *
 * Edge cases:
 *   - Midnight (0 min)  → "12:00 AM"
 *   - Noon (720 min)    → "12:00 PM"
 *   - 13:00 (780 min)   → "1:00 PM"
 *
 * @param {number} totalMinutes — Minutes since midnight (e.g., 510 for 08:30).
 * @returns {string} A formatted time string, e.g., "8:30 AM".
 */
function formatMinutesAsTimeString(totalMinutes) {
  const totalHours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  const period = totalHours >= 12 ? "PM" : "AM";

  // Convert from 24-hour to 12-hour. Hour 0 and hour 12 both display as 12.
  const twelveHourValue = totalHours % 12 === 0 ? 12 : totalHours % 12;

  // Only the minutes are zero-padded (e.g., 8:05, not 8:5).
  // The hour is left unpadded so "8:00 AM" reads more naturally than "08:00 AM".
  return twelveHourValue + ":" + String(minutes).padStart(2, "0") + " " + period;
}
