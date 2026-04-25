/**
 * settingsManager.js — Builds shift timing maps and staffing requirements for the schedule engine.
 * VERSION: 0.5.1
 *
 * In the original autoScheduleGenerator this file read directly from a Settings sheet.
 * In COMET, settings are owned by scheduleSettings.js and stored in Settings_[Dept] sheets.
 * This file now reads from getDeptSettings_() and converts the plain data objects into
 * the same in-memory structures (shiftTimingMap, staffingRequirements) that the engine expects.
 *
 * Shift timing schema (v0.5.1):
 *   startTime/endTime replaced by weekdayStart + satStart + sunStart + paidHours.
 *   End time is computed: anchorStart + paidHours×60 + (hasLunch ? 30 : 0).
 *   Use getStartMinutesForDay_(shiftDef, dayName) to get the day-specific anchor start.
 */


/**
 * Builds the shift timing map from a settings data object.
 *
 * Returns a plain object keyed by "ShiftName|Status" (e.g., "Morning|FT").
 * Each value is a ShiftDefinition:
 *   { name, status,
 *     weekdayStartMinutes, satStartMinutes, sunStartMinutes,
 *     startMinutes (alias for weekdayStartMinutes),
 *     blockMinutes, endMinutes (alias, weekday only),
 *     paidHours, blockHours, displayText,
 *     hasLunch, flexEnabled, flexWindowEarliest, flexWindowLatest }
 *
 * @param {string} deptName — Department name (used to load settings from the sheet).
 * @returns {Object} shiftTimingMap
 */
function buildShiftTimingMap(deptName) {
  const settingsData = getDeptSettings_(deptName); // scheduleSettings.js
  const shiftTimingMap = {};

  (settingsData.shifts || []).forEach(function(shift) {
    const shiftName = (shift.name || '').trim();
    const status    = (shift.ftpt || '').trim();

    if (!shiftName || !status) return;

    const weekdayStartMinutes = timeStringToMinutes_(shift.weekdayStart); // scheduleSettings.js
    // Sat/Sun anchors fall back to weekday anchor when blank.
    const satStartMinutes = shift.satStart ? timeStringToMinutes_(shift.satStart) : weekdayStartMinutes;
    const sunStartMinutes = shift.sunStart ? timeStringToMinutes_(shift.sunStart) : weekdayStartMinutes;

    const paidHours = Number(shift.paidHours || 0);
    // Block duration includes the unpaid 30-min lunch break in the clock-time span.
    const blockMinutes = paidHours * 60 + (shift.hasLunch === true ? 30 : 0);

    if (!weekdayStartMinutes && shift.flexEnabled === false) {
      console.warn('settingsManager: Fixed shift "' + shiftName + '" has no weekday anchor — skipped.');
      return;
    }

    const displayText = formatMinutesAsTimeRange(weekdayStartMinutes, weekdayStartMinutes + blockMinutes);
    const mapKey = shiftName + '|' + status;

    shiftTimingMap[mapKey] = {
      name:                 shiftName,
      status:               status,
      weekdayStartMinutes:  weekdayStartMinutes,
      satStartMinutes:      satStartMinutes,
      sunStartMinutes:      sunStartMinutes,
      startMinutes:         weekdayStartMinutes,              // backward-compat alias
      blockMinutes:         blockMinutes,
      endMinutes:           weekdayStartMinutes + blockMinutes, // backward-compat alias (weekday only)
      paidHours:            paidHours,
      blockHours:           blockMinutes / 60,
      displayText:          displayText,
      hasLunch:             shift.hasLunch === true,
      flexEnabled:          shift.flexEnabled !== false,
      flexWindowEarliest:   shift.flexWindowEarliest || '',
      flexWindowLatest:     shift.flexWindowLatest   || '',
    };
  });

  return shiftTimingMap;
}


/**
 * Returns the anchor start time in minutes-since-midnight for a specific day.
 * Uses the Saturday or Sunday override when configured; falls back to weekday anchor.
 *
 * @param {Object} shiftDef — Entry from the shift timing map.
 * @param {string} dayName  — e.g. "Saturday", "Sunday", "Monday"
 * @returns {number} minutes since midnight
 */
function getStartMinutesForDay_(shiftDef, dayName) {
  if (dayName === 'Saturday') return shiftDef.satStartMinutes || shiftDef.weekdayStartMinutes || 0;
  if (dayName === 'Sunday')   return shiftDef.sunStartMinutes || shiftDef.weekdayStartMinutes || 0;
  return shiftDef.weekdayStartMinutes || 0;
}


/**
 * Builds the staffing requirements map from a settings data object.
 *
 * Returns a plain object keyed by day name:
 *   { "Monday": { value: 6, mode: "Count" }, ... }
 *
 * @param {string} deptName — Department name.
 * @returns {Object} staffingRequirements
 */
function loadStaffingRequirements(deptName) {
  const settingsData = getDeptSettings_(deptName); // scheduleSettings.js
  const staffingRequirements = {};

  (settingsData.staffingReqs || []).forEach(function(req) {
    const day = (req.day || '').trim();
    if (!day) return;
    staffingRequirements[day] = {
      value: Number(req.count || 0),
      mode:  (req.mode || STAFFING_MODE.COUNT).trim(), // config.js
    };
  });

  // Fill in any missing days with zero to avoid undefined access in the engine.
  DAY_NAMES_IN_ORDER.forEach(function(dayName) { // config.js
    if (!staffingRequirements[dayName]) {
      console.warn('settingsManager: No staffing requirement for "' + dayName + '" — defaulting to 0.');
      staffingRequirements[dayName] = { value: 0, mode: STAFFING_MODE.COUNT };
    }
  });

  return staffingRequirements;
}


/**
 * Returns the unique shift names available for a department.
 * Used to populate dropdown options in the employee edit modal.
 *
 * @param {string} deptName
 * @returns {string[]}
 */
function readShiftNamesForDept_(deptName) {
  const settingsData = getDeptSettings_(deptName);
  const seen = new Set();
  (settingsData.shifts || []).forEach(s => {
    const name = (s.name || '').trim();
    if (name) seen.add(name);
  });
  return Array.from(seen);
}


/**
 * Reads supervisor peak traffic configuration for a department from the COMET Config sheet.
 * Returns null if no config exists (caller should use defaults from config.js).
 *
 * @param {string} deptName
 * @returns {Object|null} { department, peakProfile, minCountPerPeak, minCountPerValley, peakThreshold, lastUpdated }
 */
/**
 * Reads traffic heatmap configuration for a department from the COMET Config sheet.
 * Returns null if no config exists (caller should use defaults from config.js).
 *
 * @param {string} deptName
 * @returns {Object|null} { department, enabled, thresholds, trafficCurves, pool, weeklyOverrides, lastUpdated }
 */
function loadTrafficHeatmapConfig_(deptName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME); // config.js
  if (!configSheet) return null;

  const configKey = 'TRAFFIC_HEATMAP_' + (deptName || '').toUpperCase().replace(/\s+/g, '_');
  const configData = configSheet.getDataRange().getValues();

  for (let rowIdx = 0; rowIdx < configData.length; rowIdx++) {
    const row = configData[rowIdx];
    if (row[0] === configKey && row[1]) {
      try {
        return JSON.parse(row[1].toString());
      } catch (e) {
        logConsole_('WARNING: Failed to parse traffic heatmap config for ' + deptName);
        return null;
      }
    }
  }

  return null;
}

/**
 * Writes traffic heatmap configuration for a department to the COMET Config sheet.
 *
 * @param {string} deptName
 * @param {Object} heatmapConfig — { enabled, thresholds, trafficCurves, pool, weeklyOverrides, ... }
 */
function saveTrafficHeatmapConfig_(deptName, heatmapConfig) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME); // config.js
  if (!configSheet) {
    configSheet = workbook.insertSheet(COMET_CONFIG_SHEET_NAME);
  }

  const configKey = 'TRAFFIC_HEATMAP_' + (deptName || '').toUpperCase().replace(/\s+/g, '_');
  const configValue = JSON.stringify({
    department: deptName,
    enabled: heatmapConfig.enabled || false,
    thresholds: heatmapConfig.thresholds || HEATMAP_DEFAULTS.thresholds, // config.js
    trafficCurves: heatmapConfig.trafficCurves || HEATMAP_DEFAULTS.defaultTrafficCurves,
    pool: heatmapConfig.pool || { tier1EmployeeIds: [], tier2EmployeeIds: [], requiredRoles: [], schedulingCounts: HEATMAP_DEFAULTS.poolSchedulingCounts },
    weeklyOverrides: heatmapConfig.weeklyOverrides || {},
    lastUpdated: new Date().toISOString(),
  });

  const configData = configSheet.getDataRange().getValues();
  let foundRow = -1;

  for (let rowIdx = 0; rowIdx < configData.length; rowIdx++) {
    if (configData[rowIdx][0] === configKey) {
      foundRow = rowIdx + 1; // 1-indexed
      break;
    }
  }

  if (foundRow > 0) {
    // Update existing row
    configSheet.getRange(foundRow, 2).setValue(configValue);
  } else {
    // Append new row
    const nextRow = configData.length + 1;
    configSheet.getRange(nextRow, 1).setValue(configKey);
    configSheet.getRange(nextRow, 2).setValue(configValue);
  }
}


// ---------------------------------------------------------------------------
// Helpers (ported from autoScheduleGenerator settingsManager.js)
// ---------------------------------------------------------------------------

/**
 * Formats two minute-since-midnight values as a human-readable time range string.
 * e.g. 480, 990 → "8:00 AM - 4:30 PM"
 *
 * @param {number} startMinutes
 * @param {number} endMinutes
 * @returns {string}
 */
function formatMinutesAsTimeRange(startMinutes, endMinutes) {
  return formatMinutesAsTimeString(startMinutes) + ' - ' + formatMinutesAsTimeString(endMinutes);
}

/**
 * Converts minutes-since-midnight to a 12-hour "h:mm AM/PM" string.
 *
 * @param {number} totalMinutes
 * @returns {string}
 */
function formatMinutesAsTimeString(totalMinutes) {
  const totalHours = Math.floor(totalMinutes / 60);
  const minutes    = totalMinutes % 60;
  const period     = totalHours >= 12 ? 'PM' : 'AM';
  const twelve     = totalHours % 12 === 0 ? 12 : totalHours % 12;
  return twelve + ':' + String(minutes).padStart(2, '0') + ' ' + period;
}
