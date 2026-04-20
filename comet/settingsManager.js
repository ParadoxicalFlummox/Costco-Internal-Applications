/**
 * settingsManager.js — Builds shift timing maps and staffing requirements for the schedule engine.
 * VERSION: 0.2.3
 *
 * In the original autoScheduleGenerator this file read directly from a Settings sheet.
 * In COMET, settings are owned by scheduleSettings.js and stored in Settings_[Dept] sheets.
 * This file now reads from getDeptSettings_() and converts the plain data objects into
 * the same in-memory structures (shiftTimingMap, staffingRequirements) that the engine expects.
 *
 * The engine itself (scheduleEngine.js) is unchanged — it still calls
 * buildShiftTimingMap(settingsData) and loadStaffingRequirements(settingsData),
 * but now these functions receive a pre-loaded data object rather than a sheet reference.
 */


/**
 * Builds the shift timing map from a settings data object.
 *
 * Returns a plain object keyed by "ShiftName|Status" (e.g., "Morning|FT").
 * Each value is a ShiftDefinition:
 *   { name, status, startMinutes, endMinutes, paidHours, blockHours, displayText, hasLunch }
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

    const startMinutes = timeStringToMinutes_(shift.startTime); // scheduleSettings.js
    const endMinutes   = timeStringToMinutes_(shift.endTime);

    if (endMinutes <= startMinutes) {
      console.warn('settingsManager: Shift "' + shiftName + '" has end <= start — skipped.');
      return;
    }

    const blockHours = (endMinutes - startMinutes) / 60;
    const displayText = formatMinutesAsTimeRange(startMinutes, endMinutes);
    const mapKey = shiftName + '|' + status;

    shiftTimingMap[mapKey] = {
      name:         shiftName,
      status:       status,
      startMinutes: startMinutes,
      endMinutes:   endMinutes,
      paidHours:    Number(shift.paidHours || 0),
      blockHours:   blockHours,
      displayText:  displayText,
      hasLunch:     shift.hasLunch === true,
    };
  });

  return shiftTimingMap;
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
