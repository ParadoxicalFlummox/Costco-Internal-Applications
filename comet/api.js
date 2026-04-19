/**
 * api.js — Public API layer for COMET.
 * VERSION: 0.1.1
 * 
 * This file contains the functions that the frontend calls via google.script.run.
 * Every public function here is a thin wrapper: it validates inputs, calls the
 * appropriate backend module, and returns a plain serializable object.
 *
 * WHY THIS LAYER EXISTS:
 *   GAS only allows the frontend to call globally scoped functions. This file
 *   collects all of those entry points in one place so the rest of the backend
 *   can use private (_-suffixed) functions freely without accidentally exposing
 *   them to the web.
 *
 * RETURN CONTRACT:
 *   Every function returns a plain object: { ok: true, data: ... } on success
 *   or { ok: false, error: 'message' } on failure. The frontend checks ok before
 *   using data, so errors are always surfaced to the user rather than silently
 *   dropped.
 *
 * PHASE NOTE:
 *   Phase 1 — All functions return mock data so the frontend can be built and
 *   tested before the backend is wired up. Each stub is marked with a TODO
 *   indicating which phase and module will replace it.
 *
 * FUNCTION INVENTORY:
 *
 *   Schedule:
 *     getScheduleForWeek(deptName, mondayDate)                              → { weekSheetName, weekGrid, employeeList }
 *     generateSchedule(deptName, mondayDate)                                → { weekSheetName, weekGrid, employeeList }
 *     getDeptSettings(deptName)                                             → { shifts, staffingReqs }
 *     saveDeptSettings(deptName, data)                                      → { saved }
 *     updateCellOverride(weekSheetName, employeeId, dayIndex, newType)      → { weekGrid, employeeList }
 *     updateEmployeeScheduleFields(id, fields)                              → { updated }
 *
 *   Absences:
 *     getAbsenceLogForDay(dateString)          → { entries }
 *     logAbsenceEntry(data)                    → { rowNumber }
 *     sendNotificationForRow(sheetName, row)   → { sent }
 *
 *   Infractions:
 *     getActiveCNs()                           → { cns }
 *     runCNScan(dryRun)                        → { proposals, issued }
 *     runExpiryCheck()                         → { expired }
 *
 *   Admin:
 *     importFromUKG(rows)                      → { added, updated, skipped }
 *     getEmployeeList()                        → { employees }
 *     setEmployeeStatus(id, status)            → { updated }
 *     generateNextYearWorkbook()               → { url }
 *     getConfig()                              → { config }
 *     updateConfig(key, value)                 → { updated }
 */


// ---------------------------------------------------------------------------
// Schedule
// ---------------------------------------------------------------------------

/**
 * Returns the last-generated schedule for the given department and week.
 * Returns null weekGrid if no Week sheet exists for this dept + week.
 *
 * @param {string} deptName   — Department name.
 * @param {string} mondayDate — ISO date string "YYYY-MM-DD" for the Monday of the week.
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function getScheduleForWeek(deptName, mondayDate) {
  try {
    if (!deptName || !mondayDate) throw new Error('deptName and mondayDate are required.');
    const weekStartDate = new Date(mondayDate + 'T00:00:00');
    const result = readExistingWeekSchedule_(deptName, weekStartDate); // scheduleEngine.js
    if (!result) {
      return { ok: true, data: { weekSheetName: null, weekGrid: null, employeeList: [] } };
    }
    return {
      ok: true,
      data: {
        weekSheetName: result.weekSheetName,
        weekGrid: serializeWeekGrid_(result.weekGrid, result.employeeList),
        employeeList: serializeEmployeeList_(result.employeeList),
      },
    };
  } catch (error) {
    console.error('api: getScheduleForWeek failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Generates a new schedule for the given department and week, writes it to
 * the workbook, and returns the grid for display in the browser.
 *
 * @param {string} deptName
 * @param {string} mondayDate — ISO date string "YYYY-MM-DD".
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function generateSchedule(deptName, mondayDate) {
  try {
    if (!deptName || !mondayDate) throw new Error('deptName and mondayDate are required.');
    const weekStartDate = new Date(mondayDate + 'T00:00:00');
    const result = generateWeeklySchedule_(deptName, weekStartDate); // scheduleEngine.js

    // Write to sheet via formatter.
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = workbook.getSheetByName(result.weekSheetName);
    if (!sheet) sheet = workbook.insertSheet(result.weekSheetName);
    const staffingRequirements = loadStaffingRequirements(deptName); // settingsManager.js
    writeAndFormatSchedule(sheet, result.employeeList, result.weekGrid, staffingRequirements, weekStartDate, deptName); // formatter.js

    return {
      ok: true,
      data: {
        weekSheetName: result.weekSheetName,
        weekGrid: serializeWeekGrid_(result.weekGrid, result.employeeList),
        employeeList: serializeEmployeeList_(result.employeeList),
      },
    };
  } catch (error) {
    console.error('api: generateSchedule failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the shift definitions and staffing requirements for a department.
 *
 * @param {string} deptName
 * @returns {{ ok: boolean, data?: { shifts, staffingReqs }, error?: string }}
 */
function getDeptSettings(deptName) {
  try {
    if (!deptName) throw new Error('deptName is required.');
    const settings = getDeptSettings_(deptName); // scheduleSettings.js
    return { ok: true, data: settings };
  } catch (error) {
    console.error('api: getDeptSettings failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Saves shift definitions and staffing requirements for a department.
 *
 * @param {string} deptName
 * @param {{ shifts: Array, staffingReqs: Array }} data
 * @returns {{ ok: boolean, data?: { saved: boolean }, error?: string }}
 */
function saveDeptSettings(deptName, data) {
  try {
    if (!deptName) throw new Error('deptName is required.');
    saveDeptSettings_(deptName, data); // scheduleSettings.js
    return { ok: true, data: { saved: true } };
  } catch (error) {
    console.error('api: saveDeptSettings failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Overrides a single cell in an existing Week sheet (VAC, RDO, or SHIFT),
 * then re-runs just the SHIFT assignment phase and returns the updated grid.
 *
 * @param {string} weekSheetName — Name of the existing Week sheet tab.
 * @param {string} employeeId    — Employee ID whose row to update.
 * @param {number} dayIndex      — 0 = Monday … 6 = Sunday.
 * @param {string} newType       — 'VAC', 'RDO', or 'SHIFT'.
 * @returns {{ ok: boolean, data?: { weekGrid, employeeList }, error?: string }}
 */
function updateCellOverride(weekSheetName, employeeId, dayIndex, newType) {
  try {
    if (!weekSheetName || !employeeId || dayIndex == null || !newType) {
      throw new Error('weekSheetName, employeeId, dayIndex, and newType are all required.');
    }
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = workbook.getSheetByName(weekSheetName);
    if (!sheet) throw new Error('Week sheet "' + weekSheetName + '" not found.');

    // Derive dept name and weekStartDate from the sheet name: "Week_MM_DD_YY_DeptName"
    // Format: Week_MM_DD_YY where MM/DD/YY are two-digit month, day, 2-digit year.
    const parts = weekSheetName.split('_');
    const deptName = parts.slice(4).join('_'); // everything after Week_MM_DD_YY_
    if (!deptName) throw new Error('Could not derive department from sheet name "' + weekSheetName + '".');

    // Reconstruct weekStartDate from parts[1]=MM, parts[2]=DD, parts[3]=YY
    const mm = parts[1];
    const dd = parts[2];
    const yy = parts[3];
    const weekStartDate = new Date('20' + yy + '-' + mm + '-' + dd + 'T00:00:00');

    const employeeList = loadRosterSortedBySeniority_(deptName); // scheduleEngine.js
    const employeeIndex = employeeList.findIndex(e => String(e.employeeId) === String(employeeId));
    if (employeeIndex === -1) throw new Error('Employee ' + employeeId + ' not found in ' + deptName + ' roster.');

    // Write the override into the checkbox row of the existing sheet so re-reads are consistent.
    const checkboxCol = WEEK_SHEET.COL_MONDAY + dayIndex;
    const baseRow = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);

    // Clear the VAC and RDO checkbox rows for this employee+day, then mark the override.
    sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, checkboxCol).setValue(false);
    sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, checkboxCol).setValue(false);
    sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT, checkboxCol).setValue('');

    if (newType === 'VAC') {
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, checkboxCol).setValue(true);
    } else if (newType === 'RDO') {
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, checkboxCol).setValue(true);
    }
    // SHIFT type: leave blanks — re-generation will fill it in.

    // Rebuild the full grid from the updated sheet state, then re-run SHIFT phase.
    const weekGrid = readCheckboxStateFromSheet_(sheet, employeeList.length); // scheduleEngine.js
    const shiftTimingMap = buildShiftTimingMap(deptName);                           // settingsManager.js
    const staffingReqs = loadStaffingRequirements(deptName);                      // settingsManager.js

    runPhaseOnePreferenceAssignment_(weekGrid, employeeList, shiftTimingMap, staffingReqs, null);
    runPhaseTwoHourEnforcement_(weekGrid, employeeList, shiftTimingMap);
    runPhaseThreeGapResolution_(weekGrid, employeeList, shiftTimingMap, staffingReqs);
    runPhaseFourRoleAssignment_(weekGrid, employeeList, deptName);

    // Flush updated grid back to the sheet so it stays in sync.
    writeAndFormatSchedule(sheet, employeeList, weekGrid, staffingReqs, weekStartDate, deptName); // formatter.js

    return {
      ok: true,
      data: {
        weekGrid: serializeWeekGrid_(weekGrid, employeeList),
        employeeList: serializeEmployeeList_(employeeList),
      },
    };
  } catch (error) {
    console.error('api: updateCellOverride failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Saves schedule-specific fields (F–M) for a single employee.
 *
 * @param {string} id     — Employee ID.
 * @param {Object} fields — Object with any subset of: ftpt, dayOffPrefOne, dayOffPrefTwo,
 *                          preferredShift, qualifiedShifts, vacationDates, role.
 * @returns {{ ok: boolean, data?: { updated: boolean }, error?: string }}
 */
function updateEmployeeScheduleFields(id, fields) {
  try {
    if (!id) throw new Error('id is required.');
    const updated = updateEmployeeScheduleFields_(id, fields); // ukgImport.js
    return { ok: true, data: { updated } };
  } catch (error) {
    console.error('api: updateEmployeeScheduleFields failed —', error);
    return { ok: false, error: error.message };
  }
}


// ---------------------------------------------------------------------------
// Absences
// ---------------------------------------------------------------------------

/**
 * Returns all absence log entries for the given date.
 *
 * @param {string} dateString — ISO date string (YYYY-MM-DD).
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function getAbsenceLogForDay(dateString) {
  try {
    // Strip dateRaw (a Date object) — google.script.run only serializes primitives.
    const entries = getCallLogEntriesForDate_(dateString).map(e => ({ // callLog.js
      name: e.name,
      employeeId: e.employeeId,
      isCallout: e.isCallout,
      isFmla: e.isFmla,
      isNoShow: e.isNoShow,
      department: e.department,
      time: e.time,
      comment: e.comment,
      sheetName: e.sheetName,
      rowNumber: e.rowNumber,
    }));
    return { ok: true, data: { date: dateString, entries } };
  } catch (error) {
    console.error('api: getAbsenceLogForDay failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Writes a new absence entry to the active Call Log sheet.
 *
 * @param {{ name, employeeId, department, time, isCallout, isFmla, isNoShow, comment }} data
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function logAbsenceEntry(data) {
  try {
    const result = appendCallLogEntry_(data); // callLog.js
    return { ok: true, data: result };
  } catch (error) {
    console.error('api: logAbsenceEntry failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Sends the absence notification email for a specific row in the Call Log.
 *
 * @param {string} sheetName — The Call Log sheet tab name.
 * @param {number} rowNumber — 1-indexed row number to send.
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function sendNotificationForRow(sheetName, rowNumber) {
  try {
    const result = sendAbsenceNotification_(sheetName, rowNumber); // callLog.js
    return { ok: true, data: result };
  } catch (error) {
    console.error('api: sendNotificationForRow failed —', error);
    return { ok: false, error: error.message };
  }
}


// ---------------------------------------------------------------------------
// Infractions
// ---------------------------------------------------------------------------

/**
 * Returns all currently Active CNs from the Active CNs sheet.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function getActiveCNs() {
  try {
    const cns = getActiveCNsForDashboard_(); // cnLog.js
    return { ok: true, data: { cns } };
  } catch (error) {
    console.error('api: getActiveCNs failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Runs the infraction scanner across all active employee tabs.
 *
 * @param {boolean} dryRun — When true, logs proposals but sends no emails
 *   and writes nothing to CN_Log.
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function runCNScan(dryRun) {
  try {
    scanAndIssueCNs({ dryRun: !!dryRun }); // infractionEngine.js
    return { ok: true, data: { issued: 0, dryRun: !!dryRun } };
  } catch (error) {
    console.error('api: runCNScan failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Runs the CN expiry check, marking Active CNs older than EXPIRY_DAYS as
 * Expired and sending expiry notifications to payroll.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function runExpiryCheck() {
  try {
    expireCNsDaily(DRY_RUN); // cnLog.js — respects config.js DRY_RUN flag
    return { ok: true, data: { expired: 0 } };
  } catch (error) {
    console.error('api: runExpiryCheck failed —', error);
    return { ok: false, error: error.message };
  }
}


// ---------------------------------------------------------------------------
// Admin
// ---------------------------------------------------------------------------

/**
 * Runs first-time setup, creating all required sheets if they don't exist.
 * Safe to call multiple times — existing sheets are left untouched.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function runSetup() {
  try {
    const result = runCometSetup_(); // setup.js
    return { ok: true, data: result };
  } catch (error) {
    console.error('api: runSetup failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Upserts employee rows parsed from a UKG CSV into the Employees sheet.
 *
 * The frontend parses the CSV client-side, filters placeholder rows, and
 * sends a clean array of employee objects here. This keeps CSV parsing in
 * the browser and sheet writes on the server.
 *
 * @param {Array<{ name: string, id: string, hireDate: string, department: string }>} rows
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function importFromUKG(rows) {
  try {
    const result = importEmployeesFromUkg_(rows); // ukgImport.js
    return { ok: true, data: result };
  } catch (error) {
    console.error('api: importFromUKG failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the sorted unique department names from all Active employees.
 * Used to populate department dropdowns in the Schedule and Absences views.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function getDepartments() {
  try {
    const employees = getActiveEmployees_(); // ukgImport.js
    const departments = [...new Set(
      employees.map(e => e.department).filter(Boolean)
    )].sort();
    return { ok: true, data: { departments } };
  } catch (error) {
    console.error('api: getDepartments failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the full employee list from the Employees sheet, including
 * both Active and Archived employees.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function getEmployeeList() {
  try {
    const employees = getAllEmployees_(); // ukgImport.js
    return { ok: true, data: { employees } };
  } catch (error) {
    console.error('api: getEmployeeList failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Sets the Status column of an employee row to the given value.
 *
 * @param {string} id — Employee ID (UKG employee number).
 * @param {'Active'|'Archived'} status
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 */
function setEmployeeStatus(id, status) {
  try {
    const updated = setEmployeeStatus_(id, status); // ukgImport.js
    return { ok: true, data: { updated } };
  } catch (error) {
    console.error('api: setEmployeeStatus failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Generates the next fiscal year's attendance controller workbook and
 * returns a link to the new file in Google Drive.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 * TODO Phase 6: Replace stub with tabManager.js next-year generation call.
 */
function generateNextYearWorkbook() {
  try {
    // Phase 1 stub
    return { ok: true, data: { url: null } };
  } catch (error) {
    console.error('api: generateNextYearWorkbook failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the current values from the COMET Config sheet.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 * TODO Phase 6: Replace stub with COMET Config sheet read.
 */
function getConfig() {
  try {
    // Phase 1 stub
    return { ok: true, data: { config: {} } };
  } catch (error) {
    console.error('api: getConfig failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Writes a single key-value pair to the COMET Config sheet.
 *
 * @param {string} key — Config key (must match a row in the COMET Config sheet).
 * @param {string} value — New value to write.
 * @returns {{ ok: boolean, data?: object, error?: string }}
 *
 * TODO Phase 6: Replace stub with COMET Config sheet write.
 */
function updateConfig(key, value) {
  try {
    // Phase 1 stub
    return { ok: true, data: { updated: false } };
  } catch (error) {
    console.error('api: updateConfig failed —', error);
    return { ok: false, error: error.message };
  }
}
