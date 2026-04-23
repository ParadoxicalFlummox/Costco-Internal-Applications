/**
 * api.js — Public API layer for COMET.
 * VERSION: 0.4.0
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
 * PERF: Each public function logs its execution time and checks for timeout risk.
 * Long operations (>4.5 min) will log warnings to help identify bottlenecks.
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
 *     getCrossDeptHoursForWeek(employeeId, mondayDate)                      → { crossDeptAssignments, totalHours }
 *     getDeptSettings(deptName)                                             → { shifts, staffingReqs }
 *     saveDeptSettings(deptName, data)                                      → { saved }
 *     getSupervisorPeakConfig(deptName)                                     → { peakProfile, minCountPerPeak, ... }
 *     saveSupervisorPeakConfig(deptName, config)                            → { saved }
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
      return { ok: true, data: { weekSheetName: null, weekGrid: null, employeeList: [], staffingRequirements: {} } };
    }
    const staffingRequirements = loadStaffingRequirements(deptName); // settingsManager.js
    return {
      ok: true,
      data: {
        weekSheetName: result.weekSheetName,
        weekGrid: serializeWeekGrid_(result.weekGrid, result.employeeList),
        employeeList: serializeEmployeeList_(result.employeeList, result.weekGrid),
        staffingRequirements: staffingRequirements,
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
 * If re-generating an existing week, respects existing cell locks so manager
 * overrides are not overwritten by the algorithm.
 *
 * @param {string} deptName
 * @param {string} mondayDate — ISO date string "YYYY-MM-DD".
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function generateSchedule(deptName, mondayDate) {
  const scriptStart = Date.now();
  try {
    if (!deptName || !mondayDate) throw new Error('deptName and mondayDate are required.');
    const weekStartDate = new Date(mondayDate + 'T00:00:00');

    const { result } = logExecutionTime_('generateWeeklySchedule_', function() {
      return generateWeeklySchedule_(deptName, weekStartDate); // scheduleEngine.js
    });

    // Write to sheet via formatter.
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = workbook.getSheetByName(result.weekSheetName);
    const isFirstTimeGeneration = !sheet;
    if (!sheet) sheet = workbook.insertSheet(result.weekSheetName);

    // If this is a re-generation of an existing week, read the lock row from the sheet
    // and apply locks to the freshly generated grid. This preserves manager overrides.
    if (!isFirstTimeGeneration) {
      logExecutionTime_('Apply existing locks', function() {
        for (let ei = 0; ei < result.employeeList.length; ei++) {
          const baseRow = WEEK_SHEET.DATA_START_ROW + (ei * WEEK_SHEET.ROWS_PER_EMPLOYEE);
          const lockRow = sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
          for (let di = 0; di < WEEK_SHEET.DAYS_IN_WEEK; di++) {
            if (lockRow[di] === true && result.weekGrid[ei][di].type === 'SHIFT') {
              // This cell is locked — mark it as locked so phases skip it.
              result.weekGrid[ei][di].locked = true;
            }
          }
        }
      });
    }

    const staffingRequirements = loadStaffingRequirements(deptName); // settingsManager.js
    writeAndFormatSchedule(sheet, result.employeeList, result.weekGrid, staffingRequirements, weekStartDate, deptName); // formatter.js

    // Sort workbook tabs and clean up stale Week sheets.
    logExecutionTime_('Cleanup and sort sheets', function() {
      cleanupWeekSheets_();
      sortWorkbookSheets_();
    });

    const elapsed = Date.now() - scriptStart;
    console.log('[API] generateSchedule completed in ' + elapsed + 'ms');
    if (elapsed > MAX_SAFE_EXECUTION_MS) {
      console.warn('[TIMEOUT_RISK] generateSchedule took ' + elapsed + 'ms (limit is ' + MAX_SAFE_EXECUTION_MS + 'ms)');
    }

    return {
      ok: true,
      data: {
        weekSheetName: result.weekSheetName,
        weekGrid: serializeWeekGrid_(result.weekGrid, result.employeeList),
        employeeList: serializeEmployeeList_(result.employeeList, result.weekGrid),
        staffingRequirements: staffingRequirements,
      },
    };
  } catch (error) {
    console.error('api: generateSchedule failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns hours already scheduled for an employee across all departments for a given week.
 * Used for cross-department (split-shift) scheduling to show managers what hours are
 * already allocated in other departments.
 *
 * @param {string} employeeId  — Employee ID.
 * @param {string} mondayDate  — ISO date string "YYYY-MM-DD" for the Monday of the week.
 * @returns {{ ok: boolean, data?: { totalHours: number, assignments: Array }, error?: string }}
 *   assignments: Array of { deptName, hoursAssigned } for depts where employee is scheduled
 */
function getCrossDeptHoursForWeek(employeeId, mondayDate) {
  try {
    if (!employeeId || !mondayDate) throw new Error('employeeId and mondayDate are required.');
    const weekStartDate = new Date(mondayDate + 'T00:00:00');

    // Get employee data to look up their name and details
    const employees = getActiveEmployees_(); // ukgImport.js
    const employee = employees.find(e => String(e.id) === String(employeeId));
    if (!employee) {
      return { ok: true, data: { totalHours: 0, assignments: [] } };
    }

    // Scan all Week sheets for this week to find hours scheduled across departments
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = workbook.getSheetNames();
    const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
    const day   = String(weekStartDate.getDate()).padStart(2, '0');
    const year  = String(weekStartDate.getFullYear()).slice(-2);
    const weekBaseName = 'Week_' + month + '_' + day + '_' + year;

    const assignments = [];
    let totalHours = 0;

    sheetNames.forEach(function(sheetName) {
      if (!sheetName.startsWith(weekBaseName)) return;

      const prefix = weekBaseName + '_';
      if (!sheetName.startsWith(prefix)) return;
      const deptName = sheetName.substring(prefix.length);

      const sheet = workbook.getSheetByName(sheetName);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (!data || data.length < WEEK_SHEET.DATA_START_ROW) return;

      // Find the employee's SHIFT row and sum hours
      for (let rowIndex = WEEK_SHEET.DATA_START_ROW - 1; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        const rowLabel = row[WEEK_SHEET.COL_ROW_LABEL - 1];

        if (rowLabel !== 'SHIFT') continue;

        const empName = row[WEEK_SHEET.COL_EMPLOYEE_NAME - 1];
        if (!empName || empName.toString().trim() !== employee.name.toString().trim()) continue;

        // Found the employee in this dept. Get total hours from column J (COL_TOTAL_HOURS)
        const totalHoursCell = row[WEEK_SHEET.COL_TOTAL_HOURS - 1];
        let hoursInDept = 0;
        if (totalHoursCell && !isNaN(parseFloat(totalHoursCell))) {
          hoursInDept = parseFloat(totalHoursCell);
        }

        if (hoursInDept > 0) {
          assignments.push({ deptName: deptName, hoursAssigned: hoursInDept });
          totalHours += hoursInDept;
        }
        break;
      }
    });

    return { ok: true, data: { totalHours: totalHours, assignments: assignments } };
  } catch (error) {
    console.error('api: getCrossDeptHoursForWeek failed —', error);
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
 * Returns supervisor peak traffic configuration for a department.
 * If no config exists, returns default values from config.js.
 *
 * @param {string} deptName — Department name
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function getSupervisorPeakConfig(deptName) {
  try {
    if (!deptName) throw new Error('deptName is required.');
    let config = readSupervisorPeakConfig_(deptName); // settingsManager.js
    if (!config) {
      // Return default config from config.js
      config = {
        department: deptName,
        peakProfile: SUPERVISOR_RULES.defaultPeakProfile,
        minCountPerPeak: SUPERVISOR_RULES.minCountPerPeak,
        minCountPerValley: SUPERVISOR_RULES.minCountPerValley,
        peakThreshold: SUPERVISOR_RULES.peakThreshold,
      };
    }
    return { ok: true, data: config };
  } catch (error) {
    console.error('api: getSupervisorPeakConfig failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Saves supervisor peak traffic configuration for a department.
 *
 * @param {string} deptName — Department name
 * @param {Object} config — { peakProfile, minCountPerPeak, minCountPerValley, peakThreshold }
 * @returns {{ ok: boolean, data?: { saved: boolean }, error?: string }}
 */
function saveSupervisorPeakConfig(deptName, config) {
  try {
    if (!deptName || !config) throw new Error('deptName and config are required.');
    saveSupervisorPeakConfig_(deptName, config); // settingsManager.js
    return { ok: true, data: { saved: true } };
  } catch (error) {
    console.error('api: saveSupervisorPeakConfig failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Overrides a single cell in an existing Week sheet (VAC, RDO, or SHIFT),
 * marks it as locked so re-generations don't overwrite it, and returns the updated grid.
 *
 * For SHIFT overrides, writes directly without re-running phases. VAC/RDO overrides
 * are written as before and do NOT prevent phase re-runs on subsequent generations.
 *
 * @param {string} weekSheetName — Name of the existing Week sheet tab.
 * @param {string} employeeId    — Employee ID whose row to update.
 * @param {number} dayIndex      — 0 = Monday … 6 = Sunday.
 * @param {string} newType       — 'VAC', 'RDO', or 'SHIFT'.
 * @param {string} shiftName     — Required if newType='SHIFT'; ignored otherwise.
 *                                 Name of shift (must exist in dept settings) or '__CUSTOM__'.
 * @param {string} customDisplayText — Required if shiftName='__CUSTOM__'; display text like "10:00 AM - 3:00 PM".
 * @param {number} customPaidHours   — Required if shiftName='__CUSTOM__'; paid hours for the custom shift.
 * @returns {{ ok: boolean, data?: { weekGrid, employeeList, staffingRequirements }, error?: string }}
 */
function updateCellOverride(weekSheetName, employeeId, dayIndex, newType, shiftName, customDisplayText, customPaidHours) {
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

    const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate); // scheduleEngine.js
    const employeeIndex = employeeList.findIndex(e => String(e.employeeId) === String(employeeId));
    if (employeeIndex === -1) throw new Error('Employee ' + employeeId + ' not found in ' + deptName + ' roster.');

    const checkboxCol = WEEK_SHEET.COL_MONDAY + dayIndex;
    const baseRow = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);

    // Clear VAC/RDO for this cell, then apply the override.
    sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, checkboxCol).setValue(false);
    sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, checkboxCol).setValue(false);

    if (newType === 'VAC') {
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, checkboxCol).setValue(true);
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, checkboxCol).setValue(true);
    } else if (newType === 'RDO') {
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, checkboxCol).setValue(true);
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, checkboxCol).setValue(true);
    } else if (newType === 'SHIFT') {
      // For SHIFT: look up shift definition, write directly, and lock.
      if (!shiftName) throw new Error('shiftName is required when newType is "SHIFT".');

      let displayText = '';
      let paidHours = 0;

      if (shiftName === '__CUSTOM__') {
        // Custom shift: use provided display text and hours.
        if (!customDisplayText) throw new Error('customDisplayText is required for custom shifts.');
        if (customPaidHours == null) throw new Error('customPaidHours is required for custom shifts.');
        displayText = customDisplayText;
        paidHours = Number(customPaidHours);
      } else {
        // Standard shift: look up from settings.
        const employee = employeeList[employeeIndex];
        const shiftKey = shiftName + '|' + employee.status; // FT or PT qualifier

        // Find the shift definition matching both name and FT/PT status.
        const shiftTimingMap = buildShiftTimingMap(deptName); // settingsManager.js
        const shiftDef = shiftTimingMap[shiftKey];

        if (!shiftDef) {
          throw new Error('Shift "' + shiftName + '" not found for ' + employee.status + ' employees in ' + deptName + '.');
        }

        displayText = shiftDef.displayText;
        paidHours = shiftDef.paidHours;
      }

      // Write the shift display text and lock the cell.
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT, checkboxCol).setValue(displayText);
      sheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, checkboxCol).setValue(true);
    }

    // Re-read the grid from the updated sheet. readCheckboxStateFromSheet_ now reconstructs
    // SHIFT cells from the SHIFT row and applies lock status from the LOCK row.
    const weekGrid = readCheckboxStateFromSheet_(sheet, employeeList.length); // scheduleEngine.js
    const staffingReqs = loadStaffingRequirements(deptName); // settingsManager.js

    // Do NOT re-run phases. The override is now in place and locked, so phases will skip it.
    // Just run Phase 4 to assign roles to any SHIFT cells.
    runPhaseFourRoleAssignment_(weekGrid, employeeList, deptName);

    // Flush updated grid back to the sheet to reflect the role assignments.
    writeAndFormatSchedule(sheet, employeeList, weekGrid, staffingReqs, weekStartDate, deptName); // formatter.js

    return {
      ok: true,
      data: {
        weekGrid: serializeWeekGrid_(weekGrid, employeeList),
        employeeList: serializeEmployeeList_(employeeList, weekGrid),
        staffingRequirements: staffingReqs,
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
      manager: e.manager,
      scheduledShift: e.scheduledShift,
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
 * @param {boolean} dryRun    — When true, proposals are returned but nothing is written.
 * @param {boolean} sendEmail — When false (and not dryRun), CNs are logged but no
 *   emails are sent. When true, emails are sent normally. Ignored during dry run.
 * @returns {{ ok: boolean, data?: { proposals, issued, dryRun }, error?: string }}
 */
function runCNScan(dryRun, sendEmail) {
  try {
    const result = scanAndIssueCNs({ dryRun: !!dryRun, sendEmail: !!sendEmail }); // infractionEngine.js
    return { ok: true, data: {
      proposals: (result && result.proposals) || 0,
      issued:    (result && result.issued)    || 0,
      dryRun:    !!dryRun,
    }};
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
 * Returns the sorted unique department names.
 *
 * Unions two sources so departments appear in the dropdown as soon as
 * either source knows about them:
 *   1. Active employees in the Employees sheet (populated by UKG import).
 *   2. Existing Settings_[Dept] sheet tabs (created when settings are saved
 *      or auto-created on first generate — available even before any import).
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function getDepartments() {
  try {
    // Source 1 — Employees sheet
    const employees = getActiveEmployees_(); // ukgImport.js
    const fromEmployees = employees.map(e => e.department).filter(Boolean);

    // Source 2 — Settings_* sheet tabs
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const fromSettings = workbook.getSheets()
      .map(s => s.getName())
      .filter(n => n.startsWith(DEPT_SETTINGS_PREFIX)) // config.js
      .map(n => n.slice(DEPT_SETTINGS_PREFIX.length).trim())
      .filter(Boolean);

    const departments = [...new Set([...fromEmployees, ...fromSettings])].sort();
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
 * Each employee object includes an `attendanceSheetUrl` property (string or null)
 * so the frontend can render View as a direct <a> link without an extra round-trip.
 * window.open() after a callback is blocked by popup blockers in GAS iframes.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function getEmployeeList() {
  try {
    const workbook   = SpreadsheetApp.getActiveSpreadsheet();
    const workbookUrl = workbook.getUrl();
    const employees  = getAllEmployees_(); // ukgImport.js

    // Attach attendance sheet URL to each employee in one pass.
    employees.forEach(function(emp) {
      const tabName = emp.name + ' - ' + emp.id;
      const sheet   = workbook.getSheetByName(tabName);
      emp.attendanceSheetUrl = sheet
        ? workbookUrl + '#gid=' + sheet.getSheetId()
        : null;
    });

    return { ok: true, data: { employees } };
  } catch (error) {
    console.error('api: getEmployeeList failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the age of the employee roster (how many days since UKG import last ran).
 * Reads the ukgImportLastRan timestamp from the COMET Config sheet.
 *
 * @returns {{ ok: boolean, data?: { ageInDays, lastModified }, error?: string }}
 */
function getEmployeeSheetAge() {
  try {
    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME);
    if (!configSheet) throw new Error('COMET Config sheet not found.');

    // Search the config sheet for the "ukgImportLastRan" entry (rows 3+)
    const lastRow = configSheet.getLastRow();
    let lastImportTimestamp = null;

    if (lastRow >= 3) {
      const configData = configSheet.getRange(3, 1, lastRow - 2, 2).getValues();
      for (let i = 0; i < configData.length; i++) {
        if (String(configData[i][0] || '').trim() === 'ukgImportLastRan') {
          lastImportTimestamp = String(configData[i][1] || '').trim();
          break;
        }
      }
    }

    // If no import timestamp found, return a large age (assume very stale)
    if (!lastImportTimestamp) {
      return { ok: true, data: { ageInDays: 999, lastModified: 'Never' } };
    }

    // Parse the ISO timestamp and calculate days since
    const lastUpdated = new Date(lastImportTimestamp);
    if (isNaN(lastUpdated.getTime())) {
      // Invalid date format
      return { ok: true, data: { ageInDays: 999, lastModified: 'Invalid timestamp' } };
    }

    const now = new Date();
    const ageInMs = now - lastUpdated;
    const ageInDays = Math.floor(ageInMs / (1000 * 60 * 60 * 24));

    return { ok: true, data: { ageInDays, lastModified: lastImportTimestamp } };
  } catch (error) {
    console.error('api: getEmployeeSheetAge failed —', error);
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
 * Creates attendance controller sheet tabs for all active employees for the
 * given calendar year. Existing tabs are skipped.
 *
 * @param {number} year — e.g. 2026
 * @returns {{ ok: boolean, data?: { created, skipped }, error?: string }}
 */
function generateAttendanceTabs(year) {
  try {
    if (!year) year = new Date().getFullYear();
    const result = generateAttendanceControllerTabs_(Number(year)); // tabManager.js
    return { ok: true, data: result };
  } catch (error) {
    console.error('api: generateAttendanceTabs failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the URL of an employee's attendance controller tab so the
 * frontend can open it directly in a new browser tab.
 *
 * @param {string} employeeId
 * @returns {{ ok: boolean, data?: { url: string|null, message?: string }, error?: string }}
 */
function getAttendanceSheetUrl(employeeId) {
  try {
    if (!employeeId) throw new Error('employeeId is required.');
    const workbook  = SpreadsheetApp.getActiveSpreadsheet();
    const employees = getAllEmployees_(); // ukgImport.js
    const emp = employees.find(e => String(e.id) === String(employeeId));
    if (!emp) throw new Error('Employee ' + employeeId + ' not found.');

    const tabName = emp.name + ' - ' + emp.id;
    const sheet   = workbook.getSheetByName(tabName);
    if (!sheet) {
      return { ok: true, data: {
        url: null,
        message: 'No attendance controller tab found for ' + emp.name + '. Use "Generate Attendance Tabs" in Admin first.',
      }};
    }
    const url = workbook.getUrl() + '#gid=' + sheet.getSheetId();
    return { ok: true, data: { url } };
  } catch (error) {
    console.error('api: getAttendanceSheetUrl failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Returns the names and Monday dates of all existing Week sheets for a
 * department within the calendar month containing the given date.
 * Used by the "Load Existing" view to show all weeks at once.
 *
 * @param {string} deptName
 * @param {string} monthDate — any ISO date string within the target month
 * @returns {{ ok: boolean, data?: { weeks: Array<{weekSheetName,mondayDate}> }, error?: string }}
 */
function getScheduleWeeksForMonth(deptName, monthDate) {
  try {
    if (!deptName) throw new Error('deptName is required.');
    const anchor    = new Date((monthDate || new Date().toISOString().slice(0, 10)) + 'T00:00:00');
    const year      = anchor.getFullYear();
    const month     = anchor.getMonth(); // 0-based

    const workbook = SpreadsheetApp.getActiveSpreadsheet();
    const prefix   = 'Week_';
    const weeks    = [];

    workbook.getSheets().forEach(function(sheet) {
      const name = sheet.getName();
      if (!name.startsWith(prefix)) return;

      // Sheet name format: Week_MM_DD_YY_DeptName
      const parts = name.split('_');
      if (parts.length < 5) return;

      // Check department suffix matches
      const sheetDept = parts.slice(4).join('_');
      if (sheetDept.toLowerCase() !== deptName.toLowerCase()) return;

      // Parse date from Week_MM_DD_YY
      const sheetDate = new Date('20' + parts[3] + '-' + parts[1] + '-' + parts[2] + 'T00:00:00');
      if (isNaN(sheetDate.getTime())) return;

      // Include if the Monday falls within the target month
      if (sheetDate.getFullYear() === year && sheetDate.getMonth() === month) {
        weeks.push({
          weekSheetName: name,
          mondayDate:    sheetDate.toISOString().slice(0, 10),
        });
      }
    });

    // Sort chronologically
    weeks.sort(function(a, b) { return a.mondayDate.localeCompare(b.mondayDate); });
    return { ok: true, data: { weeks } };
  } catch (error) {
    console.error('api: getScheduleWeeksForMonth failed —', error);
    return { ok: false, error: error.message };
  }
}

/**
 * Hides Week sheets older than the current week and permanently deletes
 * Week sheets older than 2 months. Called automatically after generation.
 *
 * Hidden sheets remain accessible via right-click → Show Sheet if needed.
 * Deletion is permanent — only sheets confirmed to be more than ~9 weeks
 * old are removed.
 *
 * @returns {{ hidden: number, deleted: number }}
 */
function cleanupWeekSheets_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const now      = new Date();

  // Monday of the current week
  const today    = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const dow      = today.getDay();
  const daysToMon = (dow === 0) ? -6 : 1 - dow;
  const thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() + daysToMon);

  const twoMonthsAgo = new Date(thisMonday);
  twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 2);

  let hidden = 0;
  let deleted = 0;

  workbook.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    if (!name.startsWith('Week_')) return;

    const parts = name.split('_');
    if (parts.length < 4) return;

    const sheetDate = new Date('20' + parts[3] + '-' + parts[1] + '-' + parts[2] + 'T00:00:00');
    if (isNaN(sheetDate.getTime())) return;

    if (sheetDate < twoMonthsAgo) {
      // Older than 2 months — delete permanently
      try {
        workbook.deleteSheet(sheet);
        deleted++;
      } catch (e) {
        console.warn('cleanupWeekSheets_: could not delete "' + name + '": ' + e.message);
      }
    } else if (sheetDate < thisMonday) {
      // Older than this week but within 2 months — hide
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
        hidden++;
      }
    }
  });

  return { hidden, deleted };
}

/**
 * Generates the next fiscal year's attendance controller workbook.
 * Kept as a stub until Phase 6 multi-workbook support is added.
 *
 * @returns {{ ok: boolean, data?: object, error?: string }}
 */
function generateNextYearWorkbook() {
  try {
    return { ok: true, data: { url: null, message: 'Use "Generate Attendance Tabs" with the target year instead.' } };
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


// ---------------------------------------------------------------------------
// Internal — Workbook Tab Sorting
// ---------------------------------------------------------------------------

/**
 * Sorts all sheet tabs in the active workbook into logical groups so managers
 * can quickly find what they need without scrolling through a mix of types.
 *
 * Group order:
 *   0 — Employees (master roster)
 *   1 — COMET Config
 *   2 — Active CNs
 *   3 — CN_Log
 *   4 — (Expired CNs)
 *   5 — Settings_* sheets (alphabetical by dept name)
 *   6 — Week_* schedule sheets (newest first — desc by sheet name)
 *   7 — Call Log * sheets (newest first)
 *   8 — Everything else
 *   9 — Attendance controller tabs (Last, First - ID)
 */
function sortWorkbookSheets_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = workbook.getSheets();

  function groupOf(name) {
    if (name === EMPLOYEES_SHEET_NAME)     return 0;
    if (name === COMET_CONFIG_SHEET_NAME)  return 1;
    if (name === ACTIVE_CNS_SHEET_NAME)    return 2;
    if (name === CN_LOG_SHEET_NAME)        return 3;
    if (name === EXPIRED_CNS_SHEET_NAME)   return 4;
    if (name.startsWith(DEPT_SETTINGS_PREFIX)) return 5;
    if (name.startsWith('Week_'))          return 6;
    if (name.startsWith('Call Log'))       return 7;
    if (EMPLOYEE_TAB_PATTERN.test(name))   return 9;
    return 8;
  }

  const sorted = sheets.slice().sort(function(a, b) {
    const ga = groupOf(a.getName());
    const gb = groupOf(b.getName());
    if (ga !== gb) return ga - gb;
    // Week sheets: newest first (descending by name — Week_MM_DD_YY sorts correctly)
    if (ga === 6) return b.getName().localeCompare(a.getName());
    // Call Log sheets: newest first
    if (ga === 7) return b.getName().localeCompare(a.getName());
    // Everything else: alphabetical
    return a.getName().localeCompare(b.getName());
  });

  sorted.forEach(function(sheet, index) {
    workbook.setActiveSheet(sheet);
    workbook.moveActiveSheet(index + 1);
  });
}
