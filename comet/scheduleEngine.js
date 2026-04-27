/**
 * scheduleEngine.js — Core schedule generation algorithm for COMET.
 * VERSION: 0.6.2
 *
 * CHANGES FROM SOURCE:
 *   - loadRosterSortedBySeniority_() now reads from getActiveEmployees_() (ukgImport.js)
 *     filtered by department, instead of reading the Roster sheet directly.
 *   - generateWeeklySchedule() accepts a deptName parameter.
 *   - normalizeDeptName_() is defined here (was implicit in the original).
 *   - All four phases, coverage map functions, and utility functions are unchanged.
 *   - PERF: Added profiling wrappers to each phase for performance monitoring
 *   - SPLIT_SHIFT: Integrated multi-department scheduling; Phase 0 loads cross-dept hours,
 *     Phase 2 enforces reduced hour budgets based on hours already scheduled elsewhere.
 *   - TRAFFIC_HEATMAP: Integrated traffic heatmap system (v0.5.0)
 *     - Phases 1-4 now accept stagger map and use staggered start times
 *     - Pool members partitioned and scheduled with traffic-aware staffing
 *     - Phase 5 (old supervisor scheduling) removed; replaced by pool scheduling
 *
 * THE FIVE PHASES:
 *   Phase 0 — Bootstrap: Load roster for dept, initialize grid, load cross-dept hours, stamp vacation locks.
 *   Phase 1 — Preference Assignment: Honor day-off prefs and shift prefs, seniority order (regular employees only, with stagger).
 *   Phase 2 — Minimum Hour Enforcement: Add shifts until weekly minimum is met (regular employees only, with stagger).
 *   Phase 3 — Gap Resolution: Fill uncovered time slots by reassigning or adding employees (regular employees only, with stagger).
 *   Phase 4 — Role Assignment: Stamp primaryRole onto SHIFT cells (regular employees only).
 *   Phase 5 — Pool Scheduling: Assign pool members to shifts based on traffic level and stagger.
 */


// ---------------------------------------------------------------------------
// Top-level Entry Point
// ---------------------------------------------------------------------------

/**
 * Generates a complete weekly schedule for the given department and Monday.
 *
 * @param {string} deptName       — Department name (must match employees' department field).
 * @param {Date}   weekStartDate  — The Monday of the week to generate.
 * @returns {{ weekSheetName: string, weekGrid: Array, employeeList: Array }}
 */
function generateWeeklySchedule_(deptName, weekStartDate) {
  const shiftTimingMap      = buildShiftTimingMap(deptName);      // settingsManager.js
  const staffingRequirements = loadStaffingRequirements(deptName); // settingsManager.js

  // --- TRAFFIC HEATMAP INTEGRATION (v0.5.0) ---
  const heatmapConfig = loadTrafficHeatmapConfig_(deptName); // settingsManager.js
  let dayTrafficLevels = {}; // Default: all Moderate if heatmap disabled
  let staggerMap = {}; // Default: empty (phases use anchor times)
  let poolMembers = [];
  let regularEmployees = [];

  if (heatmapConfig && heatmapConfig.enabled) {
    // Classify traffic for each day
    dayTrafficLevels = classifyDayTrafficLevels_(heatmapConfig, weekStartDate); // trafficHeatmapEngine.js

    // Pre-compute staggered start times for each shift on each day
    staggerMap = buildStaggeredStartMap_(shiftTimingMap, dayTrafficLevels, weekStartDate); // trafficHeatmapEngine.js

    console.log('Traffic heatmap enabled: ' + Object.keys(dayTrafficLevels).map(d => d + '=' + dayTrafficLevels[d]).join(', '));
    console.log('Stagger map computed for ' + Object.keys(staggerMap).length + ' days');
  } else {
    // Heatmap disabled: default to all Moderate, no stagger
    DAY_NAMES_IN_ORDER.forEach(function(dayName) {
      dayTrafficLevels[dayName] = 'Moderate';
    });
    console.log('Traffic heatmap disabled; using static staffing');
  }

  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate);
  if (employeeList.length === 0) {
    throw new Error('No active employees found for department "' + deptName + '".');
  }

  // Partition employees into pool and regular if heatmap enabled
  if (heatmapConfig && heatmapConfig.enabled) {
    const partitioned = partitionPoolMembers_(employeeList, heatmapConfig); // trafficHeatmapEngine.js
    poolMembers = partitioned.poolMembers;
    regularEmployees = partitioned.regularEmployees;
    console.log('Pool partitioned: ' + poolMembers.length + ' pool, ' + regularEmployees.length + ' regular employees');
  } else {
    regularEmployees = employeeList;
  }

  const weekGrid = initializeWeekGrid_(employeeList, weekStartDate);

  // PERF: Wrap each phase with execution time logging
  logExecutionTime_('Phase 1: Preference Assignment (' + regularEmployees.length + ' regular employees)', function() {
    runPhaseOnePreferenceAssignment_(weekGrid, regularEmployees, shiftTimingMap, staffingRequirements, weekStartDate, staggerMap, dayTrafficLevels);
  });

  logExecutionTime_('Phase 2: Minimum Hour Enforcement', function() {
    runPhaseTwoHourEnforcement_(weekGrid, regularEmployees, shiftTimingMap, staggerMap, dayTrafficLevels);
  });

  logExecutionTime_('Phase 3: Gap Resolution', function() {
    runPhaseThreeGapResolution_(weekGrid, regularEmployees, shiftTimingMap, staffingRequirements, staggerMap, dayTrafficLevels);
  });

  logExecutionTime_('Phase 4: Role Assignment', function() {
    runPhaseFourRoleAssignment_(weekGrid, regularEmployees, deptName);
  });

  logExecutionTime_('Phase 5: Pool Scheduling', function() {
    if (poolMembers.length > 0) {
      runPhaseFivePoolScheduling_(weekGrid, employeeList, poolMembers, heatmapConfig, dayTrafficLevels, staggerMap, shiftTimingMap, weekStartDate);
    }
  });

  // Build a set of pool member IDs for the formatter to visually distinguish pool section
  const poolMemberIdSet = new Set();
  poolMembers.forEach(function(emp) {
    if (emp.id) poolMemberIdSet.add(emp.id);
  });

  return {
    weekSheetName: generateWeekSheetName_(weekStartDate, deptName),
    weekGrid:      weekGrid,
    employeeList:  employeeList,
    poolMemberIds: poolMemberIdSet,
  };
}

/**
 * Reads the schedule grid back from an existing Week sheet.
 * Used by getScheduleForWeek() to return an already-generated schedule.
 *
 * @param {string} deptName
 * @param {Date}   weekStartDate
 * @returns {{ weekSheetName, weekGrid, employeeList } | null}
 */
function readExistingWeekSchedule_(deptName, weekStartDate) {
  const sheetName = generateWeekSheetName_(weekStartDate, deptName);
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = workbook.getSheetByName(sheetName);
  if (!sheet) return null;

  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate);
  if (employeeList.length === 0) return null;

  // Read checkbox state and rebuild a grid representation.
  const weekGrid = readCheckboxStateFromSheet_(sheet, employeeList.length);

  // Re-read SHIFT row text to populate displayText.
  const shiftTimingMap = buildShiftTimingMap(deptName);
  employeeList.forEach(function(employee, employeeIndex) {
    const baseRow = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const shiftRowValues = sheet
      .getRange(baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .getValues()[0];

    shiftRowValues.forEach(function(cellText, dayIndex) {
      const current = weekGrid[employeeIndex][dayIndex];
      if (current.type === 'OFF' && cellText && cellText !== 'OFF' && cellText !== '' &&
          cellText !== 'VAC' && cellText !== 'RDO') {
        // This is a SHIFT cell — reconstruct it from the display text.
        weekGrid[employeeIndex][dayIndex] = createDayAssignment_(
          'SHIFT', null, 0, false, String(cellText)
        );
      }
    });
  });

  return { weekSheetName: sheetName, weekGrid, employeeList };
}


// ---------------------------------------------------------------------------
// Phase 0: Roster Loading
// ---------------------------------------------------------------------------

/**
 * Reads active employees for the given department from the Employees sheet
 * and returns them sorted by seniority (descending).
 *
 * Maps COMET Employees sheet columns (A–N) to the engine's EmployeeRecord shape.
 * Also loads cross-department hours for employees with secondary departments.
 *
 * @param {string} deptName       — Department name
 * @param {Date}   weekStartDate  — The Monday of the week (for cross-dept lookups)
 * @returns {Array<EmployeeRecord>}
 */
function loadRosterSortedBySeniority_(deptName, weekStartDate) {
  const normalizedTarget = normalizeDeptName_(deptName);

  const employees = getActiveEmployees_(); // ukgImport.js — Active employees only

  // Get the Employees sheet to read secondary departments (column N)
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const employeesSheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
  const employeesData = employeesSheet ? employeesSheet.getDataRange().getValues() : [];

  const deptEmployees = employees
    .filter(emp => normalizeDeptName_(emp.department) === normalizedTarget)
    .map(function(emp, _index) {
      const qualifiedShiftList = parseQualifiedShiftList_(
        emp.qualifiedShifts, emp.preferredShift
      );
      const vacationDateStrings = parseVacationDateStrings_(emp.vacationDates);

      // Parse hireDate — stored as "MM/dd/yyyy" string from ukgImport.js
      let hireDate = new Date();
      if (emp.hireDate) {
        const parsed = new Date(emp.hireDate);
        if (!isNaN(parsed.getTime())) hireDate = parsed;
      }

      // Calculate seniority rank from hire date using the SENIORITY formula in config.js.
      // Column M is not used — calculating live ensures the rank always reflects the actual
      // formula (FT_BASE/PT_BASE + days before REFERENCE_DATE) without requiring a setup re-run.
      const seniorityReferenceDate = new Date(SENIORITY.REFERENCE_DATE_STRING);
      const seniorityBase = emp.ftpt === 'FT' ? SENIORITY.FT_BASE : SENIORITY.PT_BASE;
      const seniorityDaysFromHire = Math.max(0, Math.floor(
        (seniorityReferenceDate - hireDate) / (1000 * 60 * 60 * 24)
      ));
      const calculatedSeniorityRank = seniorityBase + seniorityDaysFromHire;

      // Load secondary departments from Employees sheet column N
      let secondaryDepartments = [];
      let crossDeptHoursAlreadyScheduled = 0;
      let crossDeptScheduledDays = null; // null = not a cross-dept employee
      if (employeesData && employeesData.length > 0) {
        // Find the row for this employee (match by name or ID)
        for (let rowIdx = EMPLOYEES_DATA_START_ROW - 1; rowIdx < employeesData.length; rowIdx++) {
          const sheetRow = employeesData[rowIdx];
          const sheetEmpId = sheetRow[EMPLOYEE_COLUMN.ID - 1];
          if (sheetEmpId && sheetEmpId.toString().trim() === emp.id.toString().trim()) {
            // Found the employee's row. Read secondary departments from column N
            const secondaryDeptString = sheetRow[EMPLOYEE_COLUMN.SECONDARY_DEPARTMENTS - 1] || '';
            if (secondaryDeptString && secondaryDeptString.toString().trim()) {
              secondaryDepartments = secondaryDeptString
                .toString()
                .split(',')
                .map(function(d) { return normalizeDeptName_(d); })
                .filter(Boolean);
            }
            break;
          }
        }

        // If the employee has secondary departments, query for cross-dept hours and scheduled days.
        // crossDeptScheduledDays drives handoff scheduling: when this employee appears in a
        // secondary department's schedule, Phases 1–3 only assign shifts on days they are
        // already working in their home department — mirroring the home dept days off rather
        // than filling their open days.
        if (secondaryDepartments.length > 0 && weekStartDate) {
          crossDeptHoursAlreadyScheduled = getCrossDeptHoursForWeek_(emp, weekStartDate, deptName);
          crossDeptScheduledDays = getCrossDeptScheduledDays_(emp, weekStartDate, deptName);
        }
      }

      return {
        name:                            emp.name,
        employeeId:                      emp.id,
        hireDate:                        hireDate,
        status:                          emp.ftpt || 'PT',           // FT or PT (col F)
        dayOffPreferenceOne:             emp.dayOffPrefOne || '',
        dayOffPreferenceTwo:             emp.dayOffPrefTwo || '',
        preferredShift:                  emp.preferredShift  || '',
        qualifiedShifts:                 qualifiedShiftList,
        vacationDateStrings:             vacationDateStrings,
        seniorityRank:                   calculatedSeniorityRank,
        department:                      normalizeDeptName_(emp.department),
        primaryRole:                     emp.role || '',
        secondaryDepartments:            secondaryDepartments,
        crossDeptHoursAlreadyScheduled:  crossDeptHoursAlreadyScheduled,
        crossDeptScheduledDays:          crossDeptScheduledDays,
      };
    })
    .filter(function(employee) {
      // Exclude employees whose primary role is "manager" — they manage the schedule
      // rather than appear in it. Role field is comma-separated; first entry is primary.
      var firstRole = (employee.primaryRole || '').split(',')[0].trim().toLowerCase();
      return firstRole !== 'manager';
    });

  // Supervisors first, then remaining employees sorted alphabetically by primary role,
  // then by seniority within each role group.
  deptEmployees.sort(compareEmployeesForScheduleOrder_);
  return deptEmployees;
}

/**
 * Creates the initial WeekGrid with all cells set to OFF, then stamps vacation locks.
 */
function initializeWeekGrid_(employeeList, weekStartDate) {
  const weekGrid = [];
  employeeList.forEach(function(employee, employeeIndex) {
    weekGrid[employeeIndex] = [];
    DAY_NAMES_IN_ORDER.forEach(function(_dayName, _dayIndex) {
      weekGrid[employeeIndex][_dayIndex] = createDayAssignment_('OFF', null, 0, false);
    });
    applyVacationLocksForEmployee_(weekGrid, employeeIndex, employee, weekStartDate);
  });
  return weekGrid;
}

function applyVacationLocksForEmployee_(weekGrid, employeeIndex, employee, weekStartDate) {
  employee.vacationDateStrings.forEach(function(vacationDateString) {
    const vacationDate = parseVacationDateString_(vacationDateString, weekStartDate);
    if (!vacationDate) return;
    const dayIndex = getDayIndexForDate_(vacationDate, weekStartDate);
    if (dayIndex === -1) return;
    weekGrid[employeeIndex][dayIndex] = createDayAssignment_('VAC', null, 0, true);
  });
}

/**
 * Scans existing Week schedules for other departments to find hours already assigned
 * to an employee in the given week. Used for cross-dept scheduling (split-shift).
 *
 * @param {object}  employee        — Employee object with id and name fields
 * @param {Date}    weekStartDate   — The Monday of the week
 * @param {string}  excludeDept     — Department to exclude from scan (e.g., current dept)
 * @returns {number} Total hours already scheduled in other departments
 */
function getCrossDeptHoursForWeek_(employee, weekStartDate, excludeDept) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = workbook.getSheets().map(function(s) { return s.getName(); });
  const normalizedExcludeDept = normalizeDeptName_(excludeDept);
  const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
  const day   = String(weekStartDate.getDate()).padStart(2, '0');
  const year  = String(weekStartDate.getFullYear()).slice(-2);
  const weekBaseName = 'Week_' + month + '_' + day + '_' + year;

  let totalCrossDeptHours = 0;

  sheetNames.forEach(function(sheetName) {
    // Match sheets like "Week_MM_DD_YY_DeptName"
    if (!sheetName.startsWith(weekBaseName)) return;

    // Extract department name (everything after "Week_MM_DD_YY_")
    const prefix = weekBaseName + '_';
    if (!sheetName.startsWith(prefix)) return;
    const deptName = sheetName.substring(prefix.length);
    const normalizedDeptName = normalizeDeptName_(deptName);

    // Skip the department we're currently generating for
    if (normalizedDeptName === normalizedExcludeDept) return;

    const sheet = workbook.getSheetByName(sheetName);
    if (!sheet) return;

    // Read all data from the sheet
    const data = sheet.getDataRange().getValues();
    if (!data || data.length < WEEK_SHEET.DATA_START_ROW) return;

    // Scan for employee rows. Each employee has ROWS_PER_EMPLOYEE (5) rows.
    // Look for SHIFT rows (row offset 2 within each employee block).
    for (let rowIndex = WEEK_SHEET.DATA_START_ROW - 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      const rowLabel = row[WEEK_SHEET.COL_ROW_LABEL - 1];

      // Only look at SHIFT rows
      if (rowLabel !== 'SHIFT') continue;

      // Get employee name from column B (COL_EMPLOYEE_NAME - 1 = 1)
      const empName = row[WEEK_SHEET.COL_EMPLOYEE_NAME - 1];
      if (!empName || empName.toString().trim() !== employee.name.toString().trim()) continue;

      // Found the employee's SHIFT row in this dept. Sum hours from Mon-Sun (columns C-I).
      // Column C is COL_MONDAY = 3, so 0-indexed it's column 2
      let hoursInDept = 0;
      for (let dayCol = WEEK_SHEET.COL_MONDAY - 1; dayCol < WEEK_SHEET.COL_MONDAY - 1 + WEEK_SHEET.DAYS_IN_WEEK; dayCol++) {
        const cellValue = row[dayCol];
        // Extract hours from shift text like "8:00 AM - 4:30 PM" or from paidHours if numeric
        // For now, trust that the total is in the J column (COL_TOTAL_HOURS)
      }

      // Use the total hours cell (column J = COL_TOTAL_HOURS, 0-indexed is 9)
      const totalHoursCell = row[WEEK_SHEET.COL_TOTAL_HOURS - 1];
      if (totalHoursCell && !isNaN(parseFloat(totalHoursCell))) {
        hoursInDept = parseFloat(totalHoursCell);
      }

      totalCrossDeptHours += hoursInDept;
      break; // Found this employee in this dept, move to next dept
    }
  });

  return totalCrossDeptHours;
}


/**
 * Scans existing Week schedules for other departments to find which days an
 * employee already has a SHIFT assigned. Used for handoff scheduling: when a
 * cross-dept employee appears in a secondary department, Phases 1–3 only assign
 * shifts on days they are already working in their home department — their home
 * dept days off become the secondary dept days off automatically.
 *
 * Returns null if no other-dept schedule exists yet for this week, which means
 * the home dept hasn't been generated yet and handoff scheduling can't run.
 * In that case the engine falls back to normal (fill open days) behavior.
 *
 * @param {object}  employee        — Employee object with name field
 * @param {Date}    weekStartDate   — The Monday of the week
 * @param {string}  excludeDept     — Department to exclude (current dept being generated)
 * @returns {{ Monday: bool, Tuesday: bool, ... }|null}
 */
function getCrossDeptScheduledDays_(employee, weekStartDate, excludeDept) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = workbook.getSheets().map(function(s) { return s.getName(); });
  const normalizedExcludeDept = normalizeDeptName_(excludeDept);
  const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
  const day   = String(weekStartDate.getDate()).padStart(2, '0');
  const year  = String(weekStartDate.getFullYear()).slice(-2);
  const weekBaseName = 'Week_' + month + '_' + day + '_' + year + '_';

  // Start with all days off — any day found as SHIFT in another dept flips to true
  const scheduledDays = {
    Monday: false, Tuesday: false, Wednesday: false,
    Thursday: false, Friday: false, Saturday: false, Sunday: false,
  };
  let foundInAnyDept = false;

  sheetNames.forEach(function(sheetName) {
    if (!sheetName.startsWith(weekBaseName)) return;
    const deptNameFromSheet = sheetName.substring(weekBaseName.length);
    if (normalizeDeptName_(deptNameFromSheet) === normalizedExcludeDept) return;

    const sheet = workbook.getSheetByName(sheetName);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (!data || data.length < WEEK_SHEET.DATA_START_ROW) return;

    for (let rowIndex = WEEK_SHEET.DATA_START_ROW - 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      if (row[WEEK_SHEET.COL_ROW_LABEL - 1] !== 'SHIFT') continue;
      const empName = row[WEEK_SHEET.COL_EMPLOYEE_NAME - 1];
      if (!empName || empName.toString().trim() !== employee.name.toString().trim()) continue;

      // Found the employee's SHIFT row — check each day column for a non-empty value
      foundInAnyDept = true;
      DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
        const cellValue = row[WEEK_SHEET.COL_MONDAY - 1 + dayIndex];
        if (cellValue && String(cellValue).trim() !== '') {
          scheduledDays[dayName] = true;
        }
      });
      break;
    }
  });

  // Return null if no other-dept schedule was found — signals that the home dept
  // hasn't been generated yet and handoff mode cannot be applied this run.
  return foundInAnyDept ? scheduledDays : null;
}


// ---------------------------------------------------------------------------
// Phase 1: Preference Assignment
// ---------------------------------------------------------------------------

function runPhaseOnePreferenceAssignment_(weekGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate, staggerMap, dayTrafficLevels) {
  grantRequestedDaysOff_(weekGrid, employeeList, staffingRequirements);
  assignPreferredShifts_(weekGrid, employeeList, shiftTimingMap, staggerMap, dayTrafficLevels);
  enforceMinimumDaysOff_(weekGrid, employeeList);
}

function grantRequestedDaysOff_(weekGrid, employeeList, staffingRequirements) {
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const minimumStaffRequired = (staffingRequirements[dayName] && staffingRequirements[dayName].value) || 0;
    let availableStaffCount = 0;
    employeeList.forEach(function(_emp, ei) {
      if (weekGrid[ei][dayIndex].type !== 'VAC') availableStaffCount++;
    });
    employeeList.forEach(function(employee, ei) {
      const currentCell = weekGrid[ei][dayIndex];
      if (currentCell.locked) return;
      const requested = employee.dayOffPreferenceOne === dayName || employee.dayOffPreferenceTwo === dayName;
      if (!requested) return;
      if (availableStaffCount > minimumStaffRequired) {
        weekGrid[ei][dayIndex] = createDayAssignment_('RDO', null, 0, false);
        availableStaffCount--;
      }
    });
  });
}

function assignPreferredShifts_(weekGrid, employeeList, shiftTimingMap, staggerMap, dayTrafficLevels) {
  // Track stagger position per shift type per day (for popping next available time)
  const staggerPositions = {};

  employeeList.forEach(function(employee, ei) {
    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'OFF') return;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) return;
      // Handoff mode: cross-dept employee in a secondary dept schedule should only
      // receive shifts on days they are already working in their home department.
      // If crossDeptScheduledDays is set but false for this day, skip it — their
      // home dept scheduled that day as off and we must honour that.
      if (employee.crossDeptScheduledDays && !employee.crossDeptScheduledDays[dayName]) return;
      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (!shiftDef) return;

      // Default display text uses the day-specific anchor (sat/sun overrides applied here)
      const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName); // settingsManager.js
      let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);

      // Stagger map overrides the anchor when flex is enabled
      if (staggerMap && staggerMap[dayName]) {
        const shiftKey = shiftDef.name + '|' + shiftDef.status;
        const startTimes = staggerMap[dayName][shiftKey];
        if (startTimes && startTimes.length > 0) {
          const posKey = dayName + '|' + shiftKey;
          if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
          const staggerStartTime = startTimes[staggerPositions[posKey] % startTimes.length];
          staggerPositions[posKey]++;
          displayText = buildShiftDisplayText_(staggerStartTime, shiftDef.paidHours, shiftDef.hasLunch);
        }
      }

      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
    });
  });
}

function enforceMinimumDaysOff_(weekGrid, employeeList) {
  const maxWorkingDays = WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF; // config.js
  for (let ei = 0; ei < employeeList.length; ei++) {
    const workingDayIndices = [];
    for (let di = 0; di < WEEK_SHEET.DAYS_IN_WEEK; di++) {
      if (weekGrid[ei][di].type === 'SHIFT') workingDayIndices.push(di);
    }
    const excess = workingDayIndices.length - maxWorkingDays;
    if (excess <= 0) continue;
    const scoredDays = workingDayIndices.map(function(di) {
      let others = 0;
      for (let oi = 0; oi < employeeList.length; oi++) {
        if (oi !== ei && weekGrid[oi][di].type === 'SHIFT') others++;
      }
      return { dayIndex: di, coverage: others };
    });
    scoredDays.sort(function(a, b) { return b.coverage - a.coverage; });
    for (let i = 0; i < excess; i++) {
      weekGrid[ei][scoredDays[i].dayIndex] = createDayAssignment_('OFF', null, 0, false);
    }
  }
}


// ---------------------------------------------------------------------------
// Phase 2: Minimum Hour Enforcement
// ---------------------------------------------------------------------------

function runPhaseTwoHourEnforcement_(weekGrid, employeeList, shiftTimingMap, staggerMap, dayTrafficLevels) {
  // Track stagger position per shift type per day (for popping next available time)
  const staggerPositions = {};

  employeeList.forEach(function(employee, ei) {
    const weeklyMin = employee.status === 'FT' ? HOUR_RULES.FT_MIN : (employee.status === 'LPT' ? HOUR_RULES.LPT_MIN : HOUR_RULES.PT_MIN); // config.js
    const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
    // Reduce effective max by hours already scheduled in other departments (split-shift)
    const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
    let currentHours = getWeeklyHours_(weekGrid, ei);
    if (currentHours >= weeklyMin) return;
    const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
    if (!shiftDef) return;
    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      if (currentHours >= weeklyMin) return;
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'OFF') return;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) return;
      // Handoff mode: only schedule on days already worked in home department.
      if (employee.crossDeptScheduledDays && !employee.crossDeptScheduledDays[dayName]) return;
      if (countWorkingDays_(weekGrid, ei) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) return;
      if (currentHours + shiftDef.paidHours > effectiveMax) return;

      // Default display text uses the day-specific anchor (sat/sun overrides applied here)
      const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName); // settingsManager.js
      let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);

      // Stagger map overrides the anchor when flex is enabled
      if (staggerMap && staggerMap[dayName]) {
        const shiftKey = shiftDef.name + '|' + shiftDef.status;
        const startTimes = staggerMap[dayName][shiftKey];
        if (startTimes && startTimes.length > 0) {
          const posKey = dayName + '|' + shiftKey;
          if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
          const staggerStartTime = startTimes[staggerPositions[posKey] % startTimes.length];
          staggerPositions[posKey]++;
          displayText = buildShiftDisplayText_(staggerStartTime, shiftDef.paidHours, shiftDef.hasLunch);
        }
      }

      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
      currentHours += shiftDef.paidHours;
    });
  });
}


// ---------------------------------------------------------------------------
// Phase 3: Gap Resolution
// ---------------------------------------------------------------------------

function runPhaseThreeGapResolution_(weekGrid, employeeList, shiftTimingMap, staffingRequirements, staggerMap, dayTrafficLevels) {
  // TODO: Phase 3 currently uses default shift times. For Phase 2 enhancement,
  // integrate staggerMap into cascade assignments to use staggered start times.
  // This requires tracking stagger positions through multiple cascades.

  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const coverageWindow = COVERAGE_WINDOW[dayName] || { startMinute: 240, endMinute: 1410 }; // config.js
    const windowStartSlot = Math.max(0, Math.floor((coverageWindow.startMinute - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
    const windowEndSlot   = Math.min(COVERAGE.SLOT_COUNT, Math.floor((coverageWindow.endMinute - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));

    let dayCoverage = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);
    if (!hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot)) return;

    // Pre-compute shift definitions for all qualified shift + status combos to reduce repeated lookups
    const shiftDefCache_ = {};
    employeeList.forEach(function(emp) {
      emp.qualifiedShifts.forEach(function(shiftName) {
        const key = shiftName + '|' + emp.status;
        if (!shiftDefCache_[key]) shiftDefCache_[key] = shiftTimingMap[key];
      });
    });

    // Cascade A — reassign working employees to better shifts
    for (let ei = employeeList.length - 1; ei >= 0; ei--) {
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'SHIFT') continue;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) continue;
      const employee = employeeList[ei];
      const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
      const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
      const currentHours = getWeeklyHours_(weekGrid, ei);
      const coverageWithout = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, ei);
      const best = selectBestCoverageShift_(employee.qualifiedShifts, employee.status, coverageWithout, shiftTimingMap);
      if (!best) continue;
      const current = weekGrid[ei][dayIndex];
      const currentDef = shiftTimingMap[current.shiftName + '|' + employee.status];
      const currentScore = currentDef ? scoreCoverageForShift_(currentDef, coverageWithout) : 0;
      if (scoreCoverageForShift_(best, coverageWithout) <= currentScore) continue;
      if (currentHours + (best.paidHours - current.paidHours) > effectiveMax) continue;
      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', best.name, best.paidHours, false, best.displayText);
      dayCoverage = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);
      if (!hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot)) break;
    }

    if (!hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot)) return;

    // Cascade B — pull in employees who are off
    for (let ei = employeeList.length - 1; ei >= 0; ei--) {
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'OFF') continue;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) continue;
      if (countWorkingDays_(weekGrid, ei) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) continue;
      const employee = employeeList[ei];
      // Handoff mode: cross-dept employee must not be pulled in on days their home dept has off.
      if (employee.crossDeptScheduledDays && !employee.crossDeptScheduledDays[dayName]) continue;
      const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
      const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
      const currentHours = getWeeklyHours_(weekGrid, ei);
      const best = selectBestCoverageShift_(employee.qualifiedShifts, employee.status, dayCoverage, shiftTimingMap);
      if (!best) continue;
      if (currentHours + best.paidHours > effectiveMax) continue;
      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', best.name, best.paidHours, false, best.displayText);
      dayCoverage = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);
      if (!hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot)) break;
    }
  });
}


// ---------------------------------------------------------------------------
// Phase 4: Role Assignment
// ---------------------------------------------------------------------------

function runPhaseFourRoleAssignment_(weekGrid, employeeList, departmentName) {

  // Parse each employee's full ordered role list once (comma-separated in column L)
  const employeeRoleLists = employeeList.map(function(employee) {
    return (employee.primaryRole || '')
      .split(',')
      .map(function(r) { return r.trim(); })
      .filter(Boolean);
  });

  // Phase 4a: assign everyone their priority (first) role on days they are working
  employeeList.forEach(function(_employee, ei) {
    DAY_NAMES_IN_ORDER.forEach(function(_dayName, dayIndex) {
      const cell = weekGrid[ei][dayIndex];
      cell.role = cell.type === 'SHIFT' ? (employeeRoleLists[ei][0] || null) : null;
    });
  });

  // Phase 4b: fill days where a role has zero coverage.
  // Only pulls a backup if their current primary role stays covered by at least one other
  // person — avoids patching one gap by creating another. Unfilled gaps are left visible
  // so managers can plan around them.
  DAY_NAMES_IN_ORDER.forEach(function(_dayName, dayIndex) {

    // Collect every unique role any employee is qualified for
    const allRoles = [];
    employeeRoleLists.forEach(function(roleList) {
      roleList.forEach(function(role) {
        if (allRoles.indexOf(role) === -1) allRoles.push(role);
      });
    });

    allRoles.forEach(function(role) {
      // Skip roles that already have at least one person assigned today
      const alreadyCovered = employeeList.some(function(_emp, ei) {
        return weekGrid[ei][dayIndex].role === role;
      });
      if (alreadyCovered) return;

      // Find the most senior employee working today who is qualified for this role
      // and whose current role would still have coverage if they were pulled away
      for (let ei = 0; ei < employeeList.length; ei++) {
        const cell = weekGrid[ei][dayIndex];
        if (cell.type !== 'SHIFT') continue;
        if (employeeRoleLists[ei].indexOf(role) === -1) continue; // not qualified

        const currentRole = cell.role;
        const currentRoleStillCovered = employeeList.some(function(_emp, otherEi) {
          return otherEi !== ei && weekGrid[otherEi][dayIndex].role === currentRole;
        });
        if (!currentRole || !currentRoleStillCovered) continue; // would create a new gap

        cell.role = role;
        break;
      }
    });
  });

  // Role ratio rules (e.g., 1 Assist per Cashier)
  if (typeof ROLE_RULES === 'undefined') return; // config.js — optional
  Object.keys(ROLE_RULES).forEach(function(triggerRole) {
    const rule = ROLE_RULES[triggerRole];
    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      let triggerCount = 0, requiredCount = 0;
      employeeList.forEach(function(_emp, ei) {
        const cell = weekGrid[ei][dayIndex];
        if (cell.type !== 'SHIFT') return;
        if (cell.role === triggerRole) triggerCount++;
        if (cell.role === rule.requiresRole) requiredCount++;
      });
      const deficit = (triggerCount * rule.ratio) - requiredCount;
      if (deficit <= 0) return;
      let filled = 0;
      for (let i = employeeList.length - 1; i >= 0 && filled < deficit; i--) {
        const cell = weekGrid[i][dayIndex];
        if (cell.type !== 'SHIFT') continue;
        if (cell.role === triggerRole || cell.role === rule.requiresRole) continue;
        cell.role = rule.requiresRole;
        filled++;
      }
    });
  });
}


// ---------------------------------------------------------------------------
// Helper: Build Shift Display Text with Staggered Start Time
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
  return formatMinutesAsTimeString_(startMinutes) + ' - ' + formatMinutesAsTimeString_(endMinutes);
}

/**
 * Converts minutes-since-midnight to a 12-hour "h:mm AM/PM" string.
 *
 * @param {number} totalMinutes
 * @returns {string}
 */
function formatMinutesAsTimeString_(totalMinutes) {
  const totalHours = Math.floor(totalMinutes / 60);
  const minutes    = totalMinutes % 60;
  const period     = totalHours >= 12 ? 'PM' : 'AM';
  const twelve     = totalHours % 12 === 0 ? 12 : totalHours % 12;
  return twelve + ':' + String(minutes).padStart(2, '0') + ' ' + period;
}

/**
 * Builds a shift display text "HH:MM AM/PM - HH:MM AM/PM" from a start time (string),
 * paid hours, and whether the shift includes an unpaid 30-minute lunch break.
 *
 * @param {string}  startTimeString — e.g. "07:30" (24-hour format)
 * @param {number}  paidHours       — e.g. 8
 * @param {boolean} hasLunch        — when true, adds 30 min to get the clock-out time
 * @returns {string} e.g. "7:30 AM - 4:00 PM"
 */
function buildShiftDisplayText_(startTimeString, paidHours, hasLunch) {
  const startMinutes = timeStringToMinutes_(startTimeString);
  const endMinutes = startMinutes + (paidHours * 60) + (hasLunch ? 30 : 0);
  return formatMinutesAsTimeRange(startMinutes, endMinutes);
}

/**
 * Converts "HH:MM" string to minutes since midnight.
 *
 * @param {string} timeString — e.g. "08:00"
 * @returns {number} e.g. 480
 */
function timeStringToMinutes_(timeString) {
  if (!timeString) return 0;
  const parts = String(timeString).split(':');
  if (parts.length < 2) return 0;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}


// ---------------------------------------------------------------------------
// Phase 5: Pool Member Scheduling (replacing old Supervisor Peak Window Assignment)
// ---------------------------------------------------------------------------

/**
 * Schedules pool members (supervisors, flex workers, etc.) based on traffic level
 * and staggered start times.
 *
 * Pool members were partitioned in generateWeeklySchedule_() and are now assigned
 * to shifts with staggered start times matching the traffic level for each day.
 *
 * Algorithm:
 *   1. Build a map from pool member IDs to their indices in weekGrid (via employeeList)
 *   2. For each day:
 *      a. Determine traffic level and select pool members for that day (by seniority)
 *      b. For each selected pool member:
 *         i. Skip if already VAC/RDO (do not override manager edits)
 *         ii. If OFF, attempt to assign a shift
 *         iii. Select best shift from qualifiedShifts based on coverage need
 *         iv. Pop staggered start time from staggerMap
 *         v. Update weekGrid with staggered shift assignment
 *
 * @param {Array}  weekGrid           — Grid of day assignments (one row per employee)
 * @param {Array}  employeeList       — Full employee list (same order as weekGrid rows)
 * @param {Array}  poolMembers        — Pool members subset
 * @param {Object} heatmapConfig      — { enabled, trafficCurves, pool, ... }
 * @param {Object} dayTrafficLevels   — { Monday: "Low", ... }
 * @param {Object} staggerMap         — { dayName: { shiftKey: [startTimes...], ... }, ... }
 * @param {Object} shiftTimingMap     — Shift definitions
 * @param {Date}   weekStartDate      — The Monday of the week (unused)
 */
function runPhaseFivePoolScheduling_(weekGrid, employeeList, poolMembers, heatmapConfig, dayTrafficLevels, staggerMap, shiftTimingMap, weekStartDate) {
  if (!poolMembers || poolMembers.length === 0) {
    return;
  }

  // Build a map from pool member ID to their index in employeeList/weekGrid
  const poolMemberIndexMap = {};
  poolMembers.forEach(function(poolMember) {
    const poolId = poolMember.id || '';
    // Find this pool member's index in employeeList
    for (let ei = 0; ei < employeeList.length; ei++) {
      if ((employeeList[ei].id || '') === poolId) {
        poolMemberIndexMap[poolId] = ei;
        break;
      }
    }
  });

  // Track stagger position per shift/day to cycle through available start times
  const staggerPositions = {};

  // Process each day
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const trafficLevel = dayTrafficLevels[dayName] || 'Moderate';

    // Select which pool members should work on this day based on traffic level and seniority
    const selectedForDay = selectPoolMembers_(poolMembers, trafficLevel, heatmapConfig); // trafficHeatmapEngine.js

    // Assign each selected pool member to a shift for this day
    selectedForDay.forEach(function(poolMember) {
      const poolId = poolMember.id || '';
      const poolIndex = poolMemberIndexMap[poolId];

      // Skip if pool member's index not found (should not happen if partitioning is correct)
      if (poolIndex === undefined || poolIndex < 0) {
        return;
      }

      const cell = weekGrid[poolIndex][dayIndex];

      // Respect manager edits: do not override VAC or RDO
      if (cell.type === 'VAC' || cell.type === 'RDO') {
        return;
      }

      // Skip if already assigned to a shift (prefer existing assignment)
      if (cell.type === 'SHIFT') {
        return;
      }

      // If OFF, consider assigning a shift
      if (cell.type !== 'OFF') {
        // Some other status (e.g., sick, unpaid leave) — do not override
        return;
      }

      // Check hour constraints before assigning
      const weeklyHours = getWeeklyHours_(weekGrid, poolIndex);
      const weeklyMax = poolMember.status === 'FT' ? HOUR_RULES.FT_MAX : (poolMember.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
      const effectiveMax = weeklyMax - (poolMember.crossDeptHoursAlreadyScheduled || 0);

      // Select best shift from pool member's qualifiedShifts based on coverage
      const dayCoverage = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, poolIndex);
      const bestShift = selectBestCoverageShift_(poolMember.qualifiedShifts || [], poolMember.status, dayCoverage, shiftTimingMap);

      if (!bestShift) {
        // No eligible shift found; skip this pool member for this day
        return;
      }

      // Skip if assigning this shift would exceed hour limits
      if (weeklyHours + bestShift.paidHours > effectiveMax) {
        return;
      }

      // Get staggered start time from staggerMap
      const shiftKey = bestShift.name + '|' + poolMember.status;
      const staggerKey = dayName + '|' + shiftKey;
      const startTimes = staggerMap[dayName] && staggerMap[dayName][shiftKey] ? staggerMap[dayName][shiftKey] : [];

      // Default display text uses the day-specific anchor (sat/sun overrides applied here)
      const dayAnchorMinutes = getStartMinutesForDay_(bestShift, dayName); // settingsManager.js
      let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + bestShift.blockMinutes);

      if (startTimes && startTimes.length > 0) {
        // Stagger map overrides the anchor when flex is enabled
        if (!staggerPositions[staggerKey]) staggerPositions[staggerKey] = 0;
        const staggerStartTime = startTimes[staggerPositions[staggerKey] % startTimes.length];
        staggerPositions[staggerKey]++;
        displayText = buildShiftDisplayText_(staggerStartTime, bestShift.paidHours, bestShift.hasLunch);
      }

      // Assign pool member to shift
      weekGrid[poolIndex][dayIndex] = createDayAssignment_('SHIFT', bestShift.name, bestShift.paidHours, false, displayText);
    });
  });

  // Log completion
  console.log('Phase 5: Pool Scheduling - assigned ' + Object.keys(poolMemberIndexMap).length + ' pool members with traffic-aware staggering');
}


// ---------------------------------------------------------------------------
// Coverage Map Functions
// ---------------------------------------------------------------------------

function buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, excludeIndex) {
  const slots = new Array(COVERAGE.SLOT_COUNT).fill(0); // config.js
  employeeList.forEach(function(employee, ei) {
    if (ei === excludeIndex) return;
    const cell = weekGrid[ei][dayIndex];
    if (cell.type !== 'SHIFT') return;
    const shiftDef = shiftTimingMap[cell.shiftName + '|' + employee.status];
    if (!shiftDef) return;
    const startSlot = Math.max(0, Math.floor((shiftDef.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
    const endSlot   = Math.min(COVERAGE.SLOT_COUNT, Math.floor((shiftDef.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
    for (let s = startSlot; s < endSlot; s++) slots[s]++;
  });
  return slots;
}

function hasCoverageGaps_(slots, startSlot, endSlot) {
  for (let s = startSlot; s < endSlot; s++) {
    if (slots[s] === 0) return true;
  }
  return false;
}

function selectBestCoverageShift_(qualifiedShiftNames, status, coverageSlots, shiftTimingMap) {
  let highScore = 0, best = null;
  qualifiedShiftNames.forEach(function(name) {
    const def = shiftTimingMap[name.trim() + '|' + status];
    if (!def) return;
    const score = scoreCoverageForShift_(def, coverageSlots);
    if (score > highScore) { highScore = score; best = def; }
  });
  return best;
}

function scoreCoverageForShift_(shiftDef, coverageSlots) {
  const startSlot = Math.max(0, Math.floor((shiftDef.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
  const endSlot   = Math.min(COVERAGE.SLOT_COUNT, Math.floor((shiftDef.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
  let score = 0;
  for (let s = startSlot; s < endSlot; s++) score += 1 / (coverageSlots[s] + 1);
  return score;
}


// ---------------------------------------------------------------------------
// Serialization Helpers (for google.script.run return values)
// ---------------------------------------------------------------------------

/**
 * Strips Date objects and non-primitive values from a weekGrid so it can be
 * safely returned over google.script.run.
 *
 * @param {Array} weekGrid
 * @param {Array} employeeList
 * @returns {Array}
 */
function serializeWeekGrid_(weekGrid, employeeList) {
  return weekGrid.map(function(employeeRow, ei) {
    return employeeRow.map(function(cell) {
      return {
        type:        cell.type,
        shiftName:   cell.shiftName   || null,
        paidHours:   cell.paidHours   || 0,
        locked:      cell.locked      || false,
        displayText: cell.displayText || null,
        role:        cell.role        || null,
      };
    });
  });
}

/**
 * Strips non-serializable fields from an employee list for return to the frontend.
 * Also calculates and includes weekly hours and under-hours status.
 *
 * @param {Array} employeeList
 * @param {Array} weekGrid — Week schedule grid (used to calculate weeklyHours)
 * @returns {Array}
 */
function serializeEmployeeList_(employeeList, weekGrid) {
  return employeeList.map(function(emp, i) {
    var weeklyHours = getWeeklyHours_(weekGrid, i);
    var minHours = emp.status === 'FT' ? HOUR_RULES.FT_MIN : (emp.status === 'LPT' ? HOUR_RULES.LPT_MIN : HOUR_RULES.PT_MIN);
    return {
      name:                 emp.name,
      employeeId:           emp.employeeId,
      status:               emp.status,
      department:           emp.department,
      seniorityRank:        emp.seniorityRank,
      primaryRole:          emp.primaryRole || '',
      weeklyHours:          weeklyHours,
      underHours:           weeklyHours < minHours,
      secondaryDepartments: emp.secondaryDepartments || [],
    };
  });
}


// ---------------------------------------------------------------------------
// Utility Functions
// ---------------------------------------------------------------------------

function getWeeklyHours_(weekGrid, ei) {
  let total = 0;
  weekGrid[ei].forEach(function(cell) { if (cell.type === 'SHIFT') total += cell.paidHours; });
  return total;
}

/** Public alias — called by formatter.js without underscore. */
function getWeeklyHours(weekGrid, employeeIndex) {
  return getWeeklyHours_(weekGrid, employeeIndex);
}

function countWorkingDays_(weekGrid, ei) {
  let count = 0;
  for (let di = 0; di < WEEK_SHEET.DAYS_IN_WEEK; di++) {
    if (weekGrid[ei][di].type === 'SHIFT') count++;
  }
  return count;
}

function compareEmployeesBySeniority_(a, b) {
  if (b.seniorityRank !== a.seniorityRank) return b.seniorityRank - a.seniorityRank;
  // FT first, then PT, then LPT
  if (a.status !== b.status) return a.status === 'FT' ? -1 : (b.status === 'FT' ? 1 : (a.status === 'LPT' ? 1 : -1));
  return a.name.localeCompare(b.name);
}

/**
 * Sort order for the schedule grid:
 *   Tier 1 — Supervisors before everyone else (seniority within tier)
 *   Tier 2 — Remaining employees sorted alphabetically by primary role
 *   Tier 3 — Seniority within each role group
 *
 * @param {object} a
 * @param {object} b
 * @returns {number}
 */
function compareEmployeesForScheduleOrder_(a, b) {
  var aIsSupervisor = isSupervisorRole_(a.primaryRole);
  var bIsSupervisor = isSupervisorRole_(b.primaryRole);
  if (aIsSupervisor !== bIsSupervisor) return aIsSupervisor ? -1 : 1;

  if (!aIsSupervisor) {
    var aRole = (a.primaryRole || '').split(',')[0].trim().toLowerCase();
    var bRole = (b.primaryRole || '').split(',')[0].trim().toLowerCase();
    if (aRole !== bRole) return aRole.localeCompare(bRole);
  }

  return compareEmployeesBySeniority_(a, b);
}

/**
 * Returns true if the employee's primary role is "supervisor" (case-insensitive).
 *
 * @param {string} roleString — Comma-separated role list; first entry is the primary role.
 * @returns {boolean}
 */
function isSupervisorRole_(roleString) {
  var firstRole = (roleString || '').split(',')[0].trim().toLowerCase();
  return firstRole === 'supervisor';
}

function createDayAssignment_(type, shiftName, paidHours, locked, displayText) {
  return { type, shiftName: shiftName || null, paidHours: paidHours || 0, locked: locked || false, displayText: displayText || null, role: null };
}

function resolveShiftForEmployee_(employee, shiftTimingMap) {
  const prefKey = employee.preferredShift + '|' + employee.status;
  if (shiftTimingMap[prefKey]) return shiftTimingMap[prefKey];
  for (let i = 0; i < employee.qualifiedShifts.length; i++) {
    const key = employee.qualifiedShifts[i] + '|' + employee.status;
    if (shiftTimingMap[key]) return shiftTimingMap[key];
  }
  return null;
}

/**
 * Returns a new Date offset by dayIndex days from weekStartDate.
 * Used by formatter.js to build the week-range label in the sheet header.
 *
 * @param {Date}   weekStartDate — The Monday of the week.
 * @param {number} dayIndex      — 0 = Monday … 6 = Sunday.
 * @returns {Date}
 */
function getDateForDayIndex(weekStartDate, dayIndex) {
  const result = new Date(weekStartDate);
  result.setDate(weekStartDate.getDate() + dayIndex);
  return result;
}

function generateWeekSheetName_(weekStartDate, deptName) {
  const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
  const day   = String(weekStartDate.getDate()).padStart(2, '0');
  const year  = String(weekStartDate.getFullYear()).slice(-2);
  const base  = 'Week_' + month + '_' + day + '_' + year;
  return deptName ? base + '_' + deptName : base;
}

function normalizeDeptName_(name) {
  if (!name) return '';
  return name.toString().trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseVacationDateString_(dateString, weekStartDate) {
  const s = dateString.toString().trim();
  if (!s) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const d = new Date(s + 'T00:00:00');
    if (!isNaN(d.getTime())) return d;
  }
  if (/^\d{1,2}\/\d{1,2}$/.test(s)) {
    const parts = s.split('/');
    const d = new Date(weekStartDate.getFullYear(), parseInt(parts[0], 10) - 1, parseInt(parts[1], 10));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function parseVacationDateStrings_(raw) {
  if (!raw || raw.toString().trim() === '') return [];
  return raw.toString().split(',').map(s => s.trim()).filter(Boolean);
}

function parseQualifiedShiftList_(raw, preferred) {
  if (!raw || raw.toString().trim() === '') {
    return preferred ? [preferred.toString().trim()] : [];
  }
  const list = raw.toString().split(',').map(s => s.trim()).filter(Boolean);
  if (preferred && !list.includes(preferred.toString().trim())) {
    list.unshift(preferred.toString().trim());
  }
  return list;
}

function getDayIndexForDate_(targetDate, weekStartDate) {
  const weekStart = new Date(weekStartDate); weekStart.setHours(0,0,0,0);
  const target    = new Date(targetDate);    target.setHours(0,0,0,0);
  const diff = Math.round((target.getTime() - weekStart.getTime()) / 86400000);
  return (diff < 0 || diff > 6) ? -1 : diff;
}

function readCheckboxStateFromSheet_(weekSheet, employeeCount) {
  const state = [];
  for (let ei = 0; ei < employeeCount; ei++) {
    state[ei] = [];
    const baseRow = WEEK_SHEET.DATA_START_ROW + (ei * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    // Read VAC, RDO, SHIFT, and LOCK rows to fully reconstruct grid state.
    const vacRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const rdoRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const shiftRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const lockRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];

    for (let di = 0; di < WEEK_SHEET.DAYS_IN_WEEK; di++) {
      const isLocked = lockRow[di] === true;
      if (vacRow[di] === true) {
        state[ei][di] = createDayAssignment_('VAC', null, 0, true);
      } else if (rdoRow[di] === true) {
        state[ei][di] = createDayAssignment_('RDO', null, 0, false);
      } else if (shiftRow[di] && shiftRow[di] !== 'OFF' && shiftRow[di] !== '') {
        // This is a SHIFT cell — reconstruct it from the display text.
        // shiftName is null; displayText is the cell content.
        state[ei][di] = createDayAssignment_('SHIFT', null, 0, isLocked, String(shiftRow[di]));
      } else {
        state[ei][di] = createDayAssignment_('OFF', null, 0, isLocked);
      }
    }
  }
  return state;
}
