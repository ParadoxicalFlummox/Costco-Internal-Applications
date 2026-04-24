/**
 * scheduleEngine.js — Core schedule generation algorithm for COMET.
 * VERSION: 0.4.1
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
 *
 * THE FIVE PHASES:
 *   Phase 0 — Bootstrap: Load roster for dept, initialize grid, load cross-dept hours, stamp vacation locks.
 *   Phase 1 — Preference Assignment: Honor day-off prefs and shift prefs, seniority order (regular employees only).
 *   Phase 2 — Minimum Hour Enforcement: Add shifts until weekly minimum is met (regular employees only).
 *   Phase 3 — Gap Resolution: Fill uncovered time slots by reassigning or adding employees (regular employees only).
 *   Phase 4 — Role Assignment: Stamp primaryRole onto SHIFT cells (regular employees only).
 *   Phase 5 — Supervisor Assignment: Assign supervisors to peak traffic windows based on supervisor peak config.
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

  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate);
  if (employeeList.length === 0) {
    throw new Error('No active employees found for department "' + deptName + '".');
  }

  const weekGrid = initializeWeekGrid_(employeeList, weekStartDate);

  // PERF: Wrap each phase with execution time logging
  logExecutionTime_('Phase 1: Preference Assignment (' + employeeList.length + ' employees)', function() {
    runPhaseOnePreferenceAssignment_(weekGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate);
  });

  logExecutionTime_('Phase 2: Minimum Hour Enforcement', function() {
    runPhaseTwoHourEnforcement_(weekGrid, employeeList, shiftTimingMap);
  });

  logExecutionTime_('Phase 3: Gap Resolution', function() {
    runPhaseThreeGapResolution_(weekGrid, employeeList, shiftTimingMap, staffingRequirements);
  });

  logExecutionTime_('Phase 4: Role Assignment', function() {
    runPhaseFourRoleAssignment_(weekGrid, employeeList, deptName);
  });

  logExecutionTime_('Phase 5: Supervisor Assignment', function() {
    runPhaseFiveSupervisorAssignment_(weekGrid, employeeList, deptName, weekStartDate);
  });

  return {
    weekSheetName: generateWeekSheetName_(weekStartDate, deptName),
    weekGrid:      weekGrid,
    employeeList:  employeeList,
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

      // Load secondary departments from Employees sheet column N
      let secondaryDepartments = [];
      let crossDeptHoursAlreadyScheduled = 0;
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

        // If the employee has secondary departments, query for cross-dept hours
        if (secondaryDepartments.length > 0 && weekStartDate) {
          crossDeptHoursAlreadyScheduled = getCrossDeptHoursForWeek_(emp, weekStartDate, deptName);
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
        seniorityRank:                   Number(emp.seniorityRank || 0),
        department:                      normalizeDeptName_(emp.department),
        primaryRole:                     emp.role || '',
        secondaryDepartments:            secondaryDepartments,
        crossDeptHoursAlreadyScheduled:  crossDeptHoursAlreadyScheduled,
      };
    });

  deptEmployees.sort(compareEmployeesBySeniority_);
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
  const sheetNames = workbook.getSheetNames();
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


// ---------------------------------------------------------------------------
// Phase 1: Preference Assignment
// ---------------------------------------------------------------------------

function runPhaseOnePreferenceAssignment_(weekGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate) {
  grantRequestedDaysOff_(weekGrid, employeeList, staffingRequirements);
  assignPreferredShifts_(weekGrid, employeeList, shiftTimingMap);
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

function assignPreferredShifts_(weekGrid, employeeList, shiftTimingMap) {
  employeeList.forEach(function(employee, ei) {
    DAY_NAMES_IN_ORDER.forEach(function(_dayName, dayIndex) {
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'OFF') return;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) return;
      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (!shiftDef) return;
      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, shiftDef.displayText);
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

function runPhaseTwoHourEnforcement_(weekGrid, employeeList, shiftTimingMap) {
  employeeList.forEach(function(employee, ei) {
    const weeklyMin = employee.status === 'FT' ? HOUR_RULES.FT_MIN : HOUR_RULES.PT_MIN; // config.js
    const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;
    // Reduce effective max by hours already scheduled in other departments (split-shift)
    const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
    let currentHours = getWeeklyHours_(weekGrid, ei);
    if (currentHours >= weeklyMin) return;
    const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
    if (!shiftDef) return;
    DAY_NAMES_IN_ORDER.forEach(function(_dayName, dayIndex) {
      if (currentHours >= weeklyMin) return;
      const cell = weekGrid[ei][dayIndex];
      if (cell.type !== 'OFF') return;
      // Skip if this cell is locked (manager override via updateCellOverride).
      if (cell.locked) return;
      if (countWorkingDays_(weekGrid, ei) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) return;
      if (currentHours + shiftDef.paidHours > effectiveMax) return;
      weekGrid[ei][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, shiftDef.displayText);
      currentHours += shiftDef.paidHours;
    });
  });
}


// ---------------------------------------------------------------------------
// Phase 3: Gap Resolution
// ---------------------------------------------------------------------------

function runPhaseThreeGapResolution_(weekGrid, employeeList, shiftTimingMap, staffingRequirements) {
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
      const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;
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
      const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;
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
  employeeList.forEach(function(employee, ei) {
    DAY_NAMES_IN_ORDER.forEach(function(_dayName, dayIndex) {
      const cell = weekGrid[ei][dayIndex];
      cell.role = cell.type === 'SHIFT' ? (employee.primaryRole || null) : null;
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
// Phase 5: Supervisor Peak Window Assignment
// ---------------------------------------------------------------------------

/**
 * Assigns supervisors to peak traffic windows based on the configured peak profile.
 *
 * Supervisors (role = SUPERVISOR_ROLE) are excluded from Phases 1-4 and only assigned
 * during Phase 5 based on hourly traffic intensity. Peak windows are computed from the
 * peakProfile (24 hourly values) using peakThreshold. Supervisors are selected by
 * seniority for each peak window that requires coverage.
 *
 * @param {Array}  weekGrid        — Grid of day assignments (one row per employee)
 * @param {Array}  employeeList    — List of all employees (supervisors + regular)
 * @param {string} departmentName  — Department name (for loading peak config)
 * @param {Date}   weekStartDate   — The Monday of the week
 */
function runPhaseFiveSupervisorAssignment_(weekGrid, employeeList, departmentName, weekStartDate) {
  // Load supervisor peak config (from COMET Config sheet, or default from config.js)
  let supervisorConfig = readSupervisorPeakConfig_(departmentName);
  if (!supervisorConfig) {
    // No config found; use default
    supervisorConfig = {
      department: departmentName,
      peakProfile: SUPERVISOR_RULES.defaultPeakProfile,
      enabled: SUPERVISOR_RULES.enabled || false,
      membersPerSupervisor: SUPERVISOR_RULES.membersPerSupervisor || 75,
      maxDoorCount: SUPERVISOR_RULES.maxDoorCount || 500,
    };
  }

  // Check if supervisor scheduling is enabled for this department
  if (!supervisorConfig.enabled) {
    return;
  }

  // Extract list of supervisors from employeeList
  const supervisors = [];
  const supervisorIndices = [];
  employeeList.forEach(function(employee, index) {
    if (employee.primaryRole === SUPERVISOR_RULES.supervisorRole) {
      supervisors.push(employee);
      supervisorIndices.push(index);
    }
  });

  if (supervisors.length === 0) {
    return; // No supervisors to assign
  }

  const membersPerSupervisor = supervisorConfig.membersPerSupervisor || 75;

  // For each day of the week, assign supervisors based on door count
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const peakProfile = supervisorConfig.peakProfile[dayName] || [];

    // peakProfile is an 8-element array (one per 3 hours: 0, 3, 6, 9, 12, 15, 18, 21)
    for (var pointIndex = 0; pointIndex < peakProfile.length; pointIndex++) {
      var doorCount = peakProfile[pointIndex] || 0;

      // Calculate how many supervisors are needed at this time slot
      var supervisorsNeeded = Math.ceil(doorCount / membersPerSupervisor);

      if (supervisorsNeeded === 0) {
        continue; // No supervisors needed for this time slot
      }

      // Try to assign supervisorsNeeded supervisors to this time slot
      var assignedCount = 0;
      var attemptedIndices = new Set();

      while (assignedCount < supervisorsNeeded && attemptedIndices.size < supervisors.length) {
        // Find highest-seniority supervisor available (not VAC/RDO, not already assigned)
        let selectedSupervisorIndex = -1;
        let highestSeniority = -1;

        supervisors.forEach(function(supervisor, supervisorIndex) {
          if (attemptedIndices.has(supervisorIndex)) {
            return; // Already tried this supervisor
          }

          const supervisorEmployeeIndex = supervisorIndices[supervisorIndex];
          const cell = weekGrid[supervisorEmployeeIndex][dayIndex];

          // Check if supervisor is VAC or RDO
          if (cell.type === 'VAC' || cell.type === 'RDO') {
            return; // Skip this supervisor
          }

          // Check if supervisor is available (currently OFF)
          if (cell.type !== 'OFF') {
            return; // Skip (already assigned to something else)
          }

          // Select supervisor with highest seniority
          if (supervisor.seniorityRank > highestSeniority) {
            highestSeniority = supervisor.seniorityRank;
            selectedSupervisorIndex = supervisorIndex;
          }
        });

        if (selectedSupervisorIndex === -1) {
          // No more supervisors available for this time slot
          break;
        }

        // Assign the selected supervisor
        const supervisorEmployeeIndex = supervisorIndices[selectedSupervisorIndex];
        const timeLabels = ['12am', '3am', '6am', '9am', '12pm', '3pm', '6pm', '9pm'];
        const displayText = 'Peak (' + timeLabels[pointIndex] + ')';

        weekGrid[supervisorEmployeeIndex][dayIndex] = createDayAssignment_(
          'SHIFT',
          'Peak',
          3,  // Rough estimate: 3-hour block
          false,
          displayText
        );

        assignedCount++;
        attemptedIndices.add(selectedSupervisorIndex);
      }

      if (assignedCount < supervisorsNeeded) {
        logConsole_(
          'WARNING: Day ' + dayName + ' at ' + ['12am', '3am', '6am', '9am', '12pm', '3pm', '6pm', '9pm'][pointIndex] +
          ' (' + doorCount + ' members): needed ' + supervisorsNeeded + ' supervisors, but only ' + assignedCount + ' available'
        );
      }
    }
  });
}

/**
 * Computes peak windows from a 24-hour traffic profile.
 *
 * Scans through hourly values and identifies contiguous blocks where traffic >= peakThreshold.
 * Returns array of [startHour, endHour) pairs.
 *
 * @param {Array}  peakProfile — 24-element array of hourly traffic intensities (0-1)
 * @param {number} peakThreshold — Traffic level above which a window is "peak" (default 0.7)
 * @returns {Array<[number, number]>} Array of [startHour, endHour] pairs
 */
function computePeakWindows_(peakProfile, peakThreshold) {
  if (!peakProfile || peakProfile.length !== 24) {
    return [];
  }

  const windows = [];
  let currentWindowStart = -1;

  for (let hour = 0; hour < 24; hour++) {
    const intensity = peakProfile[hour];
    const isPeak = intensity >= (peakThreshold || 0.7);

    if (isPeak && currentWindowStart === -1) {
      // Start of a new peak window
      currentWindowStart = hour;
    } else if (!isPeak && currentWindowStart !== -1) {
      // End of current peak window
      windows.push([currentWindowStart, hour]);
      currentWindowStart = -1;
    }
  }

  // Handle window that extends to end of day
  if (currentWindowStart !== -1) {
    windows.push([currentWindowStart, 24]);
  }

  return windows;
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
    var minHours = emp.status === 'FT' ? HOUR_RULES.FT_MIN : HOUR_RULES.PT_MIN;
    return {
      name:           emp.name,
      employeeId:     emp.employeeId,
      status:         emp.status,
      department:     emp.department,
      seniorityRank:  emp.seniorityRank,
      primaryRole:    emp.primaryRole || '',
      weeklyHours:    weeklyHours,
      underHours:     weeklyHours < minHours,
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
  if (a.status !== b.status) return a.status === 'FT' ? -1 : 1;
  return a.name.localeCompare(b.name);
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
