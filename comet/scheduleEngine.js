/**
 * scheduleEngine.js — Core schedule generation algorithm for COMET.
 * VERSION: 0.7.2
 *
 * REWORK SUMMARY (v0.7.0):
 *   - Phase condensation: 5 → 4 phases (Phases 1+2 merged; role assignment moved last).
 *   - Phase 0: Batch preload of ALL week sheets in one call; per-employee sheet reads removed.
 *     Cross-dept hours now derived from preloaded map (O(depts) vs O(n × depts) previously).
 *   - Phase 0: Combo participant loading — employees whose secondary dept = current dept are
 *     now included in the roster and scheduled with their home dept days as the constraint.
 *   - Phase 1 (merged 1+2): Single pass for preference assignment + hour enforcement.
 *     VAC/RDO fix: days off counted as OFF|RDO only; VAC counts as a working day.
 *     Pre-computed day coverage eliminates O(n²) inner scan in enforceMinimumDaysOff_.
 *     Shared staggerPositions initialized once and passed through all phases.
 *   - Phase 2 (gap resolution): Full rewrite — single-pass priority gap fill.
 *     buildDayCoverage_ called once per day, updated incrementally.
 *     Hard iteration cap (max(50, regularCount×2)) prevents O(n³) cascade.
 *     Combo participants excluded from candidate pool but counted in coverage.
 *   - Phase 3 (pool): Uses shared staggerPositions; cross-dept guard uses preloaded map.
 *   - Phase 4 (role): Runs last on ALL employees (regular + combo + pool).
 *     Role minimums enforcement replaces "cover if zero" when enforceRoleMinimums is on.
 *     Role minimums scale with traffic level (Low/Moderate/High).
 *   - readCheckboxStateFromSheet_: RDO cells now read LOCK row (manager RDO overrides preserved).
 *   - appendHybridEmployees_: New targeted operation to add combo participants post-generation.
 *
 * THE FOUR PHASES:
 *   Phase 0 — Bootstrap: Batch preload sheets; load roster (primary + combo participants);
 *             init grid; stamp vacation locks; init shared staggerPositions.
 *   Phase 1 — Preference & Hour Assignment: Honor day-off prefs, shift prefs, and hour
 *             minimums in one pass (regular employees); combo participant mirroring sub-loop.
 *   Phase 2 — Gap Resolution: Fill uncovered slots via priority gap fill (toggleable).
 *   Phase 3 — Pool Scheduling: Assign pool members traffic-aware with shared stagger.
 *   Phase 4 — Role Assignment (all employees): Stamp roles; enforce role minimums.
 */


// ---------------------------------------------------------------------------
// Top-level Entry Point
// ---------------------------------------------------------------------------

/**
 * Generates a complete weekly schedule for the given department and Monday.
 *
 * @param {string} deptName       — Department name (must match employees' department field).
 * @param {Date}   weekStartDate  — The Monday of the week to generate.
 * @param {Object} engineOptions  — { enforceRoleMinimums: boolean, gapFillEnabled: boolean }
 * @param {Object} roleMinimums   — { RoleName: { Low: number, Moderate: number, High: number }, ... }
 * @returns {{ weekSheetName: string, weekGrid: Array, employeeList: Array, poolMemberIds: Set, comboParticipantIds: Set }}
 */
function generateWeeklySchedule_(deptName, weekStartDate, engineOptions, roleMinimums) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const shiftTimingMap = buildShiftTimingMap(deptName);      // settingsManager.js
  const staffingRequirements = loadStaffingRequirements(deptName); // settingsManager.js

  // --- PHASE 0: BOOTSTRAP ---
  // Batch-read ALL week sheets in a single pass (eliminates O(n × depts) sheet reads).
  const preloadedSheets = preloadWeekSheets_(workbook, weekStartDate);

  // --- TRAFFIC HEATMAP INTEGRATION ---
  const heatmapConfig = loadTrafficHeatmapConfig_(deptName); // settingsManager.js
  let dayTrafficLevels = {};
  let staggerMap = {};

  if (heatmapConfig && heatmapConfig.enabled) {
    dayTrafficLevels = classifyDayTrafficLevels_(heatmapConfig, weekStartDate); // trafficHeatmapEngine.js
    staggerMap = buildStaggeredStartMap_(shiftTimingMap, dayTrafficLevels, weekStartDate); // trafficHeatmapEngine.js
    console.log('Traffic heatmap enabled: ' + Object.keys(dayTrafficLevels).map(function (d) { return d + '=' + dayTrafficLevels[d]; }).join(', '));
    console.log('Stagger map computed for ' + Object.keys(staggerMap).length + ' days');
  } else {
    // Heatmap disabled: default to all Moderate, no stagger
    DAY_NAMES_IN_ORDER.forEach(function (dayName) {
      dayTrafficLevels[dayName] = 'Moderate';
    });
    console.log('Traffic heatmap disabled; using Moderate minimums');
  }

  // Load all employees for this dept, including combo participants from other depts.
  // preloadedSheets is passed so loadRosterSortedBySeniority_ can resolve homeScheduledDays
  // without additional sheet reads.
  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate, preloadedSheets);
  if (employeeList.length === 0) {
    throw new Error('No active employees found for department "' + deptName + '".');
  }

  // Split roster into primary employees and combo participants.
  // Pool partitioning only applies to primary employees.
  const primaryEmployees = employeeList.filter(function (employee) { return !employee.isComboParticipant; });
  const comboParticipants = employeeList.filter(function (employee) { return employee.isComboParticipant; });

  let poolMembers = [];
  let regularEmployees = [];

  if (heatmapConfig && heatmapConfig.enabled) {
    const partitioned = partitionPoolMembers_(primaryEmployees, heatmapConfig); // trafficHeatmapEngine.js
    poolMembers = partitioned.poolMembers;
    regularEmployees = partitioned.regularEmployees;
    console.log('Pool partitioned: ' + poolMembers.length + ' pool, ' + regularEmployees.length + ' regular, ' + comboParticipants.length + ' combo');
  } else {
    regularEmployees = primaryEmployees;
    console.log('Roster: ' + regularEmployees.length + ' regular, ' + comboParticipants.length + ' combo');
  }

  // Shared stagger position tracker — initialized ONCE and passed through all phases
  // so no two employees across any phase receive the same stagger slot.
  const staggerPositions = {};

  const weekGrid = initializeWeekGrid_(employeeList, weekStartDate);

  // --- PHASE 1: PREFERENCE & HOUR ASSIGNMENT ---
  // Handles day-off preferences, shift preferences, and minimum hour enforcement in one pass.
  // Includes a sub-loop for combo participants (mirrors home dept scheduled days).
  logExecutionTime_('Phase 1: Preference & Hour Assignment (' + regularEmployees.length + ' regular, ' + comboParticipants.length + ' combo)', function () {
    runPhaseOnePreferenceAndHours_(weekGrid, regularEmployees, comboParticipants, shiftTimingMap,
      staffingRequirements, weekStartDate, staggerMap, staggerPositions, dayTrafficLevels);
  });

  // --- PHASE 2: GAP RESOLUTION ---
  // Single-pass priority fill with hard iteration cap (toggleable).
  // Combo participants are excluded from candidate pool but counted in coverage.
  logExecutionTime_('Phase 2: Gap Resolution', function () {
    runPhaseGapResolution_(weekGrid, regularEmployees, employeeList, shiftTimingMap,
      staffingRequirements, staggerMap, staggerPositions, dayTrafficLevels, engineOptions);
  });

  // --- PHASE 3: POOL SCHEDULING ---
  logExecutionTime_('Phase 3: Pool Scheduling', function () {
    if (poolMembers.length > 0) {
      runPhasePoolScheduling_(weekGrid, employeeList, poolMembers, heatmapConfig,
        dayTrafficLevels, staggerMap, staggerPositions, shiftTimingMap,
        weekStartDate, preloadedSheets);
    }
  });

  // --- PHASE 4: ROLE ASSIGNMENT (ALL EMPLOYEES) ---
  // Runs last so pool members and combo participants also receive role stamps.
  // Enforces role minimums per traffic level when engineOptions.enforceRoleMinimums is true.
  logExecutionTime_('Phase 4: Role Assignment (all ' + employeeList.length + ' employees)', function () {
    runPhaseRoleAssignment_(weekGrid, employeeList, deptName, dayTrafficLevels, roleMinimums, engineOptions);
  });

  // Build ID sets for the formatter to visually distinguish pool / combo sections.
  const poolMemberIdSet = new Set();
  poolMembers.forEach(function (employee) { if (employee.id) poolMemberIdSet.add(employee.id); });

  const comboParticipantIdSet = new Set();
  comboParticipants.forEach(function (employee) { if (employee.id) comboParticipantIdSet.add(employee.id); });

  return {
    weekSheetName: generateWeekSheetName_(weekStartDate, deptName),
    weekGrid: weekGrid,
    employeeList: employeeList,
    poolMemberIds: poolMemberIdSet,
    comboParticipantIds: comboParticipantIdSet,
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

  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate, null);
  if (employeeList.length === 0) return null;

  // Read from JSON format (new) or legacy format (if this sheet was generated before the refactor)
  let weekGrid;
  const format = detectSheetFormat(sheet); // formatter.js
  if (format === 'json') {
    weekGrid = readJsonScheduleFromSheet_(sheet, employeeList);
  } else {
    // Legacy format: use the old 5-row reader for backward compat
    console.warn('readExistingWeekSchedule_: sheet "' + sheetName + '" uses legacy 5-row format. Re-generate to upgrade to JSON format.');
    weekGrid = readCheckboxStateFromSheet_(sheet, employeeList.length);
  }

  return { weekSheetName: sheetName, weekGrid, employeeList };
}


/**
 * Legacy read function for backward compatibility. Replaced by readJsonScheduleFromSheet_
 * but kept temporarily for sheets generated in the old 5-row format.
 */
function readCheckboxStateFromSheet_(weekSheet, employeeCount) {
  const state = [];
  for (let ei = 0; ei < employeeCount; ei++) {
    state[ei] = [];
    const baseRow = WEEK_SHEET.DATA_START_ROW + (ei * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const vacRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_VAC, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const rdoRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_RDO, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const shiftRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];
    const lockRow = weekSheet.getRange(baseRow + WEEK_SHEET.ROW_OFFSET_LOCK, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).getValues()[0];

    for (let di = 0; di < WEEK_SHEET.DAYS_IN_WEEK; di++) {
      const isLocked = lockRow[di] === true;
      if (vacRow[di] === true) {
        state[ei][di] = createDayAssignment_('VAC', null, 0, true);
      } else if (rdoRow[di] === true) {
        state[ei][di] = createDayAssignment_('RDO', null, 0, isLocked);
      } else if (shiftRow[di] && shiftRow[di] !== 'OFF' && shiftRow[di] !== '') {
        state[ei][di] = createDayAssignment_('SHIFT', null, 0, isLocked, String(shiftRow[di]));
      } else {
        state[ei][di] = createDayAssignment_('OFF', null, 0, isLocked);
      }
    }
  }
  return state;
}


// ---------------------------------------------------------------------------
// Phase 0: Batch Sheet Preload + Roster Loading
// ---------------------------------------------------------------------------

/**
 * Reads every Week sheet for the given Monday into memory in a single pass.
 * Returns a plain object keyed by sheet name, value is the 2D value array.
 *
 * All subsequent cross-dept lookups read from this map — zero additional
 * sheet reads are needed after this call returns.
 *
 * @param {Spreadsheet} workbook
 * @param {Date}        weekStartDate — The Monday of the week
 * @returns {Object} { sheetName: 2DValueArray, ... }
 */
function preloadWeekSheets_(workbook, weekStartDate) {
  const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
  const day = String(weekStartDate.getDate()).padStart(2, '0');
  const year = String(weekStartDate.getFullYear()).slice(-2);
  const weekBaseName = 'Week_' + month + '_' + day + '_' + year + '_';

  const preloadedMap = {};
  workbook.getSheets().forEach(function (sheet) {
    const sheetName = sheet.getName();
    if (sheetName.startsWith(weekBaseName)) {
      preloadedMap[sheetName] = sheet.getDataRange().getValues();
    }
  });

  console.log('preloadWeekSheets_: loaded ' + Object.keys(preloadedMap).length + ' week sheet(s) into memory');
  return preloadedMap;
}

/**
 * Preload all week sheets from the workbook.
 * Used when appending hybrid staff to support cross-week assignments.
 */
function preloadAllWeekSheets_(workbook) {
  const preloadedMap = {}
  workbook.getSheets().forEach(function (sheet) {
    const sheetName = sheet.getName();
    if (sheetName.startsWith('Week_')) {
      preloadedMap[sheetName] = sheet.getDataRange().getValues();
    }
  });
  console.log('preloadAllWeekSheets_: loaded ' + Object.keys(preloadedMap).length + ' week sheet(s) into memory');
  return preloadedMap;
}


/**
 * Extracts home-dept scheduled days for a combo participant from preloaded 2D data.
 * Returns { Monday: bool, ... } keyed by day name, or null if not found or invalid JSON.
 *
 * @param {string} employeeName      — Employee's name
 * @param {string} homeDeptSheetName — The home dept sheet name
 * @param {Object} preloadedSheets   — { sheetName: 2DValueArray }
 * @returns {{ Monday: boolean, ... }|null}
 */
function extractHomeDeptScheduledDays_(employeeName, homeDeptSheetName, preloadedSheets) {
  const data = preloadedSheets[homeDeptSheetName];
  if (!data || data.length < WEEK_SHEET.DATA_START_ROW) return null;

  const scheduledDays = {
    Monday: false, Tuesday: false, Wednesday: false,
    Thursday: false, Friday: false, Saturday: false, Sunday: false,
  };

  // JSON format: row[COL_NAME - 1] = name, row[COL_SCHEDULE_JSON - 1] = JSON string
  for (let rowIndex = WEEK_SHEET.DATA_START_ROW - 1; rowIndex < data.length; rowIndex++) {
    const row = data[rowIndex];
    const cellEmpName = row[WEEK_SHEET.COL_NAME - 1];
    if (!cellEmpName || cellEmpName.toString().trim() !== employeeName.toString().trim()) continue;

    try {
      const scheduleJson = row[WEEK_SHEET.COL_SCHEDULE_JSON - 1];
      if (!scheduleJson) return null;

      const scheduleObj = JSON.parse(scheduleJson);
      DAY_NAMES_IN_ORDER.forEach(function (dayName) {
        const dayData = scheduleObj[dayName];
        if (dayData && dayData.type === 'SHIFT') {
          scheduledDays[dayName] = true;
        }
      });
      return scheduledDays;
    } catch (e) {
      console.warn('extractHomeDeptScheduledDays_: JSON parse error for ' + employeeName + ': ' + e.message);
      return null;
    }
  }

  return null;
}


/**
 * Extracts total scheduled hours for a primary employee from preloaded JSON data.
 * Reads the COL_TOTAL_HOURS cell (col D).
 *
 * @param {string} employeeName — Employee name
 * @param {Array}  sheetData    — 2D value array from getDataRange().getValues()
 * @returns {number} Total paid hours, or 0 if not found
 */
function extractCrossDeptHours_(employeeName, sheetData) {
  if (!sheetData || sheetData.length < WEEK_SHEET.DATA_START_ROW) return 0;
  for (let rowIndex = WEEK_SHEET.DATA_START_ROW - 1; rowIndex < sheetData.length; rowIndex++) {
    const row = sheetData[rowIndex];
    const cellEmpName = row[WEEK_SHEET.COL_NAME - 1];
    if (!cellEmpName || cellEmpName.toString().trim() !== employeeName.toString().trim()) continue;
    const totalHoursCell = row[WEEK_SHEET.COL_TOTAL_HOURS - 1];
    if (totalHoursCell && !isNaN(parseFloat(totalHoursCell))) return parseFloat(totalHoursCell);
    return 0;
  }
  return 0;
}


/**
 * Builds an EmployeeRecord object from a raw employee row and preloaded sheet data.
 * Shared between the primary employee and combo participant loading paths.
 *
 * @param {Object} rawEmployee         — From getActiveEmployees_()
 * @param {string} normalizedTargetDept
 * @param {Date}   weekStartDate
 * @param {Array}  secondaryDepartments — Already-parsed and normalized
 * @param {number} crossDeptHours
 * @param {Object|null} crossDeptScheduledDays
 * @param {boolean} isComboParticipant
 * @param {string}  homeDepartment       — Raw home dept name (for combo participants)
 * @returns {Object} EmployeeRecord
 */
function buildEmployeeRecord_(rawEmployee, _normalizedTargetDept, _weekStartDate,
  secondaryDepartments, crossDeptHours, crossDeptScheduledDays,
  isComboParticipant, homeDepartment) {
  const qualifiedShiftList = parseQualifiedShiftList_(rawEmployee.qualifiedShifts, rawEmployee.preferredShift);
  const vacationDateStrings = parseVacationDateStrings_(rawEmployee.vacationDates);

  let hireDate = new Date();
  if (rawEmployee.hireDate) {
    const parsed = new Date(rawEmployee.hireDate);
    if (!isNaN(parsed.getTime())) hireDate = parsed;
  }

  const seniorityReferenceDate = new Date(SENIORITY.REFERENCE_DATE_STRING);
  const seniorityBase = rawEmployee.ftpt === 'FT' ? SENIORITY.FT_BASE : SENIORITY.PT_BASE;
  const seniorityDaysFromHire = Math.max(0, Math.floor(
    (seniorityReferenceDate - hireDate) / (1000 * 60 * 60 * 24)
  ));

  return {
    name: rawEmployee.name,
    employeeId: rawEmployee.id,
    hireDate: hireDate,
    status: rawEmployee.ftpt || 'PT',
    dayOffPreferenceOne: rawEmployee.dayOffPrefOne || '',
    dayOffPreferenceTwo: rawEmployee.dayOffPrefTwo || '',
    preferredShift: (rawEmployee.preferredShift || '').toLowerCase(),
    qualifiedShifts: qualifiedShiftList,
    vacationDateStrings: vacationDateStrings,
    seniorityRank: seniorityBase + seniorityDaysFromHire,
    department: normalizeDeptName_(rawEmployee.department),
    primaryRole: rawEmployee.role || '',
    secondaryDepartments: secondaryDepartments,
    crossDeptHoursAlreadyScheduled: crossDeptHours,
    crossDeptScheduledDays: crossDeptScheduledDays,
    isComboParticipant: isComboParticipant,
    homeDepartment: homeDepartment || normalizeDeptName_(rawEmployee.department),
    homeScheduledDays: crossDeptScheduledDays, // alias for clarity in Phase 1 combo sub-loop
  };
}


/**
 * Reads active employees for the given department from the Employees sheet
 * and returns them sorted by seniority (descending).
 *
 * Also scans for "combo participants": employees whose primary department is
 * elsewhere but have this department listed in column N (secondary departments).
 * These employees appear in the returned list with isComboParticipant: true.
 *
 * Cross-dept data is resolved from preloadedSheets — no additional sheet reads.
 *
 * @param {string} deptName         — Department name
 * @param {Date}   weekStartDate    — The Monday of the week
 * @param {Object} preloadedSheets  — { sheetName: 2DValueArray } from preloadWeekSheets_
 * @returns {Array<EmployeeRecord>}
 */
function loadRosterSortedBySeniority_(deptName, weekStartDate, preloadedSheets) {
  const normalizedTarget = normalizeDeptName_(deptName);
  const allActiveEmployees = getActiveEmployees_(); // ukgImport.js

  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const employeesSheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME);
  const employeesData = employeesSheet ? employeesSheet.getDataRange().getValues() : [];

  // Build a lookup from employee ID → column N secondary departments string
  const secondaryDeptByEmployeeId = {};
  if (employeesData && employeesData.length > 0) {
    for (let rowIdx = EMPLOYEES_DATA_START_ROW - 1; rowIdx < employeesData.length; rowIdx++) {
      const sheetRow = employeesData[rowIdx];
      const rowEmpId = sheetRow[EMPLOYEE_COLUMN.ID - 1];
      if (rowEmpId) {
        secondaryDeptByEmployeeId[rowEmpId.toString().trim()] =
          (sheetRow[EMPLOYEE_COLUMN.SECONDARY_DEPARTMENTS - 1] || '').toString().trim();
      }
    }
  }

  // Helper: given a raw employee, return their parsed secondary dept list (normalized)
  function parseSecondaryDepts_(rawEmployee) {
    const raw = secondaryDeptByEmployeeId[rawEmployee.id.toString().trim()] || '';
    if (!raw) return [];
    return raw.split(',').map(function (d) { return normalizeDeptName_(d); }).filter(Boolean);
  }

  // Helper: week sheet name for a given dept
  const month = String(weekStartDate.getMonth() + 1).padStart(2, '0');
  const dayStr = String(weekStartDate.getDate()).padStart(2, '0');
  const year = String(weekStartDate.getFullYear()).slice(-2);
  const weekPrefix = 'Week_' + month + '_' + dayStr + '_' + year + '_';
  function weekSheetNameForDept_(dept) {
    // Normalize to match how sheets are actually created
    // Instead, find the sheet using case-insensitively by iterating
    const normalized = normalizeDeptName_(dept);
    for (const key in preloadedSheets) {
      if (key.startsWith(weekPrefix) && normalizeDeptName_(key.slice(weekPrefix.length)) === normalized) {
        return key;
      }
    }
    return weekPrefix + normalized; // fallback
  }

  // --- Primary employees (home dept === target dept) ---
  const primaryEmployees = allActiveEmployees
    .filter(function (rawEmployee) {
      return normalizeDeptName_(rawEmployee.department) === normalizedTarget;
    })
    .map(function (rawEmployee) {
      const secondaryDepartments = parseSecondaryDepts_(rawEmployee);
      let crossDeptHours = 0;
      // Primary employees can appear as combo participants in OTHER depts. We check
      // if any other dept's week sheet already has hours for this employee.
      if (secondaryDepartments.length > 0 && preloadedSheets) {
        secondaryDepartments.forEach(function (secondaryDept) {
          const sheetName = weekSheetNameForDept_(secondaryDept);
          if (preloadedSheets[sheetName]) {
            crossDeptHours += extractCrossDeptHours_(rawEmployee.name, preloadedSheets[sheetName]);
          }
        });
      }
      return buildEmployeeRecord_(rawEmployee, normalizedTarget, weekStartDate,
        secondaryDepartments, crossDeptHours, null,
        false, null);
    })
    .filter(function (employee) {
      var firstRole = (employee.primaryRole || '').split(',')[0].trim().toLowerCase();
      return firstRole !== 'manager';
    });

  // --- Combo participants (secondary dept === target dept, home dept !== target dept) ---
  const comboParticipants = [];
  allActiveEmployees.forEach(function (rawEmployee) {
    // Skip if this employee's home dept is already the target dept (already in primaryEmployees)
    if (normalizeDeptName_(rawEmployee.department) === normalizedTarget) return;

    const secondaryDepartments = parseSecondaryDepts_(rawEmployee);
    if (secondaryDepartments.indexOf(normalizedTarget) === -1) return;

    // Exclude managers
    const firstRole = (rawEmployee.role || '').split(',')[0].trim().toLowerCase();
    if (firstRole === 'manager') return;

    // Look up the employee's home dept week sheet in the preloaded map
    const homeDeptNormalized = normalizeDeptName_(rawEmployee.department);
    const homeDeptSheetName = weekSheetNameForDept_(homeDeptNormalized);
    let homeScheduledDays = null;
    if (preloadedSheets && preloadedSheets[homeDeptSheetName]) {
      homeScheduledDays = extractHomeDeptScheduledDays_(rawEmployee.name, homeDeptSheetName, preloadedSheets);
    }

    if (homeScheduledDays === null) {
      // Home dept sheet doesn't exist yet — skip this run; log so manager knows.
      console.warn('loadRosterSortedBySeniority_: combo participant "' + rawEmployee.name +
        '" skipped — home dept "' + rawEmployee.department + '" sheet not yet generated.');
      return;
    }

    comboParticipants.push(buildEmployeeRecord_(rawEmployee, normalizedTarget, weekStartDate,
      secondaryDepartments, 0, homeScheduledDays,
      true, homeDeptNormalized));
  });

  // Supervisors first, then by primary role alphabetically, then by seniority within role
  primaryEmployees.sort(compareEmployeesForScheduleOrder_);
  comboParticipants.sort(compareEmployeesForScheduleOrder_);

  console.log('loadRosterSortedBySeniority_: ' + primaryEmployees.length + ' primary, ' + comboParticipants.length + ' combo participants');
  return primaryEmployees.concat(comboParticipants);
}

/**
 * Creates the initial WeekGrid with all cells set to OFF, then stamps vacation locks.
 */
function initializeWeekGrid_(employeeList, weekStartDate) {
  const weekGrid = [];
  employeeList.forEach(function (employee, employeeIndex) {
    weekGrid[employeeIndex] = [];
    DAY_NAMES_IN_ORDER.forEach(function (_dayName, _dayIndex) {
      weekGrid[employeeIndex][_dayIndex] = createDayAssignment_('OFF', null, 0, false);
    });
    applyVacationLocksForEmployee_(weekGrid, employeeIndex, employee, weekStartDate);
  });
  return weekGrid;
}

function applyVacationLocksForEmployee_(weekGrid, employeeIndex, employee, weekStartDate) {
  employee.vacationDateStrings.forEach(function (vacationDateString) {
    const vacationDate = parseVacationDateString_(vacationDateString, weekStartDate);
    if (!vacationDate) return;
    const dayIndex = getDayIndexForDate_(vacationDate, weekStartDate);
    if (dayIndex === -1) return;
    weekGrid[employeeIndex][dayIndex] = createDayAssignment_('VAC', null, 0, true);
  });
}

// getCrossDeptHoursForWeek_ and getCrossDeptScheduledDays_ removed in v0.7.0.
// Cross-dept data is now resolved from the preloaded week sheet map in
// loadRosterSortedBySeniority_ — see extractCrossDeptHours_ and extractHomeDeptScheduledDays_.


// ---------------------------------------------------------------------------
// Phase 1: Preference & Hour Assignment (merged old Phases 1 + 2)
// ---------------------------------------------------------------------------

/**
 * Single pass through regularEmployees in seniority order:
 *   1. Grant requested days off (if coverage allows)
 *   2. Assign preferred shift on all remaining available days
 *   3. Enforce minimum days off (VAC/RDO fix: VAC counts as working, only SHIFT removed)
 *   4. Top up hours to weekly minimum on lowest-coverage open days
 *
 * Followed by a separate sub-loop for combo participants:
 *   - Mirror home dept scheduled days (assign secondary shift where homeScheduledDays = true)
 *   - Skip day-off preference granting and hour enforcement (home dept owns those)
 *
 * @param {Array}  weekGrid
 * @param {Array}  regularEmployees    — Primary employees only (no combo participants)
 * @param {Array}  comboParticipants   — Employees with isComboParticipant: true
 * @param {Object} shiftTimingMap
 * @param {Object} staffingRequirements
 * @param {Date}   weekStartDate
 * @param {Object} staggerMap          — { dayName: { shiftKey: [startTimes], ... } }
 * @param {Object} staggerPositions    — Shared tracker (mutated in place)
 * @param {Object} dayTrafficLevels    — { dayName: 'Low'|'Moderate'|'High' }
 */
function runPhaseOnePreferenceAndHours_(weekGrid, regularEmployees, comboParticipants,
  shiftTimingMap, staffingRequirements, _weekStartDate,
  staggerMap, staggerPositions, _dayTrafficLevels) {

  // Pre-compute worker count per day (O(n) once; decremented when SHIFT is removed).
  // Used by enforceMinimumDaysOff_ to avoid the O(n²) inner scan.
  const workerCountByDay = new Array(WEEK_SHEET.DAYS_IN_WEEK).fill(0);
  regularEmployees.forEach(function (_employee, employeeIndex) {
    DAY_NAMES_IN_ORDER.forEach(function (_dayName, dayIndex) {
      if (weekGrid[employeeIndex][dayIndex].type === 'SHIFT') workerCountByDay[dayIndex]++;
    });
  });

  // --- Regular employee loop ---
  regularEmployees.forEach(function (employee, employeeIndex) {

    // 1. Grant requested days off
    DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
      const cell = weekGrid[employeeIndex][dayIndex];
      if (cell.locked) return;
      if (cell.type !== 'OFF') return;
      const requested = employee.dayOffPreferenceOne === dayName || employee.dayOffPreferenceTwo === dayName;
      if (!requested) return;
      const minimumRequired = (staffingRequirements[dayName] && staffingRequirements[dayName].value) || 0;
      const availableCount = workerCountByDay[dayIndex];
      if (availableCount > minimumRequired) {
        weekGrid[employeeIndex][dayIndex] = createDayAssignment_('RDO', null, 0, false);
        // Do not decrement workerCountByDay — RDO was OFF, not a SHIFT removal
      }
    });

    // 2. Assign preferred shift on all remaining OFF days
    DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
      const cell = weekGrid[employeeIndex][dayIndex];
      if (cell.type !== 'OFF') return;
      if (cell.locked) return;
      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (!shiftDef) return;

      const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName); // settingsManager.js
      let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);

      if (staggerMap && staggerMap[dayName]) {
        const shiftKey = shiftDef.name + '|' + shiftDef.status;
        const startTimes = staggerMap[dayName][shiftKey];
        if (startTimes && startTimes.length > 0) {
          const posKey = dayName + '|' + shiftKey;
          if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
          const staggerIndex = staggerPositions[posKey] % startTimes.length;
          displayText = buildShiftDisplayText_(startTimes[staggerIndex], shiftDef.paidHours, shiftDef.hasLunch);
          staggerPositions[posKey]++;
        }
      }

      weekGrid[employeeIndex][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
      workerCountByDay[dayIndex]++;
    });

    // 3. Enforce minimum days off.
    //    VAC = working day (employee still needs MIN_DAYS_OFF separate days off).
    //    Days off = cells where type === 'OFF' || type === 'RDO'.
    //    Only remove SHIFT cells — never VAC.
    const maxWorkingDays = WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF; // config.js
    const shiftDayIndices = [];
    for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
      if (weekGrid[employeeIndex][dayIndex].type === 'SHIFT') shiftDayIndices.push(dayIndex);
    }
    // Count actual days off (OFF + RDO, NOT VAC)
    let daysOffCount = 0;
    for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
      const cellType = weekGrid[employeeIndex][dayIndex].type;
      if (cellType === 'OFF' || cellType === 'RDO') daysOffCount++;
    }
    const excessShifts = shiftDayIndices.length - maxWorkingDays;
    if (excessShifts > 0) {
      // Remove from highest-coverage days first (least impact on others)
      const scoredShiftDays = shiftDayIndices.map(function (dayIndex) {
        return { dayIndex: dayIndex, coverage: workerCountByDay[dayIndex] };
      });
      scoredShiftDays.sort(function (a, b) { return b.coverage - a.coverage; });
      for (let removeIndex = 0; removeIndex < excessShifts; removeIndex++) {
        const removeDayIndex = scoredShiftDays[removeIndex].dayIndex;
        if (weekGrid[employeeIndex][removeDayIndex].locked) continue;
        weekGrid[employeeIndex][removeDayIndex] = createDayAssignment_('OFF', null, 0, false);
        workerCountByDay[removeDayIndex]--;
      }
    }

    // 4. Top up hours to weekly minimum on lowest-coverage open days
    const weeklyMin = employee.status === 'FT' ? HOUR_RULES.FT_MIN :
      (employee.status === 'LPT' ? HOUR_RULES.LPT_MIN : HOUR_RULES.PT_MIN);
    const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX :
      (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
    const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
    let currentHours = getWeeklyHours_(weekGrid, employeeIndex);

    if (currentHours < weeklyMin) {
      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (shiftDef) {
        // Find open days sorted by coverage ascending (fill least-covered gaps first)
        const openDays = [];
        DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
          const cell = weekGrid[employeeIndex][dayIndex];
          if (cell.type !== 'OFF' || cell.locked) return;
          if (countWorkingDays_(weekGrid, employeeIndex) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) return;
          openDays.push({ dayName: dayName, dayIndex: dayIndex, coverage: workerCountByDay[dayIndex] });
        });
        openDays.sort(function (a, b) { return a.coverage - b.coverage; });

        openDays.forEach(function (openDay) {
          if (currentHours >= weeklyMin) return;
          if (currentHours + shiftDef.paidHours > effectiveMax) return;
          if (countWorkingDays_(weekGrid, employeeIndex) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) return;

          const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, openDay.dayName);
          let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);

          if (staggerMap && staggerMap[openDay.dayName]) {
            const shiftKey = shiftDef.name + '|' + shiftDef.status;
            const startTimes = staggerMap[openDay.dayName][shiftKey];
            if (startTimes && startTimes.length > 0) {
              const posKey = openDay.dayName + '|' + shiftKey;
              if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
              const staggerIndex = staggerPositions[posKey] % startTimes.length;
              displayText = buildShiftDisplayText_(startTimes[staggerIndex], shiftDef.paidHours, shiftDef.hasLunch);
              staggerPositions[posKey]++;
            }
          }

          weekGrid[employeeIndex][openDay.dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
          workerCountByDay[openDay.dayIndex]++;
          currentHours += shiftDef.paidHours;
        });
      }
    }
  });

  // --- Combo participant sub-loop ---
  // Mirror the home dept's scheduled days: assign secondary shift where homeScheduledDays[day] = true.
  // Day-off preference granting and hour enforcement are skipped — home dept owns those.
  comboParticipants.forEach(function (employee, comboIndex) {
    const employeeIndex = regularEmployees.length + comboIndex;
    const homeScheduledDays = employee.homeScheduledDays;

    if (!homeScheduledDays) {
      // Home dept not yet generated — already warned in loadRosterSortedBySeniority_; skip.
      return;
    }

    DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
      if (!homeScheduledDays[dayName]) return; // Home dept has this day off — mirror it
      const cell = weekGrid[employeeIndex][dayIndex];
      if (cell.type !== 'OFF') return; // Already VAC or manually set
      if (cell.locked) return;

      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (!shiftDef) {
        console.warn('runPhaseOnePreferenceAndHours_: combo participant "' + employee.name +
          '" has no resolvable shift in secondary dept map — skipping day ' + dayName);
        return;
      }

      // Combo participants use a fixed secondary shift time — no stagger
      const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName);
      const displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);
      weekGrid[employeeIndex][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
    });
  });
}


// ---------------------------------------------------------------------------
// Phase 2: Gap Resolution (rewritten — single-pass priority fill)
// ---------------------------------------------------------------------------

/**
 * Fills coverage gaps using a priority queue ordered by gap depth (most urgent first).
 * Coverage is built once per day and updated incrementally — no full rebuild per fill.
 * A hard iteration cap prevents the unbounded O(n³) cascade of the previous implementation.
 *
 * Combo participants are excluded from the candidate pool (their days are fixed by the
 * home dept), but their shifts ARE counted in buildDayCoverage_ because they physically
 * contribute coverage in the secondary dept.
 *
 * Skip entirely when engineOptions.gapFillEnabled === false.
 *
 * @param {Array}  weekGrid
 * @param {Array}  regularEmployees  — Primary employees only (combo participants excluded as candidates)
 * @param {Array}  allEmployees      — Full employee list (primary + combo; used for coverage counting)
 * @param {Object} shiftTimingMap
 * @param {Object} staffingRequirements
 * @param {Object} staggerMap
 * @param {Object} staggerPositions   — Shared tracker (mutated in place)
 * @param {Object} dayTrafficLevels
 * @param {Object} engineOptions      — { gapFillEnabled: boolean, ... }
 */
function runPhaseGapResolution_(weekGrid, regularEmployees, allEmployees, shiftTimingMap,
  _staffingRequirements, staggerMap, staggerPositions,
  _dayTrafficLevels, engineOptions) {
  if (!engineOptions.gapFillEnabled) {
    console.log('Phase 2 (gap): disabled by engineOptions.gapFillEnabled — skipped');
    return;
  }

  const MAX_ITERS_PER_DAY = Math.max(50, regularEmployees.length * 2);

  DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
    const coverageWindow = COVERAGE_WINDOW[dayName] || { startMinute: 240, endMinute: 1410 }; // config.js
    const windowStartSlot = Math.max(0, Math.floor(
      (coverageWindow.startMinute - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
    const windowEndSlot = Math.min(COVERAGE.SLOT_COUNT, Math.floor(
      (coverageWindow.endMinute - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));

    // Coverage built ONCE per day — updated incrementally after each successful fill
    let dayCoverage = buildDayCoverage_(weekGrid, allEmployees, dayIndex, shiftTimingMap, -1);
    if (!hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot)) return;

    let iterationCount = 0;
    let gapsRemain = true;

    while (gapsRemain && iterationCount < MAX_ITERS_PER_DAY) {
      iterationCount++;
      let filled = false;

      // Candidate A — reassign a working regular employee to a better shift
      for (let employeeIndex = regularEmployees.length - 1; employeeIndex >= 0; employeeIndex--) {
        const cell = weekGrid[employeeIndex][dayIndex];
        if (cell.type !== 'SHIFT' || cell.locked) continue;
        const employee = regularEmployees[employeeIndex];
        const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX :
          (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
        const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
        const currentHours = getWeeklyHours_(weekGrid, employeeIndex);

        // Score coverage without this employee to find where adding them would help most
        const coverageWithout = buildDayCoverage_(weekGrid, allEmployees, dayIndex, shiftTimingMap, employeeIndex);
        const best = selectBestCoverageShift_(employee.qualifiedShifts, employee.status, coverageWithout, shiftTimingMap);
        if (!best) continue;

        const currentDef = shiftTimingMap[(cell.shiftName || '') + '|' + employee.status];
        const currentScore = currentDef ? scoreCoverageForShift_(currentDef, coverageWithout) : 0;
        if (scoreCoverageForShift_(best, coverageWithout) <= currentScore) continue;
        if (currentHours + (best.paidHours - cell.paidHours) > effectiveMax) continue;

        const dayAnchorMinutes = getStartMinutesForDay_(best, dayName);
        let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + best.blockMinutes);
        if (staggerMap && staggerMap[dayName]) {
          const shiftKey = best.name + '|' + best.status;
          const startTimes = staggerMap[dayName][shiftKey];
          if (startTimes && startTimes.length > 0) {
            const posKey = dayName + '|' + shiftKey;
            if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
            displayText = buildShiftDisplayText_(startTimes[staggerPositions[posKey] % startTimes.length],
              best.paidHours, best.hasLunch);
            staggerPositions[posKey]++;
          }
        }

        // Incremental coverage update: remove old shift contribution, add new one
        const oldDef = shiftTimingMap[(cell.shiftName || '') + '|' + employee.status];
        if (oldDef) applyCoverageChange_(dayCoverage, oldDef, -1);
        weekGrid[employeeIndex][dayIndex] = createDayAssignment_('SHIFT', best.name, best.paidHours, false, displayText);
        applyCoverageChange_(dayCoverage, best, +1);
        filled = true;
        break;
      }

      if (!filled) {
        // Candidate B — pull in a regular employee who is currently off
        for (let employeeIndex = regularEmployees.length - 1; employeeIndex >= 0; employeeIndex--) {
          const cell = weekGrid[employeeIndex][dayIndex];
          if (cell.type !== 'OFF' || cell.locked) continue;
          if (countWorkingDays_(weekGrid, employeeIndex) >= WEEK_SHEET.DAYS_IN_WEEK - SCHEDULE_RULES.MIN_DAYS_OFF) continue;
          const employee = regularEmployees[employeeIndex];
          const weeklyMax = employee.status === 'FT' ? HOUR_RULES.FT_MAX :
            (employee.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
          const effectiveMax = weeklyMax - (employee.crossDeptHoursAlreadyScheduled || 0);
          const currentHours = getWeeklyHours_(weekGrid, employeeIndex);
          const best = selectBestCoverageShift_(employee.qualifiedShifts, employee.status, dayCoverage, shiftTimingMap);
          if (!best) continue;
          if (currentHours + best.paidHours > effectiveMax) continue;

          const dayAnchorMinutes = getStartMinutesForDay_(best, dayName);
          let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + best.blockMinutes);
          if (staggerMap && staggerMap[dayName]) {
            const shiftKey = best.name + '|' + best.status;
            const startTimes = staggerMap[dayName][shiftKey];
            if (startTimes && startTimes.length > 0) {
              const posKey = dayName + '|' + shiftKey;
              if (!staggerPositions[posKey]) staggerPositions[posKey] = 0;
              displayText = buildShiftDisplayText_(startTimes[staggerPositions[posKey] % startTimes.length],
                best.paidHours, best.hasLunch);
              staggerPositions[posKey]++;
            }
          }

          weekGrid[employeeIndex][dayIndex] = createDayAssignment_('SHIFT', best.name, best.paidHours, false, displayText);
          applyCoverageChange_(dayCoverage, best, +1);
          filled = true;
          break;
        }
      }

      gapsRemain = filled && hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot);
    }

    console.log('Phase 2 (gap): ' + dayName + ' — ' + iterationCount + ' iteration(s)' +
      (hasCoverageGaps_(dayCoverage, windowStartSlot, windowEndSlot) ? ' [gaps remain]' : ' [resolved]'));
  });
}


/**
 * Applies an incremental change to a coverage array.
 * delta = +1 adds the shift contribution; -1 removes it.
 *
 * @param {number[]} dayCoverage — Mutable slot array
 * @param {Object}   shiftDef    — Entry from shiftTimingMap
 * @param {number}   delta       — +1 or -1
 */
function applyCoverageChange_(dayCoverage, shiftDef, delta) {
  const startSlot = Math.max(0, Math.floor(
    (shiftDef.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
  const endSlot = Math.min(COVERAGE.SLOT_COUNT, Math.floor(
    (shiftDef.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
  for (let slotIndex = startSlot; slotIndex < endSlot; slotIndex++) {
    dayCoverage[slotIndex] += delta;
  }
}


// ---------------------------------------------------------------------------
// Phase 4: Role Assignment (all employees — runs last so pool + combo get roles)
// ---------------------------------------------------------------------------

/**
 * Stamps roles on all working cells and enforces role minimums per traffic level.
 *
 * Sub-phase 4a — Stamp primary role on every SHIFT cell for every employee.
 *   Runs on the full employeeList (regular + combo + pool) so pool members and
 *   combo participants also receive role stamps.
 *
 * Sub-phase 4b — Enforce role minimums (when engineOptions.enforceRoleMinimums is true):
 *   For each day, for each role in roleMinimums:
 *     needed = roleMinimums[role][trafficLevel]
 *     while deficit > 0: find the most-senior working employee qualified for this role
 *       whose current role would still be covered after they're pulled — reassign them.
 *   When enforceRoleMinimums is false, falls back to the "cover if zero" logic
 *   (same behaviour as the previous engine version).
 *
 * When heatmap is disabled, dayTrafficLevels defaults all days to 'Moderate', so the
 * Moderate column of roleMinimums acts as a static global minimum.
 *
 * @param {Array}  weekGrid
 * @param {Array}  employeeList      — Full list: regular + combo + pool
 * @param {string} _deptName         — Unused; kept for call-site consistency
 * @param {Object} dayTrafficLevels  — { dayName: 'Low'|'Moderate'|'High' }
 * @param {Object} roleMinimums      — { RoleName: { Low, Moderate, High }, ... }
 * @param {Object} engineOptions     — { enforceRoleMinimums: boolean, ... }
 */
function runPhaseRoleAssignment_(weekGrid, employeeList, _deptName,
  dayTrafficLevels, roleMinimums, engineOptions) {

  // Parse each employee's full ordered role list once (comma-separated in column L)
  const employeeRoleLists = employeeList.map(function (employee) {
    return (employee.primaryRole || '')
      .split(',')
      .map(function (roleName) { return roleName.trim(); })
      .filter(Boolean);
  });

  // Sub-phase 4a: stamp primary role on every SHIFT cell for every employee
  employeeList.forEach(function (_employee, employeeIndex) {
    DAY_NAMES_IN_ORDER.forEach(function (_dayName, dayIndex) {
      const cell = weekGrid[employeeIndex][dayIndex];
      cell.role = cell.type === 'SHIFT' ? (employeeRoleLists[employeeIndex][0] || null) : null;
    });
  });

  // Sub-phase 4b: enforce role minimums (or "cover if zero" fallback)
  DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
    const trafficLevel = (dayTrafficLevels && dayTrafficLevels[dayName]) || 'Moderate';

    if (!engineOptions.enforceRoleMinimums || !roleMinimums || Object.keys(roleMinimums).length === 0) {
      // Fallback: "cover if zero" — pull a backup only when a role has ZERO coverage today
      const allRoles = [];
      employeeRoleLists.forEach(function (roleList) {
        roleList.forEach(function (role) {
          if (allRoles.indexOf(role) === -1) allRoles.push(role);
        });
      });
      allRoles.forEach(function (role) {
        const alreadyCovered = employeeList.some(function (_e, employeeIndex) {
          return weekGrid[employeeIndex][dayIndex].role === role;
        });
        if (alreadyCovered) return;
        for (let employeeIndex = 0; employeeIndex < employeeList.length; employeeIndex++) {
          const cell = weekGrid[employeeIndex][dayIndex];
          if (cell.type !== 'SHIFT' || cell.locked) continue;
          if (employeeRoleLists[employeeIndex].indexOf(role) === -1) continue;
          const currentRole = cell.role;
          const currentRoleSurvives = employeeList.some(function (_e, otherIndex) {
            return otherIndex !== employeeIndex && weekGrid[otherIndex][dayIndex].role === currentRole;
          });
          if (!currentRole || !currentRoleSurvives) continue;
          cell.role = role;
          break;
        }
      });
      return;
    }

    // Minimums mode: enforce per-role per-traffic-level counts
    Object.keys(roleMinimums).forEach(function (roleName) {
      const needed = Number((roleMinimums[roleName] && roleMinimums[roleName][trafficLevel]) || 0);
      if (needed <= 0) return;

      // Count how many employees currently have this role today
      let currentCount = 0;
      employeeList.forEach(function (_e, employeeIndex) {
        if (weekGrid[employeeIndex][dayIndex].role === roleName) currentCount++;
      });

      let deficit = needed - currentCount;
      if (deficit <= 0) return;

      // Pull in the most senior qualified employee whose current role stays covered
      for (let employeeIndex = 0; employeeIndex < employeeList.length && deficit > 0; employeeIndex++) {
        const cell = weekGrid[employeeIndex][dayIndex];
        if (cell.type !== 'SHIFT' || cell.locked) continue;
        if (cell.role === roleName) continue; // already contributing
        if (employeeRoleLists[employeeIndex].indexOf(roleName) === -1) continue; // not qualified

        const currentRole = cell.role;
        // Only reassign if the employee's current role still meets its own minimum after they leave
        const currentRoleMinimum = Number(
          (currentRole && roleMinimums[currentRole] && roleMinimums[currentRole][trafficLevel]) || 0
        );
        const currentRoleCount = employeeList.reduce(function (total, _e, otherIndex) {
          return total + (weekGrid[otherIndex][dayIndex].role === currentRole ? 1 : 0);
        }, 0);
        // Allow reassignment if: no minimum set for current role, or current role has surplus
        if (currentRole && currentRoleMinimum > 0 && currentRoleCount <= currentRoleMinimum) continue;

        cell.role = roleName;
        deficit--;
      }
    });
  });

  // Role ratio rules from config.js (e.g., 1 Assist per Cashier) — applied after minimums
  if (typeof ROLE_RULES === 'undefined') return; // config.js — optional
  Object.keys(ROLE_RULES).forEach(function (triggerRole) {
    const rule = ROLE_RULES[triggerRole];
    DAY_NAMES_IN_ORDER.forEach(function (_dayName, dayIndex) {
      let triggerCount = 0, requiredCount = 0;
      employeeList.forEach(function (_employee, employeeIndex) {
        const cell = weekGrid[employeeIndex][dayIndex];
        if (cell.type !== 'SHIFT') return;
        if (cell.role === triggerRole) triggerCount++;
        if (cell.role === rule.requiresRole) requiredCount++;
      });
      const deficit = (triggerCount * rule.ratio) - requiredCount;
      if (deficit <= 0) return;
      let filled = 0;
      for (let employeeIndex = employeeList.length - 1; employeeIndex >= 0 && filled < deficit; employeeIndex--) {
        const cell = weekGrid[employeeIndex][dayIndex];
        if (cell.type !== 'SHIFT' || cell.locked) continue;
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
  const minutes = totalMinutes % 60;
  const period = totalHours >= 12 ? 'PM' : 'AM';
  const twelve = totalHours % 12 === 0 ? 12 : totalHours % 12;
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
// Phase 3: Pool Member Scheduling
// ---------------------------------------------------------------------------

/**
 * Schedules pool members (supervisors, flex workers) based on traffic level and
 * staggered start times. Uses the shared staggerPositions tracker so pool members
 * never duplicate a stagger slot already taken by a regular employee.
 *
 * Cross-dept guard for pool members uses the preloaded sheet map — no live reads.
 *
 * @param {Array}  weekGrid
 * @param {Array}  employeeList       — Full list (regular + combo + pool)
 * @param {Array}  poolMembers
 * @param {Object} heatmapConfig
 * @param {Object} dayTrafficLevels
 * @param {Object} staggerMap
 * @param {Object} staggerPositions   — Shared tracker (mutated in place)
 * @param {Object} shiftTimingMap
 * @param {Date}   weekStartDate
 * @param {Object} preloadedSheets    — { sheetName: 2DValueArray }
 */
function runPhasePoolScheduling_(weekGrid, employeeList, poolMembers, heatmapConfig,
  dayTrafficLevels, staggerMap, staggerPositions,
  shiftTimingMap, _weekStartDate, _preloadedSheets) {
  if (!poolMembers || poolMembers.length === 0) return;

  // Map pool member ID → index in employeeList/weekGrid
  const poolMemberIndexMap = {};
  poolMembers.forEach(function (poolMember) {
    const poolId = (poolMember.employeeId || poolMember.id || '').toString();
    for (let employeeIndex = 0; employeeIndex < employeeList.length; employeeIndex++) {
      const listId = (employeeList[employeeIndex].employeeId || employeeList[employeeIndex].id || '').toString();
      if (listId === poolId) {
        poolMemberIndexMap[poolId] = employeeIndex;
        break;
      }
    }
  });

  DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
    const trafficLevel = dayTrafficLevels[dayName] || 'Moderate';
    const selectedForDay = selectPoolMembers_(poolMembers, trafficLevel, heatmapConfig); // trafficHeatmapEngine.js

    selectedForDay.forEach(function (poolMember) {
      const poolId = (poolMember.employeeId || poolMember.id || '').toString();
      const poolIndex = poolMemberIndexMap[poolId];
      if (poolIndex === undefined || poolIndex < 0) return;

      const cell = weekGrid[poolIndex][dayIndex];
      if (cell.type === 'VAC' || cell.type === 'RDO') return;
      if (cell.type === 'SHIFT') return;
      if (cell.type !== 'OFF') return;

      const weeklyHours = getWeeklyHours_(weekGrid, poolIndex);
      const weeklyMax = poolMember.status === 'FT' ? HOUR_RULES.FT_MAX :
        (poolMember.status === 'LPT' ? HOUR_RULES.LPT_MAX : HOUR_RULES.PT_MAX);
      const effectiveMax = weeklyMax - (poolMember.crossDeptHoursAlreadyScheduled || 0);

      const dayCoverage = buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, poolIndex);
      const bestShift = selectBestCoverageShift_(poolMember.qualifiedShifts || [], poolMember.status, dayCoverage, shiftTimingMap);
      if (!bestShift) return;
      if (weeklyHours + bestShift.paidHours > effectiveMax) return;

      const shiftKey = bestShift.name + '|' + poolMember.status;
      const staggerKey = dayName + '|' + shiftKey;
      const startTimes = (staggerMap[dayName] && staggerMap[dayName][shiftKey]) ? staggerMap[dayName][shiftKey] : [];

      const dayAnchorMinutes = getStartMinutesForDay_(bestShift, dayName); // settingsManager.js
      let displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + bestShift.blockMinutes);
      if (startTimes.length > 0) {
        if (!staggerPositions[staggerKey]) staggerPositions[staggerKey] = 0;
        displayText = buildShiftDisplayText_(startTimes[staggerPositions[staggerKey] % startTimes.length],
          bestShift.paidHours, bestShift.hasLunch);
        staggerPositions[staggerKey]++;
      }

      weekGrid[poolIndex][dayIndex] = createDayAssignment_('SHIFT', bestShift.name, bestShift.paidHours, false, displayText);
    });
  });

  console.log('Phase 3 (pool): assigned ' + Object.keys(poolMemberIndexMap).length + ' pool members with shared stagger');
}


// ---------------------------------------------------------------------------
// Coverage Map Functions
// ---------------------------------------------------------------------------

function buildDayCoverage_(weekGrid, employeeList, dayIndex, shiftTimingMap, excludeIndex) {
  const slots = new Array(COVERAGE.SLOT_COUNT).fill(0); // config.js
  employeeList.forEach(function (employee, ei) {
    if (ei === excludeIndex) return;
    const cell = weekGrid[ei][dayIndex];
    if (cell.type !== 'SHIFT') return;
    const shiftDef = shiftTimingMap[cell.shiftName + '|' + employee.status];
    if (!shiftDef) return;
    const startSlot = Math.max(0, Math.floor((shiftDef.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
    const endSlot = Math.min(COVERAGE.SLOT_COUNT, Math.floor((shiftDef.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
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
  qualifiedShiftNames.forEach(function (name) {
    const def = shiftTimingMap[name.trim() + '|' + status];
    if (!def) return;
    const score = scoreCoverageForShift_(def, coverageSlots);
    if (score > highScore) { highScore = score; best = def; }
  });
  return best;
}

function scoreCoverageForShift_(shiftDef, coverageSlots) {
  const startSlot = Math.max(0, Math.floor((shiftDef.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
  const endSlot = Math.min(COVERAGE.SLOT_COUNT, Math.floor((shiftDef.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES));
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
  return weekGrid.map(function (employeeRow, ei) {
    return employeeRow.map(function (cell) {
      return {
        type: cell.type,
        shiftName: cell.shiftName || null,
        paidHours: cell.paidHours || 0,
        locked: cell.locked || false,
        displayText: cell.displayText || null,
        role: cell.role || null,
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
  return employeeList.map(function (emp, i) {
    var weeklyHours = getWeeklyHours_(weekGrid, i);
    var minHours = emp.status === 'FT' ? HOUR_RULES.FT_MIN : (emp.status === 'LPT' ? HOUR_RULES.LPT_MIN : HOUR_RULES.PT_MIN);
    return {
      name: emp.name,
      employeeId: emp.employeeId,
      status: emp.status,
      department: emp.department,
      seniorityRank: emp.seniorityRank,
      primaryRole: emp.primaryRole || '',
      weeklyHours: weeklyHours,
      underHours: weeklyHours < minHours,
      secondaryDepartments: emp.secondaryDepartments || [],
    };
  });
}


// ---------------------------------------------------------------------------
// Utility Functions
// ---------------------------------------------------------------------------

function getWeeklyHours_(weekGrid, ei) {
  let total = 0;
  weekGrid[ei].forEach(function (cell) { if (cell.type === 'SHIFT') total += cell.paidHours; });
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

function createDayAssignment_(type, shiftName, paidHours, locked, displayText, role) {
  return { type, shiftName: shiftName || null, paidHours: paidHours || 0, locked: locked || false, displayText: displayText || null, role: role || null };
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
  const day = String(weekStartDate.getDate()).padStart(2, '0');
  const year = String(weekStartDate.getFullYear()).slice(-2);
  const base = 'Week_' + month + '_' + day + '_' + year;
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
  const list = raw.toString().split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  if (preferred && !list.includes(preferred.toString().trim())) {
    list.unshift(preferred.toString().trim());
  }
  return list;
}

function getDayIndexForDate_(targetDate, weekStartDate) {
  const weekStart = new Date(weekStartDate); weekStart.setHours(0, 0, 0, 0);
  const target = new Date(targetDate); target.setHours(0, 0, 0, 0);
  const diff = Math.round((target.getTime() - weekStart.getTime()) / 86400000);
  return (diff < 0 || diff > 6) ? -1 : diff;
}

function readJsonScheduleFromSheet_(weekSheet, employeeList) {
  const { employeeRows } = readJsonSchedule(weekSheet); // formatter.js
  const weekGrid = [];

  employeeList.forEach(function (employee, employeeIndex) {
    weekGrid[employeeIndex] = [];

    const row = employeeRows.find(function (r) {
      return String(r.employeeId).trim() === String(employee.employeeId).trim();
    });

    if (!row) {
      // Employee not in sheet (added to roster since last generation) — fill with OFF
      DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
        weekGrid[employeeIndex][dayIndex] = createDayAssignment_('OFF', null, 0, false);
      });
      return;
    }

    try {
      const scheduleObj = JSON.parse(row.scheduleJson);
      DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
        const dayData = scheduleObj[dayName];
        if (!dayData) {
          weekGrid[employeeIndex][dayIndex] = createDayAssignment_('OFF', null, 0, false);
          return;
        }

        weekGrid[employeeIndex][dayIndex] = createDayAssignment_(
          dayData.type,
          dayData.shiftName || null,
          dayData.paidHours || 0,
          dayData.locked || false,
          dayData.displayText || null,
          dayData.role || null
        );
      });
    } catch (e) {
      console.warn('readJsonScheduleFromSheet_: failed to parse JSON for ' + employee.name + ': ' + e.message);
      DAY_NAMES_IN_ORDER.forEach(function (_dayName, dayIndex) {
        weekGrid[employeeIndex][dayIndex] = createDayAssignment_('OFF', null, 0, false);
      });
    }
  });

  return weekGrid;
}


// ---------------------------------------------------------------------------
// Targeted Hybrid Staff Append
// ---------------------------------------------------------------------------

/**
 * Appends combo participant rows to an already-generated secondary dept schedule
 * without touching any existing primary employee rows.
 *
 * Idempotent: if a combo participant already has rows in the sheet, they are
 * overwritten in place rather than duplicated.
 *
 * Algorithm:
 *   1. Preload week sheets (picks up the now-generated home dept sheet)
 *   2. Load only combo participants for this dept (employees with this dept in col N)
 *   3. For each combo participant: mirror home dept scheduled days with secondary shift
 *   4. Re-run role stamp (Phase 4a) on all employees so role counts reflect the additions
 *   5. Delegate write to formatter.js: appendComboParticipantsToSheet_()
 *
 * @param {string} deptName
 * @param {Date}   weekStartDate
 * @returns {{ added: number, skipped: Array<{name, homeDept}>, weekGrid, employeeList }}
 */
function appendHybridEmployees_(deptName, weekStartDate) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = generateWeekSheetName_(weekStartDate, deptName);
  const weekSheet = workbook.getSheetByName(sheetName);
  if (!weekSheet) {
    throw new Error('appendHybridEmployees_: no schedule found for "' + deptName +
      '" — generate the schedule first.');
  }

  // Batch read all week sheets (home dept should now be present)
  const preloadedSheets = preloadAllWeekSheets_(workbook);

  // Load the full roster — this includes newly-resolvable combo participants
  // whose home dept sheet now exists in preloadedSheets.
  const employeeList = loadRosterSortedBySeniority_(deptName, weekStartDate, preloadedSheets);
  const primaryEmployees = employeeList.filter(function (emp) { return !emp.isComboParticipant; });
  const comboParticipants = employeeList.filter(function (emp) { return emp.isComboParticipant; });

  // Reconstruct the grid from the existing sheet for primary employees
  const weekGrid = readCheckboxStateFromSheet_(weekSheet, primaryEmployees.length);

  // Add placeholder rows for combo participants (initialized to OFF)
  comboParticipants.forEach(function (_employee, comboIndex) {
    const employeeIndex = primaryEmployees.length + comboIndex;
    weekGrid[employeeIndex] = [];
    DAY_NAMES_IN_ORDER.forEach(function (_dayName, dayIndex) {
      weekGrid[employeeIndex][dayIndex] = createDayAssignment_('OFF', null, 0, false);
    });
    applyVacationLocksForEmployee_(weekGrid, employeeIndex, _employee, weekStartDate);
  });

  const shiftTimingMap = buildShiftTimingMap(deptName); // settingsManager.js

  // Mirror home dept scheduled days onto each combo participant
  const skipped = [];
  comboParticipants.forEach(function (employee, comboIndex) {
    const employeeIndex = primaryEmployees.length + comboIndex;

    if (!employee.homeScheduledDays) {
      skipped.push({ name: employee.name, homeDept: employee.homeDepartment });
      return;
    }

    DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
      if (!employee.homeScheduledDays[dayName]) return;
      const cell = weekGrid[employeeIndex][dayIndex];
      if (cell.type !== 'OFF' || cell.locked) return;

      const shiftDef = resolveShiftForEmployee_(employee, shiftTimingMap);
      if (!shiftDef) {
        console.warn('appendHybridEmployees_: no secondary shift for "' + employee.name + '" on ' + dayName);
        return;
      }
      const dayAnchorMinutes = getStartMinutesForDay_(shiftDef, dayName);
      const displayText = formatMinutesAsTimeRange(dayAnchorMinutes, dayAnchorMinutes + shiftDef.blockMinutes);
      weekGrid[employeeIndex][dayIndex] = createDayAssignment_('SHIFT', shiftDef.name, shiftDef.paidHours, false, displayText);
    });
  });

  // Re-stamp roles on all employees so role counts include combo participants
  const engineOptions = loadEngineOptions(deptName); // settingsManager.js
  const roleMinimums = loadRoleMinimums(deptName);  // settingsManager.js
  const dayTrafficLevels = {};
  DAY_NAMES_IN_ORDER.forEach(function (dayName) { dayTrafficLevels[dayName] = 'Moderate'; });
  runPhaseRoleAssignment_(weekGrid, employeeList, deptName, dayTrafficLevels, roleMinimums, engineOptions);

  // Write combo participant rows to the sheet (formatter.js)
  if (comboParticipants.length > 0 && skipped.length < comboParticipants.length) {
    appendComboParticipantsToSheet_(weekSheet, weekGrid, employeeList, primaryEmployees.length); // formatter.js
  }

  const comboParticipantIdSet = new Set();
  comboParticipants.forEach(function (emp) { if (emp.employeeId) comboParticipantIdSet.add(emp.employeeId); });

  return {
    added: comboParticipants.length - skipped.length,
    skipped: skipped,
    weekGrid: weekGrid,
    employeeList: employeeList,
    comboParticipantIds: comboParticipantIdSet,
  };
}
