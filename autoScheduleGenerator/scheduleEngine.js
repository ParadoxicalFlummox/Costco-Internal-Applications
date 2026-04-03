/**
 * scheduleEngine.js — The core schedule generation algorithm.
 * VERSION 0.3.0
 *
 * This file contains the 4-phase algorithm that turns a roster of employees and
 * a set of shift/staffing rules into a complete weekly schedule grid.
 *
 * THE FOUR PHASES:
 *   Phase 0 — Bootstrap: Load the roster, initialize the grid, stamp vacation locks.
 *   Phase 1 — Preference Assignment: Honor preferred days off and preferred shifts, in seniority order.
 *   Phase 2 — Minimum Hour Enforcement: Add shifts until every employee meets their weekly minimum.
 *   Phase 3 — Gap Resolution: Detect uncovered time slots and fill them via shift reassignment or adding employees.
 *
 * DATA MODEL:
 *   The schedule is represented as a WeekGrid: a 2D array where
 *   weekGrid[employeeIndex][dayIndex] is a DayAssignment object:
 *   {
 *     type:      "SHIFT" | "OFF" | "VAC" | "RDO"
 *     shiftName: string  (e.g., "Morning") — only populated when type === "SHIFT"
 *     paidHours: number  — only populated when type === "SHIFT"
 *     locked:    boolean — true for VAC cells; no phase may overwrite a locked cell
 *   }
 *
 * INDEPENDENCE BETWEEN WEEKS:
 *   Each call to generateWeeklySchedule() produces a completely independent result.
 *   The grid is built from scratch from the Roster sheet each time. This means
 *   week 2 and week 3 of a 3-week generation run are not influenced by week 1's output.
 */


// ---------------------------------------------------------------------------
// Top-level entry point
// ---------------------------------------------------------------------------

/**
 * Generates a complete weekly schedule for the given Monday and returns the result.
 *
 * This function is an orchestrator — it calls the phase functions in order and
 * returns their combined output. It contains no scheduling logic itself.
 *
 * @param {Date} weekStartDate — The Monday of the week to generate (time portion is ignored).
 * @returns {{ weekSheetName: string, weekGrid: Array, employeeList: Array }}
 *   The generated grid, the list of employees (in seniority order), and the sheet name
 *   to use when writing the schedule. The caller (ui.js) is responsible for writing
 *   the grid to the sheet and calling the formatter.
 */
function generateWeeklySchedule(weekStartDate) {
  // Load settings once and pass them through all phases.
  // GAS sheet reads are expensive — loading both maps here prevents repeated reads
  // during the inner loops of Phases 1–3.
  const shiftTimingMap        = buildShiftTimingMap();
  const staffingRequirements  = loadStaffingRequirements();

  // --- Phase 0: Bootstrap ---
  const employeeList = loadRosterSortedBySeniority();

  if (employeeList.length === 0) {
    throw new Error(
      "The Roster sheet is empty. Sync the roster from the Ingestion sheet before generating a schedule."
    );
  }

  const weekGrid = initializeWeekGrid(employeeList, weekStartDate);

  // --- Phase 1: Preference Assignment ---
  runPhaseOnePreferenceAssignment(weekGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate);

  // --- Phase 2: Minimum Hour Enforcement ---
  runPhaseTwoHourEnforcement(weekGrid, employeeList, shiftTimingMap);

  // --- Phase 3: Gap Resolution ---
  runPhaseThreeGapResolution(weekGrid, employeeList, shiftTimingMap, staffingRequirements);

  return {
    weekSheetName: generateWeekSheetName(weekStartDate),
    weekGrid:      weekGrid,
    employeeList:  employeeList,
  };
}


// ---------------------------------------------------------------------------
// Phase 0: Bootstrap — Roster loading and grid initialization
// ---------------------------------------------------------------------------

/**
 * Reads all employee records from the Roster sheet and returns them sorted by seniority.
 *
 * Seniority order determines which employees get first pick of their preferred days off
 * and preferred shifts in Phase 1. The most senior employee is at index 0.
 *
 * Sort order:
 *   1. Descending seniority rank (higher number = more senior).
 *   2. If ranks are equal (same hire date AND same status), full-time before part-time.
 *   3. If still equal, alphabetical by name (deterministic tiebreak).
 *
 * @returns {Array<EmployeeRecord>} Employees in descending seniority order.
 */
function loadRosterSortedBySeniority() {
  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    throw new Error("Roster sheet not found. Run \"Setup Sheets\" from the Schedule Admin menu.");
  }

  const lastRow = rosterSheet.getLastRow();

  if (lastRow < ROSTER_DATA_START_ROW) {
    return [];
  }

  const dataRowCount = lastRow - ROSTER_DATA_START_ROW + 1;

  // Read the entire roster in a single call to minimize sheet read time.
  const allRosterValues = rosterSheet
    .getRange(ROSTER_DATA_START_ROW, 1, dataRowCount, ROSTER_COLUMN.SENIORITY_RANK)
    .getValues();

  const employeeList = [];

  allRosterValues.forEach(function(row) {
    const name        = row[ROSTER_COLUMN.NAME - 1];
    const employeeId  = row[ROSTER_COLUMN.EMPLOYEE_ID - 1];
    const hireDate    = row[ROSTER_COLUMN.HIRE_DATE - 1];
    const status      = row[ROSTER_COLUMN.STATUS - 1];

    // Skip blank rows that may appear at the bottom of a partially filled roster.
    if (!name || !employeeId) {
      return;
    }

    const dayOffPreferenceOne = row[ROSTER_COLUMN.DAY_OFF_PREF_ONE - 1];
    const dayOffPreferenceTwo = row[ROSTER_COLUMN.DAY_OFF_PREF_TWO - 1];
    const preferredShift      = row[ROSTER_COLUMN.PREFERRED_SHIFT - 1];
    const qualifiedShiftsRaw  = row[ROSTER_COLUMN.QUALIFIED_SHIFTS - 1];
    const vacationDatesRaw    = row[ROSTER_COLUMN.VACATION_DATES - 1];
    const seniorityRank       = row[ROSTER_COLUMN.SENIORITY_RANK - 1];

    // Parse the qualified shifts field from a comma-separated string into an array.
    // If the field is blank, default to using the preferred shift as the only qualified shift.
    const qualifiedShiftList = parseQualifiedShiftList(qualifiedShiftsRaw, preferredShift);

    // Vacation dates are stored as a comma-separated string in the Roster sheet.
    // They are parsed against the week's date range at grid initialization time,
    // not here, to keep this function focused on reading and transforming data.
    const vacationDateStrings = parseVacationDateStrings(vacationDatesRaw);

    employeeList.push({
      name:                 name.toString().trim(),
      employeeId:           employeeId.toString().trim(),
      hireDate:             hireDate instanceof Date ? hireDate : new Date(hireDate),
      status:               status.toString().trim(),
      dayOffPreferenceOne:  dayOffPreferenceOne ? dayOffPreferenceOne.toString().trim() : "",
      dayOffPreferenceTwo:  dayOffPreferenceTwo ? dayOffPreferenceTwo.toString().trim() : "",
      preferredShift:       preferredShift ? preferredShift.toString().trim() : "",
      qualifiedShifts:      qualifiedShiftList,
      vacationDateStrings:  vacationDateStrings,
      seniorityRank:        Number(seniorityRank) || 0,
    });
  });

  // Sort employees in descending seniority order so that the most senior employee
  // is processed first in all phases.
  employeeList.sort(compareEmployeesBySeniority);

  return employeeList;
}


/**
 * Creates the initial WeekGrid with all cells set to OFF, then stamps vacation locks.
 *
 * The grid is a 2D array indexed as weekGrid[employeeIndex][dayIndex] where dayIndex
 * 0 = Monday and dayIndex 6 = Sunday. Each cell is a DayAssignment object.
 *
 * Vacation locking happens here (Phase 0) rather than in Phase 1 so that vacation
 * cells are permanently protected before any scheduling logic runs. A locked cell
 * cannot be overwritten by Phase 1, 2, or 3 under any circumstances.
 *
 * @param {Array}  employeeList  — Employees in seniority order (from loadRosterSortedBySeniority).
 * @param {Date}   weekStartDate — The Monday of the week being generated.
 * @returns {Array} A 2D array of DayAssignment objects.
 */
function initializeWeekGrid(employeeList, weekStartDate) {
  const weekGrid = [];

  employeeList.forEach(function(employee, employeeIndex) {
    weekGrid[employeeIndex] = [];

    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      // All cells start as OFF. Phases 1–3 will convert these to SHIFT, RDO, or VAC as needed.
      weekGrid[employeeIndex][dayIndex] = createDayAssignment("OFF", null, 0, false);
    });

    // Stamp vacation locks for any of this employee's vacation dates that fall
    // within the current week. A locked VAC cell cannot be changed by any later phase.
    applyVacationLocksForEmployee(weekGrid, employeeIndex, employee, weekStartDate);
  });

  return weekGrid;
}


/**
 * Locks the grid cells for an employee's vacation days that fall within the given week.
 *
 * Vacation dates in the Roster sheet may span multiple weeks. This function checks
 * each vacation date string against the week's Monday–Sunday range and only locks
 * cells for dates that actually fall within this particular week. Dates outside
 * the week's range are silently ignored — they will be handled when that week is generated.
 *
 * @param {Array}  weekGrid       — The grid being initialized (mutated in place).
 * @param {number} employeeIndex  — The index of the employee in the grid.
 * @param {Object} employee       — The employee record (provides vacationDateStrings).
 * @param {Date}   weekStartDate  — The Monday of the week being generated.
 */
function applyVacationLocksForEmployee(weekGrid, employeeIndex, employee, weekStartDate) {
  employee.vacationDateStrings.forEach(function(vacationDateString) {
    const vacationDate = parseVacationDateString(vacationDateString, weekStartDate);

    if (!vacationDate) {
      // parseVacationDateString already logged the parse failure; skip this entry.
      return;
    }

    const dayIndex = getDayIndexForDate(vacationDate, weekStartDate);

    if (dayIndex === -1) {
      // This vacation date does not fall within the current week. It will be
      // handled when the relevant week is generated.
      return;
    }

    // Mark the cell as VAC and lock it so no phase can overwrite it.
    weekGrid[employeeIndex][dayIndex] = createDayAssignment("VAC", null, 0, true);
  });
}


// ---------------------------------------------------------------------------
// Phase 1: Preference Assignment
// ---------------------------------------------------------------------------

/**
 * Phase 1 orchestrator — grants RDO requests then assigns preferred shifts, in seniority order.
 *
 * Phase 1 is split into two sub-steps:
 *   1. grantRequestedDaysOff(): Decide which RDO requests to honor (seniority order, up to staffing floor).
 *   2. assignPreferredShifts(): Assign shifts to all non-VAC, non-RDO cells.
 *
 * These steps are separated so that RDO decisions are finalized before shift assignment begins.
 * If they ran together, a senior employee's RDO decision could change the coverage picture
 * mid-loop and affect a junior employee's shift assignment in an unpredictable way.
 *
 * @param {Array}  weekGrid              — The schedule grid (mutated in place).
 * @param {Array}  employeeList          — Employees in seniority order.
 * @param {Object} shiftTimingMap        — From buildShiftTimingMap().
 * @param {Object} staffingRequirements  — From loadStaffingRequirements().
 * @param {Date}   weekStartDate         — The Monday of the week being generated.
 */
function runPhaseOnePreferenceAssignment(weekGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate) {
  grantRequestedDaysOff(weekGrid, employeeList, staffingRequirements);
  assignPreferredShifts(weekGrid, employeeList, shiftTimingMap);
}


/**
 * Grants each employee's preferred days off in seniority order, up to the staffing floor.
 *
 * The staffing floor prevents too many employees from having the same day off.
 * For each day, the engine counts how many employees are still scheduled (not VAC or RDO),
 * and stops granting RDO requests once the count would drop below the minimum required staff.
 *
 * Employees whose RDO request cannot be honored (because the staffing floor would be breached)
 * are silently denied — their cell remains OFF, which Phase 1 will convert to a SHIFT assignment.
 * This is intentional: the schedule is a draft, and the manager can override any assignment manually.
 *
 * @param {Array}  weekGrid             — The schedule grid (mutated in place).
 * @param {Array}  employeeList         — Employees in descending seniority order.
 * @param {Object} staffingRequirements — Minimum staff per day of week.
 */
function grantRequestedDaysOff(weekGrid, employeeList, staffingRequirements) {
  // Process one day at a time so that the staffing floor check is accurate for each day independently.
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const minimumStaffRequired = staffingRequirements[dayName] || 0;

    // Count how many employees are available (not VAC) on this day.
    // VAC days are locked and cannot be counted as available staff.
    let availableStaffCount = 0;
    employeeList.forEach(function(employee, employeeIndex) {
      if (weekGrid[employeeIndex][dayIndex].type !== "VAC") {
        availableStaffCount++;
      }
    });

    // Grant RDO requests in seniority order (index 0 is most senior).
    // Stop granting once granting another would drop us below minimumStaffRequired.
    employeeList.forEach(function(employee, employeeIndex) {
      const currentCell = weekGrid[employeeIndex][dayIndex];

      // Never modify a locked VAC cell.
      if (currentCell.locked) {
        return;
      }

      const hasRequestedThisDayOff =
        employee.dayOffPreferenceOne === dayName ||
        employee.dayOffPreferenceTwo === dayName;

      if (!hasRequestedThisDayOff) {
        return;
      }

      // Only grant the RDO if we still have enough staff after removing this employee.
      // The check uses > rather than >= because we are checking if we can afford to
      // remove one more employee and still meet the minimum.
      if (availableStaffCount > minimumStaffRequired) {
        weekGrid[employeeIndex][dayIndex] = createDayAssignment("RDO", null, 0, false);
        // Decrement the available count to reflect that this employee is now off.
        availableStaffCount--;
      }
      // If the floor would be breached, do nothing — the cell stays as OFF and
      // this employee will be assigned a shift in the next step.
    });
  });
}


/**
 * Assigns a shift to every cell that is currently OFF (not VAC, RDO, or already SHIFT).
 *
 * For each working cell, the function looks up the employee's preferred shift in the
 * shift timing map. If the preferred shift is not found (e.g., it was deleted from Settings),
 * the function falls back to the first shift in the employee's qualified shift list.
 * If no valid shift is found at all, the cell is left as OFF and a warning is logged.
 *
 * @param {Array}  weekGrid       — The schedule grid (mutated in place).
 * @param {Array}  employeeList   — Employees in seniority order.
 * @param {Object} shiftTimingMap — From buildShiftTimingMap().
 */
function assignPreferredShifts(weekGrid, employeeList, shiftTimingMap) {
  employeeList.forEach(function(employee, employeeIndex) {
    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      const currentCell = weekGrid[employeeIndex][dayIndex];

      // Only assign shifts to cells that are currently OFF.
      // VAC and RDO cells are already finalized.
      if (currentCell.type !== "OFF") {
        return;
      }

      const shiftDefinition = resolveShiftForEmployee(employee, shiftTimingMap);

      if (!shiftDefinition) {
        // No valid shift found — leave as OFF and warn. The under-hours highlight
        // will flag this employee for the manager's attention.
        Logger.log(
          "WARNING: No valid shift found for employee \"" + employee.name + "\" on " +
          dayName + ". Their preferred shift (\"" + employee.preferredShift + "\") and all " +
          "qualified shifts are missing from the Settings sheet. This day will remain OFF."
        );
        return;
      }

      weekGrid[employeeIndex][dayIndex] = createDayAssignment(
        "SHIFT",
        shiftDefinition.name,
        shiftDefinition.paidHours,
        false,
        shiftDefinition.displayText
      );
    });
  });
}


// ---------------------------------------------------------------------------
// Phase 2: Minimum Hour Enforcement
// ---------------------------------------------------------------------------

/**
 * Ensures every employee meets their weekly paid hour minimum by converting OFF cells to shifts.
 *
 * After Phase 1, some employees may be below their minimum because they had many RDO grants
 * or because they work PT and their preferred shift only covers part of the week. Phase 2
 * converts remaining OFF cells to the employee's preferred shift until the minimum is met.
 *
 * Iteration order: employees are processed in order (most senior first). Within each
 * employee, days are processed Monday through Sunday. This deterministically fills days
 * from the start of the week, which is easier for managers to review.
 *
 * Constraints:
 *   - Never converts a VAC cell (locked).
 *   - Never converts an RDO cell — doing so would silently break a preference that was
 *     explicitly granted. If an employee is below minimum even after Phase 2 exhausts
 *     all OFF cells, the under-hours highlight in the formatter flags them.
 *   - Respects the weekly maximum — stops converting if adding another shift would
 *     push the employee over their status-based maximum.
 *
 * @param {Array}  weekGrid       — The schedule grid (mutated in place).
 * @param {Array}  employeeList   — Employees in seniority order.
 * @param {Object} shiftTimingMap — From buildShiftTimingMap().
 */
function runPhaseTwoHourEnforcement(weekGrid, employeeList, shiftTimingMap) {
  employeeList.forEach(function(employee, employeeIndex) {
    const weeklyMinimum = employee.status === "FT" ? HOUR_RULES.FT_MIN : HOUR_RULES.PT_MIN;
    const weeklyMaximum = employee.status === "FT" ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;

    let currentWeeklyHours = getWeeklyHours(weekGrid, employeeIndex);

    if (currentWeeklyHours >= weeklyMinimum) {
      // This employee already meets their minimum — nothing to do.
      return;
    }

    const shiftDefinition = resolveShiftForEmployee(employee, shiftTimingMap);

    if (!shiftDefinition) {
      // No valid shift available. The under-hours highlight will flag this employee.
      return;
    }

    // Scan through the week and convert OFF days to shifts until the minimum is met.
    DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
      // Re-check every iteration because the previous iteration may have added hours.
      if (currentWeeklyHours >= weeklyMinimum) {
        return;
      }

      const currentCell = weekGrid[employeeIndex][dayIndex];

      // Only convert OFF cells — VAC and RDO are intentional and must not be touched.
      if (currentCell.type !== "OFF") {
        return;
      }

      // Guard against exceeding the weekly maximum. This can happen if a partial
      // shift (e.g., a 5-hour PT shift) would push the employee past their cap.
      if (currentWeeklyHours + shiftDefinition.paidHours > weeklyMaximum) {
        return;
      }

      weekGrid[employeeIndex][dayIndex] = createDayAssignment(
        "SHIFT",
        shiftDefinition.name,
        shiftDefinition.paidHours,
        false,
        shiftDefinition.displayText
      );

      currentWeeklyHours += shiftDefinition.paidHours;
    });
  });
}


// ---------------------------------------------------------------------------
// Phase 3: Gap Resolution
// ---------------------------------------------------------------------------

/**
 * Detects uncovered 30-minute time slots for each day and fills them by reassigning
 * employees (Cascade A) or adding employees who are currently off (Cascade B).
 *
 * A "gap" is a 30-minute window between 04:00 and 23:30 where no employee is scheduled.
 * The engine first tries to cover gaps by shifting already-working employees to alternative
 * shifts that provide better coverage (Cascade A). If gaps remain, it pulls in employees
 * who are currently off (Cascade B).
 *
 * VAC cells are never touched. RDO cells are not touched by Cascade A. Cascade B may
 * convert OFF cells but not RDO cells — a granted day-off preference is respected even
 * in gap resolution.
 *
 * Junior employees are processed first in both cascades because it is fairer to adjust
 * the schedule of a junior employee than a senior one.
 *
 * @param {Array}  weekGrid             — The schedule grid (mutated in place).
 * @param {Array}  employeeList         — Employees in seniority order (index 0 = most senior).
 * @param {Object} shiftTimingMap       — From buildShiftTimingMap().
 * @param {Object} staffingRequirements — From loadStaffingRequirements() (used as a guard in Cascade B).
 */
function runPhaseThreeGapResolution(weekGrid, employeeList, shiftTimingMap, staffingRequirements) {
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    // Build the coverage map fresh for this day before starting either cascade.
    // The coverage map is a 39-element array where each element counts how many
    // employees are present during that 30-minute window.
    let dayCoverageSlots = buildDayCoverage(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);

    // Check whether there are any gaps to fill before running cascades.
    if (!hasCoverageGaps(dayCoverageSlots)) {
      return;
    }

    // --- Cascade A: Reassign working employees to alternative shifts ---
    // Process from most junior (last index) to most senior (index 0).
    for (let employeeIndex = employeeList.length - 1; employeeIndex >= 0; employeeIndex--) {
      const currentCell = weekGrid[employeeIndex][dayIndex];

      // Only consider employees who are currently scheduled to work.
      if (currentCell.type !== "SHIFT") {
        continue;
      }

      const employee = employeeList[employeeIndex];
      const weeklyMaximum = employee.status === "FT" ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;
      const currentWeeklyHours = getWeeklyHours(weekGrid, employeeIndex);

      // Compute coverage as if this employee were not scheduled, so we can fairly
      // score whether a different shift would provide better coverage.
      const coverageWithoutThisEmployee = buildDayCoverage(
        weekGrid, employeeList, dayIndex, shiftTimingMap, employeeIndex
      );

      // Find the best alternative shift from this employee's qualified shifts.
      const bestAlternativeShift = selectBestCoverageShift(
        employee.qualifiedShifts,
        employee.status,
        coverageWithoutThisEmployee,
        shiftTimingMap
      );

      if (!bestAlternativeShift) {
        continue;
      }

      // Only reassign if the alternative shift provides better gap coverage than the current shift.
      const currentShiftDefinition = shiftTimingMap[currentCell.shiftName + "|" + employee.status];
      const currentShiftScore  = currentShiftDefinition
        ? scoreCoverageForShift(currentShiftDefinition, coverageWithoutThisEmployee)
        : 0;
      const alternativeShiftScore = scoreCoverageForShift(bestAlternativeShift, coverageWithoutThisEmployee);

      if (alternativeShiftScore <= currentShiftScore) {
        // The alternative shift does not fill more gaps — do not reassign.
        continue;
      }

      // Check that switching to the alternative shift does not push the employee over their maximum.
      const hourDifference = bestAlternativeShift.paidHours - currentCell.paidHours;
      if (currentWeeklyHours + hourDifference > weeklyMaximum) {
        continue;
      }

      // Reassign this employee to the alternative shift.
      weekGrid[employeeIndex][dayIndex] = createDayAssignment(
        "SHIFT",
        bestAlternativeShift.name,
        bestAlternativeShift.paidHours,
        false,
        bestAlternativeShift.displayText
      );

      // Rebuild the coverage map to reflect this reassignment before evaluating the next employee.
      dayCoverageSlots = buildDayCoverage(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);

      if (!hasCoverageGaps(dayCoverageSlots)) {
        break; // All gaps filled by Cascade A — no need to continue.
      }
    }

    // Check again after Cascade A — Cascade B only runs if gaps remain.
    if (!hasCoverageGaps(dayCoverageSlots)) {
      return;
    }

    // --- Cascade B: Pull in employees who are currently OFF ---
    // Process from most junior to most senior, same as Cascade A.
    for (let employeeIndex = employeeList.length - 1; employeeIndex >= 0; employeeIndex--) {
      const currentCell = weekGrid[employeeIndex][dayIndex];

      // Only pull in employees who are currently OFF (not VAC, RDO, or already SHIFT).
      if (currentCell.type !== "OFF") {
        continue;
      }

      const employee = employeeList[employeeIndex];
      const weeklyMaximum = employee.status === "FT" ? HOUR_RULES.FT_MAX : HOUR_RULES.PT_MAX;
      const currentWeeklyHours = getWeeklyHours(weekGrid, employeeIndex);

      // Find the shift from this employee's qualified shifts that best fills the gaps.
      const bestGapFillingShift = selectBestCoverageShift(
        employee.qualifiedShifts,
        employee.status,
        dayCoverageSlots,
        shiftTimingMap
      );

      if (!bestGapFillingShift) {
        continue;
      }

      // Respect the weekly maximum — do not schedule an employee beyond their cap.
      if (currentWeeklyHours + bestGapFillingShift.paidHours > weeklyMaximum) {
        continue;
      }

      // Pull this employee in from OFF to cover the gap.
      weekGrid[employeeIndex][dayIndex] = createDayAssignment(
        "SHIFT",
        bestGapFillingShift.name,
        bestGapFillingShift.paidHours,
        false,
        bestGapFillingShift.displayText
      );

      // Rebuild the coverage map to reflect the addition.
      dayCoverageSlots = buildDayCoverage(weekGrid, employeeList, dayIndex, shiftTimingMap, -1);

      if (!hasCoverageGaps(dayCoverageSlots)) {
        break; // All gaps filled — stop pulling in employees.
      }
    }
  });
}


// ---------------------------------------------------------------------------
// Coverage Map Functions
// ---------------------------------------------------------------------------

/**
 * Builds a 39-element coverage array for a single day, representing staff presence
 * in 30-minute windows from 04:00 to 23:30.
 *
 * Each element of the returned array is a count of how many employees are physically
 * present during that 30-minute window. A value of 0 means a gap — no coverage.
 *
 * The coverage array is rebuilt from scratch each time it is needed rather than
 * maintained incrementally. This is simpler and avoids subtle bugs from partial updates,
 * at the cost of some extra computation. For typical roster sizes (< 50 employees)
 * this is not a performance concern.
 *
 * @param {Array}  weekGrid        — The current schedule grid.
 * @param {Array}  employeeList    — The full employee list.
 * @param {number} dayIndex        — The day column index (0 = Monday, 6 = Sunday).
 * @param {Object} shiftTimingMap  — From buildShiftTimingMap().
 * @param {number} excludeIndex    — Employee index to exclude from the count (-1 to include all).
 *   Passing an employee's index here computes coverage as if they were not scheduled,
 *   which is used in Cascade A to evaluate whether a shift change would help.
 * @returns {Array<number>} 39-element array of coverage counts per 30-minute slot.
 */
function buildDayCoverage(weekGrid, employeeList, dayIndex, shiftTimingMap, excludeIndex) {
  // Initialize all 39 slots to zero.
  const coverageSlots = new Array(COVERAGE.SLOT_COUNT).fill(0);

  employeeList.forEach(function(employee, employeeIndex) {
    // Exclude the specified employee if requested (used in Cascade A scoring).
    if (employeeIndex === excludeIndex) {
      return;
    }

    const cell = weekGrid[employeeIndex][dayIndex];

    // Only employees with SHIFT assignments contribute to coverage.
    // OFF, RDO, and VAC cells represent employees who are absent.
    if (cell.type !== "SHIFT") {
      return;
    }

    const shiftKey        = cell.shiftName + "|" + employee.status;
    const shiftDefinition = shiftTimingMap[shiftKey];

    if (!shiftDefinition) {
      // The shift was removed from Settings after the schedule was generated.
      // Log and skip — this employee will not contribute to coverage until
      // their shift is reassigned to a valid one.
      Logger.log(
        "WARNING: Shift key \"" + shiftKey + "\" not found in shift timing map during " +
        "coverage calculation. Employee: \"" + employee.name + "\"."
      );
      return;
    }

    // Calculate the slot range this shift covers.
    // The block time (start to end including unpaid lunch) determines physical presence,
    // not the paid hours. An employee on an 8.5-hour FT shift covers 17 slots.
    const startSlotIndex = Math.max(
      0,
      Math.floor((shiftDefinition.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES)
    );
    const endSlotIndex = Math.min(
      COVERAGE.SLOT_COUNT,
      Math.floor((shiftDefinition.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES)
    );

    // Increment the count for every slot this shift covers.
    // endSlotIndex is exclusive — the employee covers [startSlotIndex, endSlotIndex).
    for (let slotIndex = startSlotIndex; slotIndex < endSlotIndex; slotIndex++) {
      coverageSlots[slotIndex]++;
    }
  });

  return coverageSlots;
}


/**
 * Returns true if any slot in the coverage array has zero employees.
 *
 * A zero-count slot represents a 30-minute window where no employee is scheduled,
 * which is the definition of a coverage gap that Phase 3 attempts to fill.
 *
 * @param {Array<number>} coverageSlots — 39-element coverage array from buildDayCoverage().
 * @returns {boolean}
 */
function hasCoverageGaps(coverageSlots) {
  return coverageSlots.some(function(slotCount) {
    return slotCount === 0;
  });
}


/**
 * Given a list of shift names and the current coverage state, returns the ShiftDefinition
 * that would most effectively fill uncovered time slots.
 *
 * Shift scoring: for each slot a shift covers, its score is incremented by
 * 1 / (existingCoverage + 1). This formula gives higher scores to shifts that
 * cover slots with zero or low coverage, and lower scores to shifts that cover
 * already-well-covered periods. The +1 prevents division by zero when coverage is 0.
 *
 * @param {Array<string>} qualifiedShiftNames — Shift names from the employee's qualified shifts list.
 * @param {string}        employmentStatus    — "FT" or "PT", used to look up the correct shift variant.
 * @param {Array<number>} coverageSlots       — Current 39-element coverage array.
 * @param {Object}        shiftTimingMap      — From buildShiftTimingMap().
 * @returns {Object|null} The ShiftDefinition with the highest coverage score, or null if none found.
 */
function selectBestCoverageShift(qualifiedShiftNames, employmentStatus, coverageSlots, shiftTimingMap) {
  let highestScore          = -1;
  let bestShiftDefinition   = null;

  qualifiedShiftNames.forEach(function(shiftName) {
    const shiftKey        = shiftName.trim() + "|" + employmentStatus;
    const shiftDefinition = shiftTimingMap[shiftKey];

    if (!shiftDefinition) {
      // This shift name is in the employee's qualified list but not in Settings.
      // This is a data quality issue — log it and skip.
      Logger.log(
        "WARNING: Qualified shift \"" + shiftKey + "\" not found in shift timing map. " +
        "Check that this shift name is defined in the Settings sheet for status " + employmentStatus + "."
      );
      return;
    }

    const score = scoreCoverageForShift(shiftDefinition, coverageSlots);

    if (score > highestScore) {
      highestScore        = score;
      bestShiftDefinition = shiftDefinition;
    }
  });

  return bestShiftDefinition;
}


/**
 * Scores a shift definition against the current coverage map.
 *
 * Higher scores indicate the shift would fill more uncovered (or under-covered) slots.
 * The scoring formula 1 / (coverage + 1) is used because:
 *   - A slot with 0 coverage contributes 1.0 to the score (highest possible per slot).
 *   - A slot with 1 existing employee contributes 0.5 (half as valuable).
 *   - A slot with 2 employees contributes 0.33, and so on.
 * This means the function naturally prioritizes filling zero-coverage gaps first,
 * then slots with the least coverage, without any explicit conditional logic.
 *
 * @param {Object}        shiftDefinition — A ShiftDefinition object from the shift timing map.
 * @param {Array<number>} coverageSlots   — The current 39-element coverage array.
 * @returns {number} The coverage score for this shift.
 */
function scoreCoverageForShift(shiftDefinition, coverageSlots) {
  const startSlotIndex = Math.max(
    0,
    Math.floor((shiftDefinition.startMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES)
  );
  const endSlotIndex = Math.min(
    COVERAGE.SLOT_COUNT,
    Math.floor((shiftDefinition.endMinutes - COVERAGE.COVERAGE_START_MINUTE) / COVERAGE.SLOT_DURATION_MINUTES)
  );

  let totalScore = 0;

  for (let slotIndex = startSlotIndex; slotIndex < endSlotIndex; slotIndex++) {
    // Add 1 to existing coverage to avoid division by zero, and to reduce the
    // marginal value of covering already-covered slots.
    totalScore += 1 / (coverageSlots[slotIndex] + 1);
  }

  return totalScore;
}


// ---------------------------------------------------------------------------
// Seniority and Hour Utilities
// ---------------------------------------------------------------------------

/**
 * Calculates the seniority rank integer for a given employee.
 *
 * The rank is a single number that encodes both employment status and length of service:
 *   - Full-time employees receive a base of 200,000,000.
 *   - Part-time employees receive a base of 100,000,000.
 *   - The number of days between the hire date and a future reference date (2050-01-01)
 *     is added to the base.
 *
 * This means earlier hire dates produce larger day values (they are further from 2050),
 * so more senior employees naturally receive higher ranks without any conditional logic.
 * The 100,000,000 gap between FT and PT bases is large enough that no real hire date
 * can close it, ensuring FT employees always outrank PT employees hired on the same date.
 *
 * @param {Date}   hireDate         — The employee's hire date.
 * @param {string} employmentStatus — "FT" or "PT".
 * @returns {number} The seniority rank integer.
 */
function calculateSeniorityRank(hireDate, employmentStatus) {
  const statusBase = employmentStatus === "FT" ? SENIORITY.FT_BASE : SENIORITY.PT_BASE;

  // The reference date is a future anchor point. Subtracting the hire date from this
  // reference produces a larger integer for employees hired earlier (longer ago).
  const referenceDate = new Date(SENIORITY.REFERENCE_DATE_STRING);

  const validHireDate = hireDate instanceof Date && !isNaN(hireDate.getTime())
    ? hireDate
    : new Date(); // Fallback to today if the hire date is invalid.

  // Divide by 86,400,000 to convert milliseconds to days.
  const daysFromHireToReference = Math.floor(
    (referenceDate.getTime() - validHireDate.getTime()) / 86400000
  );

  return statusBase + daysFromHireToReference;
}


/**
 * Recalculates and writes seniority ranks for all employees in the Roster sheet.
 *
 * This function is called after every roster sync and whenever a manager changes
 * an employee's status (FT/PT). It reads each row's hire date and status, computes
 * the seniority rank, and writes it back to column J.
 */
function refreshAllSeniorityRanks() {
  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    return;
  }

  const lastRow = rosterSheet.getLastRow();

  if (lastRow < ROSTER_DATA_START_ROW) {
    return;
  }

  const dataRowCount = lastRow - ROSTER_DATA_START_ROW + 1;

  // Read hire dates and statuses in a single batch to minimize sheet reads.
  const hireDateValues = rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.HIRE_DATE, dataRowCount, 1)
    .getValues();

  const statusValues = rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.STATUS, dataRowCount, 1)
    .getValues();

  // Build an array of new seniority ranks to write back in a single batch.
  const newSeniorityRankValues = hireDateValues.map(function(hireDateRow, rowOffset) {
    const hireDate         = hireDateRow[0];
    const employmentStatus = statusValues[rowOffset][0];
    return [calculateSeniorityRank(hireDate, employmentStatus ? employmentStatus.toString() : "PT")];
  });

  // Write all ranks in one call — much faster than writing cell-by-cell.
  rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.SENIORITY_RANK, dataRowCount, 1)
    .setValues(newSeniorityRankValues);
}


/**
 * Recalculates and writes the seniority rank for a single Roster row.
 *
 * Called by onEdit() when a manager changes the status of a single employee.
 * Recalculating only the affected row is faster than running refreshAllSeniorityRanks()
 * for every individual edit.
 *
 * @param {number} rowNumber — The 1-indexed row number in the Roster sheet.
 */
function recalculateSeniorityRankForRow(rowNumber) {
  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    return;
  }

  const hireDate         = rosterSheet.getRange(rowNumber, ROSTER_COLUMN.HIRE_DATE).getValue();
  const employmentStatus = rosterSheet.getRange(rowNumber, ROSTER_COLUMN.STATUS).getValue();

  const newRank = calculateSeniorityRank(hireDate, employmentStatus ? employmentStatus.toString() : "PT");

  rosterSheet.getRange(rowNumber, ROSTER_COLUMN.SENIORITY_RANK).setValue(newRank);
}


/**
 * Sums the paid hours for all SHIFT cells in one employee's week.
 *
 * This is a pure computation — it reads from the in-memory grid and does not
 * touch any sheet. It is called frequently (in every phase) and must be fast.
 *
 * @param {Array}  weekGrid      — The current schedule grid.
 * @param {number} employeeIndex — The index of the employee to sum.
 * @returns {number} Total paid hours scheduled for the employee this week.
 */
function getWeeklyHours(weekGrid, employeeIndex) {
  let totalHours = 0;

  weekGrid[employeeIndex].forEach(function(cell) {
    if (cell.type === "SHIFT") {
      totalHours += cell.paidHours;
    }
  });

  return totalHours;
}


// ---------------------------------------------------------------------------
// Date and Parsing Utilities
// ---------------------------------------------------------------------------

/**
 * Generates the sheet tab name for a generated schedule, e.g., "Week_04_07_26".
 *
 * The name format "Week_MM_DD_YY" was chosen because:
 *   1. It sorts chronologically in the Sheets tab bar (tabs are ordered by insertion,
 *      but a consistent naming pattern makes them easy to scan and locate).
 *   2. The underscores make it easy to parse programmatically if needed.
 *   3. The two-digit year keeps the name short enough to read in the tab bar.
 *
 * @param {Date} weekStartDate — The Monday of the week.
 * @returns {string} The sheet name, e.g., "Week_04_07_26".
 */
function generateWeekSheetName(weekStartDate) {
  const month = String(weekStartDate.getMonth() + 1).padStart(2, "0");
  const day   = String(weekStartDate.getDate()).padStart(2, "0");
  const year  = String(weekStartDate.getFullYear()).slice(-2);
  return "Week_" + month + "_" + day + "_" + year;
}


/**
 * Returns the date (as a Date object) for a given day of the week relative to the week start.
 *
 * @param {Date}   weekStartDate — The Monday of the week (day index 0).
 * @param {number} dayIndex      — 0 = Monday, 6 = Sunday.
 * @returns {Date} The calendar date for that column.
 */
function getDateForDayIndex(weekStartDate, dayIndex) {
  const result = new Date(weekStartDate);
  result.setDate(weekStartDate.getDate() + dayIndex);
  return result;
}


/**
 * Returns the day index (0 = Monday, 6 = Sunday) for a given date relative to the week start.
 * Returns -1 if the date does not fall within the week (Monday through Sunday).
 *
 * @param {Date} targetDate    — The date to check.
 * @param {Date} weekStartDate — The Monday of the week.
 * @returns {number} Day index 0–6, or -1 if out of range.
 */
function getDayIndexForDate(targetDate, weekStartDate) {
  const weekStartTime = new Date(weekStartDate);
  weekStartTime.setHours(0, 0, 0, 0);

  const targetTime = new Date(targetDate);
  targetTime.setHours(0, 0, 0, 0);

  const differenceInDays = Math.round(
    (targetTime.getTime() - weekStartTime.getTime()) / 86400000
  );

  if (differenceInDays < 0 || differenceInDays > 6) {
    return -1;
  }

  return differenceInDays;
}


/**
 * Parses a vacation date string into a Date object anchored to the year of the given week.
 *
 * Accepted formats:
 *   "YYYY-MM-DD" — ISO format (unambiguous; preferred)
 *   "MM/DD"      — Month/day shorthand (year is inferred from the week start date's year)
 *   "M/D"        — Single-digit month/day shorthand
 *
 * If parsing fails for all formats, logs a warning and returns null so the caller
 * can skip this entry rather than crash.
 *
 * @param {string} dateString    — The vacation date string from the Roster sheet.
 * @param {Date}   weekStartDate — Used to infer the year for "MM/DD" format strings.
 * @returns {Date|null} The parsed Date, or null if parsing failed.
 */
function parseVacationDateString(dateString, weekStartDate) {
  const trimmedString = dateString.toString().trim();

  if (!trimmedString) {
    return null;
  }

  // Try ISO format "YYYY-MM-DD" first — unambiguous and highest priority.
  if (/^\d{4}-\d{2}-\d{2}$/.test(trimmedString)) {
    const parsed = new Date(trimmedString + "T00:00:00");
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  // Try "MM/DD" or "M/D" shorthand — infer the year from the week start date.
  if (/^\d{1,2}\/\d{1,2}$/.test(trimmedString)) {
    const parts = trimmedString.split("/");
    const month = parseInt(parts[0], 10);
    const day   = parseInt(parts[1], 10);
    const year  = weekStartDate.getFullYear();

    const parsed = new Date(year, month - 1, day);
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  Logger.log(
    "WARNING: Could not parse vacation date string \"" + dateString + "\". " +
    "Use YYYY-MM-DD (e.g., 2026-04-07) or MM/DD (e.g., 4/7) format. " +
    "This date will be skipped."
  );
  return null;
}


/**
 * Parses the raw vacation dates cell value into an array of trimmed date strings.
 *
 * The Roster sheet stores vacation dates as a comma-separated string in one cell.
 * This function splits that string, trims whitespace, and filters out blank entries.
 *
 * @param {*} rawCellValue — The raw value from the Roster vacation dates cell.
 * @returns {Array<string>} An array of trimmed date strings.
 */
function parseVacationDateStrings(rawCellValue) {
  if (!rawCellValue || rawCellValue.toString().trim() === "") {
    return [];
  }

  return rawCellValue.toString()
    .split(",")
    .map(function(entry) { return entry.trim(); })
    .filter(function(entry) { return entry !== ""; });
}


/**
 * Parses the raw qualified shifts cell value into an array of trimmed shift name strings.
 *
 * If the cell is blank, defaults to an array containing only the employee's preferred shift
 * so that at minimum the engine always has one valid shift to work with.
 *
 * @param {*}      rawCellValue    — The raw value from the Roster qualified shifts cell.
 * @param {string} preferredShift  — Fallback shift name if the cell is blank.
 * @returns {Array<string>} An array of trimmed shift name strings.
 */
function parseQualifiedShiftList(rawCellValue, preferredShift) {
  if (!rawCellValue || rawCellValue.toString().trim() === "") {
    return preferredShift ? [preferredShift.toString().trim()] : [];
  }

  const parsedList = rawCellValue.toString()
    .split(",")
    .map(function(entry) { return entry.trim(); })
    .filter(function(entry) { return entry !== ""; });

  // Ensure the preferred shift is always in the qualified list so it can be
  // used as a fallback even if the manager forgot to include it explicitly.
  if (preferredShift && !parsedList.includes(preferredShift.toString().trim())) {
    parsedList.unshift(preferredShift.toString().trim());
  }

  return parsedList;
}


// ---------------------------------------------------------------------------
// Comparator and Factory Functions
// ---------------------------------------------------------------------------

/**
 * Comparator function for sorting employees in descending seniority order.
 *
 * Sort logic:
 *   1. Higher seniority rank first (descending).
 *   2. If ranks are equal: FT before PT (redundant with the rank formula for same
 *      hire date, but included as an explicit guard for edge cases).
 *   3. If still equal: alphabetical by name (A before Z), providing a deterministic
 *      and consistent tiebreak that managers can predict.
 *
 * @param {Object} employeeA — First employee record.
 * @param {Object} employeeB — Second employee record.
 * @returns {number} Negative if A should come first, positive if B should come first.
 */
function compareEmployeesBySeniority(employeeA, employeeB) {
  // Primary sort: descending seniority rank.
  if (employeeB.seniorityRank !== employeeA.seniorityRank) {
    return employeeB.seniorityRank - employeeA.seniorityRank;
  }

  // Secondary sort: FT before PT at equal seniority.
  if (employeeA.status !== employeeB.status) {
    return employeeA.status === "FT" ? -1 : 1;
  }

  // Tertiary sort: alphabetical by name.
  return employeeA.name.localeCompare(employeeB.name);
}


/**
 * Creates a DayAssignment object — the building block of the WeekGrid.
 *
 * Using a factory function (rather than constructing object literals inline) ensures
 * every cell in the grid has the same shape, which prevents undefined property errors
 * when phase code checks cell.type or cell.paidHours.
 *
 * @param {string}  assignmentType — "SHIFT", "OFF", "VAC", or "RDO".
 * @param {string|null} shiftName  — The shift name, e.g., "Morning". Null for non-SHIFT types.
 * @param {number}  paidHours      — Paid hours for this shift. 0 for non-SHIFT types.
 * @param {boolean} isLocked       — True for VAC cells; prevents any phase from overwriting.
 * @returns {Object} A DayAssignment object.
 */
function createDayAssignment(assignmentType, shiftName, paidHours, isLocked, displayText) {
  return {
    type:      assignmentType,
    shiftName: shiftName,
    paidHours: paidHours,
    locked:    isLocked,
    displayText: displayText || null,
    // display text is the human readable time range written into the employees schedule cell,
    // for example "8:00 AM - 4:30 PM", only populated for cells with a valid shift
  };
}


/**
 * Resolves the best available shift definition for an employee.
 *
 * Tries the employee's preferred shift first. If it is not in the shift timing map
 * for their status (e.g., the shift was renamed in Settings), falls back to the
 * first shift in their qualified shift list that does exist.
 *
 * @param {Object} employee       — The employee record.
 * @param {Object} shiftTimingMap — From buildShiftTimingMap().
 * @returns {Object|null} A ShiftDefinition, or null if no valid shift is found.
 */
function resolveShiftForEmployee(employee, shiftTimingMap) {
  // Try the preferred shift first.
  const preferredShiftKey = employee.preferredShift + "|" + employee.status;
  if (shiftTimingMap[preferredShiftKey]) {
    return shiftTimingMap[preferredShiftKey];
  }

  // Fall back to the first qualified shift that exists in the map.
  for (let shiftIndex = 0; shiftIndex < employee.qualifiedShifts.length; shiftIndex++) {
    const qualifiedShiftKey = employee.qualifiedShifts[shiftIndex] + "|" + employee.status;
    if (shiftTimingMap[qualifiedShiftKey]) {
      return shiftTimingMap[qualifiedShiftKey];
    }
  }

  return null;
}


/**
 * Reads the VAC and RDO checkbox state from an existing Week sheet and returns a
 * partial grid that represents the manager's manual decisions.
 *
 * This function is called by resolveEntireWeek() in ui.js when a manager checks or
 * unchecks a VAC or RDO checkbox. The partial grid is used to reinitialize the week
 * grid with the manager's decisions locked in before re-running Phases 1–3.
 *
 * @param {Sheet} weekSheet — The Week_MM_DD_YY sheet object.
 * @param {number} employeeCount — The number of employees on the Roster.
 * @returns {Array} A 2D array of partial DayAssignment objects reflecting checkbox state.
 */
function readCheckboxStateFromSheet(weekSheet, employeeCount) {
  const checkboxState = [];

  for (let employeeIndex = 0; employeeIndex < employeeCount; employeeIndex++) {
    checkboxState[employeeIndex] = [];

    const baseRow = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const vacationRow = baseRow + WEEK_SHEET.ROW_OFFSET_VAC;
    const requestedDayOffRow = baseRow + WEEK_SHEET.ROW_OFFSET_RDO;

    // Read all 7 day columns for VAC and RDO rows in two batch reads.
    const vacationValues = weekSheet
      .getRange(vacationRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .getValues()[0];

    const requestedDayOffValues = weekSheet
      .getRange(requestedDayOffRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .getValues()[0];

    for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
      const isVacation        = vacationValues[dayIndex] === true;
      const isRequestedDayOff = requestedDayOffValues[dayIndex] === true;

      if (isVacation) {
        checkboxState[employeeIndex][dayIndex] = createDayAssignment("VAC", null, 0, true);
      } else if (isRequestedDayOff) {
        checkboxState[employeeIndex][dayIndex] = createDayAssignment("RDO", null, 0, false);
      } else {
        checkboxState[employeeIndex][dayIndex] = createDayAssignment("OFF", null, 0, false);
      }
    }
  }

  return checkboxState;
}
