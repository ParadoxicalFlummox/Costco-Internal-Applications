/**
 * formatter.js — Writes a generated WeekGrid to a Google Sheet and applies all visual formatting.
 * VERSION: 0.5.4
 *
 * This file is the only place in the codebase that writes to a Week schedule sheet.
 * The schedule engine (scheduleEngine.js) produces a pure JavaScript data structure (the WeekGrid).
 * This file translates that data structure into what the manager actually sees on screen.
 *
 * PERF: Optimized to batch write operations—all SHIFT row values across all employees
 * are written in one API call, not per-employee. Shift colors and role colors are
 * similarly batched via setBackgrounds() for each row type.
 *
 * SEPARATION OF CONCERNS:
 * Every visual concern is its own function:
 *   writeWeekHeader()               — Rows 1–4 (title, timestamp, department)
 *   writeColumnHeaders()            — Row 5 (Mon, Tue, ... column labels)
 *   writeEmployeeBlocks()           — Employee VAC/RDO/SHIFT rows
 *   writeStaffingSummary()          — REQUIRED/ACTUAL/STATUS footer
 *   applyShiftColors()              — Cell background colors by assignment type
 *   applyUnderHoursHighlight()      — Red name cell for employees below minimum hours
 *   applyStatusRowConditionalFormat() — Green/red STATUS row cells
 *   applyStructuralFormatting()     — Borders, column widths, freeze rows
 *
 * If a visual bug occurs, it can be traced to exactly one function.
 *
 * RE-GENERATION BEHAVIOR:
 * When a manager checks a VAC or RDO checkbox on an existing schedule sheet, the engine
 * re-runs Phases 1–3 and then this formatter re-writes the sheet. The formatter clears
 * only the SHIFT row cells — the VAC and RDO checkboxes set by the manager are preserved.
 */


/**
 * The single entry point for writing and formatting a schedule sheet.
 *
 * This function is an orchestrator — it calls all the write and format functions in the
 * correct order and passes their shared inputs. It contains no formatting logic itself.
 *
 * POOL SECTION (v0.5.0):
 * If poolMemberIds is provided (from traffic heatmap system), pool members are written
 * in their own section above regular employees. Pool members are visually distinguished
 * with a background color and "POOL" label. Pool section is always first (rows after header),
 * followed by the regular employee section.
 *
 * @param {Sheet}  scheduleSheet       — The Week_MM_DD_YY sheet to write to.
 * @param {Array}  employeeList        — Employees in seniority order (from scheduleEngine.js).
 * @param {Array}  weekGrid            — The generated schedule grid (from scheduleEngine.js).
 * @param {Object} staffingRequirements — From loadStaffingRequirements().
 * @param {Date}   weekStartDate        — The Monday of the week being written.
 * @param {string} departmentName       — The department name to display in the sheet header.
 * @param {Set}    poolMemberIds       — (Optional) Set of employee IDs who are pool members (from traffic heatmap).
 */
function writeAndFormatSchedule(scheduleSheet, employeeList, weekGrid, staffingRequirements, weekStartDate, departmentName, poolMemberIds) {
  // Partition employees into pool and regular based on poolMemberIds
  let poolMembers = [];
  let regularEmployees = [];

  if (poolMemberIds && poolMemberIds.size > 0) {
    employeeList.forEach(function(emp) {
      if (poolMemberIds.has(emp.id)) {
        poolMembers.push(emp);
      } else {
        regularEmployees.push(emp);
      }
    });
  } else {
    // No pool section: all employees are regular
    regularEmployees = employeeList.slice();
  }

  // Write all content first, then apply formatting.
  // Interleaving content writes and formatting calls would slow down rendering
  // because GAS batches API calls — writing all values first then formatting is faster.

  logExecutionTime_('Write Week Header', function() {
    writeWeekHeader(scheduleSheet, weekStartDate, departmentName);
  });

  logExecutionTime_('Write Column Headers', function() {
    writeColumnHeaders(scheduleSheet);
  });

  // Write summary rows BEFORE employee blocks so they occupy fixed rows 6/7/8.
  // Employee blocks grow downward from row 9; hybrid-pass appends land below
  // without ever disturbing the summary row positions.
  logExecutionTime_('Write Staffing Summary', function() {
    writeStaffingSummary(scheduleSheet, employeeList, weekGrid, staffingRequirements, poolMembers.length);
  });

  logExecutionTime_('Write Pool Section (' + poolMembers.length + ' pool members)', function() {
    if (poolMembers.length > 0) {
      writePoolSection_(scheduleSheet, poolMembers, weekGrid, employeeList);
    }
  });

  logExecutionTime_('Write Employee Blocks (' + regularEmployees.length + ' regular employees)', function() {
    writeEmployeeBlocks(scheduleSheet, regularEmployees, weekGrid, poolMembers.length);
  });

  // Flush all pending content writes before starting formatting.
  // This ensures GAS does not hold too many deferred operations in memory,
  // which reduces mid-run timeout risk on large rosters or multi-department runs.
  SpreadsheetApp.flush();

  // Apply visual formatting after all content is written.
  logExecutionTime_('Apply Shift Colors', function() {
    applyShiftColors(scheduleSheet, employeeList, weekGrid);
  });

  logExecutionTime_('Apply Role Row Colors', function() {
    applyRoleRowColors(scheduleSheet, employeeList, weekGrid);
  });

  logExecutionTime_('Apply Under-Hours Highlight', function() {
    applyUnderHoursHighlight(scheduleSheet, regularEmployees, weekGrid, poolMembers.length);
  });

  logExecutionTime_('Apply Status Row Conditional Format', function() {
    applyStatusRowConditionalFormat(scheduleSheet);
  });

  logExecutionTime_('Apply Structural Formatting', function() {
    applyStructuralFormatting(scheduleSheet, employeeList.length);
  });
}


// ---------------------------------------------------------------------------
// Content Writers
// ---------------------------------------------------------------------------

/**
 * Writes the schedule sheet header rows (rows 1–4).
 *
 * The header contains:
 *   Row 1: "Week of [Month Day] – [Day], [Year]" (merged across all columns)
 *   Row 2: "Generated: [timestamp]"
 *   Row 3: "Department: [departmentName]"
 *   Row 4: (blank spacer row)
 *
 * @param {Sheet}  scheduleSheet  — The schedule sheet to write to.
 * @param {Date}   weekStartDate  — The Monday of the week.
 * @param {string} departmentName — The department name from the Ingestion sheet.
 */
function writeWeekHeader(scheduleSheet, weekStartDate, departmentName) {
  const weekEndDate = getDateForDayIndex(weekStartDate, 6); // Sunday

  // GAS's V8 Intl implementation does not produce clean output for partial date
  // option sets like { day, year } without month — it renders "(day: 12) 2026".
  // Build the label manually: "April 6 – 12, 2026".
  const weekLabel =
    "Week of " +
    weekStartDate.toLocaleDateString("en-US", { month: "long", day: "numeric" }) +
    " \u2013 " +
    weekEndDate.getDate() + ", " + weekEndDate.getFullYear();

  // Row 1: Week label — merged across all 10 columns for visual impact.
  const titleRange = scheduleSheet.getRange(
    WEEK_SHEET.HEADER_ROW, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS
  );
  titleRange.merge();
  titleRange.setValue(weekLabel);
  titleRange.setFontSize(14);
  titleRange.setFontWeight("bold");
  titleRange.setBackground(COLORS.HEADER_BG);
  titleRange.setFontColor(COLORS.HEADER_TEXT);
  titleRange.setHorizontalAlignment("center");
  titleRange.setVerticalAlignment("middle");

  // Row 2: Generation timestamp — formatted for human readability.
  // Format: "Generated: April 20, 2026 at 3:45 PM"
  const now = new Date();
  const timestampText = "Generated: " +
    now.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" }) +
    " at " +
    now.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });

  const timestampRange = scheduleSheet.getRange(WEEK_SHEET.TIMESTAMP_ROW, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS);
  timestampRange.merge();
  timestampRange.setValue(timestampText);
  timestampRange.setFontSize(10);
  timestampRange.setFontColor("#666666");
  timestampRange.setHorizontalAlignment("left");

  // Row 3: Department name — styled for consistency.
  const deptText = "Department: " + departmentName;
  const deptRange = scheduleSheet.getRange(WEEK_SHEET.DEPARTMENT_ROW, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS);
  deptRange.merge();
  deptRange.setValue(deptText);
  deptRange.setFontSize(11);
  deptRange.setFontWeight("bold");
  deptRange.setFontColor("#000000");
  deptRange.setHorizontalAlignment("left");

  // Row 4: Spacer — light background for visual separation (between department and column headers).
  const spacerRange = scheduleSheet.getRange(WEEK_SHEET.COLUMN_HEADER_ROW - 1, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS);
  spacerRange.setBackground("#EEEEEE");
}


/**
 * Writes the column header row (row 5) with day names and "Total Hrs".
 *
 * This row is frozen so that it remains visible when the manager scrolls down
 * through a long roster. The freeze is applied in applyStructuralFormatting().
 *
 * @param {Sheet} scheduleSheet — The schedule sheet to write to.
 */
function writeColumnHeaders(scheduleSheet) {
  const headerRowValues = [
    ["Label", "Employee", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun", "Total Hrs"]
  ];

  scheduleSheet
    .getRange(WEEK_SHEET.COLUMN_HEADER_ROW, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS)
    .setValues(headerRowValues);

  // Style the header row to stand out visually from the employee data below it.
  const headerRange = scheduleSheet.getRange(
    WEEK_SHEET.COLUMN_HEADER_ROW, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS
  );
  headerRange.setBackground(COLORS.HEADER_BG);
  headerRange.setFontColor(COLORS.HEADER_TEXT);
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
}


/**
 * Writes pool member blocks to the schedule sheet with "POOL" section label.
 * Pool members appear first (above regular employees) in the final sheet.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet to write to.
 * @param {Array} poolMembers   — Pool members in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid (includes all employees).
 * @param {Array} allEmployees  — Full employee list (for grid indexing).
 */
function writePoolSection_(scheduleSheet, poolMembers, weekGrid, allEmployees) {
  const poolStartRow = WEEK_SHEET.DATA_START_ROW;

  poolMembers.forEach(function(employee, poolIndex) {
    const baseRow            = poolStartRow + (poolIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const vacationRow        = baseRow + WEEK_SHEET.ROW_OFFSET_VAC;
    const requestedDayOffRow = baseRow + WEEK_SHEET.ROW_OFFSET_RDO;
    const shiftRow           = baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT;
    const roleRow            = baseRow + WEEK_SHEET.ROW_OFFSET_ROLE;
    const lockRow            = baseRow + WEEK_SHEET.ROW_OFFSET_LOCK;

    // Write row labels with "POOL" prefix
    scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_LABEL).setValue('POOL-VAC');
    scheduleSheet.getRange(requestedDayOffRow, WEEK_SHEET.COL_LABEL).setValue('POOL-RDO');
    scheduleSheet.getRange(shiftRow, WEEK_SHEET.COL_LABEL).setValue('POOL-SHIFT');
    scheduleSheet.getRange(roleRow, WEEK_SHEET.COL_LABEL).setValue('POOL-ROLE');
    scheduleSheet.getRange(lockRow, WEEK_SHEET.COL_LABEL).setValue('POOL-LOCK');

    // Merge and write employee name across all five rows
    scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_EMPLOYEE_NAME, WEEK_SHEET.ROWS_PER_EMPLOYEE, 1)
      .merge()
      .setValue(employee.name)
      .setVerticalAlignment('middle');

    // Write VAC/RDO checkboxes (empty on first generation)
    scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).insertCheckboxes();
    scheduleSheet.getRange(requestedDayOffRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).insertCheckboxes();

    // Write SHIFT and ROLE rows from grid
    const gridRowIndex = allEmployees.indexOf(employee);
    if (gridRowIndex >= 0) {
      const shiftRowValues = [];
      const roleRowValues = [];
      for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
        const cell = weekGrid[gridRowIndex][dayIndex];
        shiftRowValues.push([cell.displayText || '']);
        roleRowValues.push([cell.role && cell.type === 'SHIFT' ? cell.role : '—']);
      }
      scheduleSheet.getRange(shiftRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).setValues([shiftRowValues.map(v => v[0])]);
      scheduleSheet.getRange(roleRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).setValues([roleRowValues.map(v => v[0])]);
    }

    // Write LOCK checkboxes (hidden)
    scheduleSheet.getRange(lockRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK).insertCheckboxes();

    // Apply pool section background color to the entire block (5 rows)
    scheduleSheet.getRange(baseRow, 1, WEEK_SHEET.ROWS_PER_EMPLOYEE, WEEK_SHEET.COL_TOTAL_HOURS)
      .setBackground(COLORS.POOL_SECTION_BG); // config.js
  });
}


/**
 * Writes all employee data blocks (VAC row, RDO row, SHIFT row, ROLE row, LOCK row) to the schedule sheet.
 *
 * Each employee occupies five consecutive rows:
 *   Row 1 of block (VAC):   "VAC" label | employee name | checkboxes for Mon–Sun
 *   Row 2 of block (RDO):   "RDO" label | (name merged from VAC row) | checkboxes for Mon–Sun
 *   Row 3 of block (SHIFT): "SHIFT" label | (name merged) | shift text for Mon–Sun | total hours
 *   Row 4 of block (ROLE):  "ROLE" label | (name merged) | role name for working days, "—" otherwise
 *   Row 5 of block (LOCK):  "LOCK" label | (name merged) | hidden lock checkboxes for Mon–Sun
 *
 * The employee name cell is merged across all five rows in the block and vertically centered.
 * This makes it visually clear which rows belong to one employee.
 *
 * RE-GENERATION NOTE: On re-generation (when a manager edits a checkbox), this function
 * writes only the SHIFT and ROLE row values. The VAC, RDO, and LOCK checkboxes are not touched
 * because they represent the manager's explicit decisions. The checkboxes are only cleared
 * and re-inserted on the first generation of a new week sheet.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet to write to.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 * @param {number} poolRowOffset — (Optional) Number of pool member rows before regular employees.
 */
function writeEmployeeBlocks(scheduleSheet, employeeList, weekGrid, poolRowOffset) {
  poolRowOffset = poolRowOffset || 0;
  // Determine if this is a first-time write or a re-generation.
  // Check by looking for content in the first REGULAR employee's name cell (accounting for pool offset).
  const firstEmployeeNameCell = scheduleSheet.getRange(
    WEEK_SHEET.DATA_START_ROW + (poolRowOffset * WEEK_SHEET.ROWS_PER_EMPLOYEE) + WEEK_SHEET.ROW_OFFSET_VAC,
    WEEK_SHEET.COL_EMPLOYEE_NAME
  );
  const isFirstTimeGeneration = firstEmployeeNameCell.getValue() === "";

  // Collect total hours per employee during the loop; write as one batched call after.
  // This replaces 57 individual setValue() calls with a single setValues() call,
  // reducing write time from ~70s to <1s for large departments.
  const totalHoursCollected = new Array(employeeList.length).fill(0);

  // On first-time generation, collect lock row numbers to batch setRowHeight after loop.
  const lockRowsToHide = [];

  employeeList.forEach(function(employee, employeeIndex) {
    const baseRow            = WEEK_SHEET.DATA_START_ROW + (poolRowOffset * WEEK_SHEET.ROWS_PER_EMPLOYEE) + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const vacationRow        = baseRow + WEEK_SHEET.ROW_OFFSET_VAC;
    const requestedDayOffRow = baseRow + WEEK_SHEET.ROW_OFFSET_RDO;
    const shiftRow           = baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT;
    const roleRow            = baseRow + WEEK_SHEET.ROW_OFFSET_ROLE;
    const lockRow            = baseRow + WEEK_SHEET.ROW_OFFSET_LOCK;

    if (isFirstTimeGeneration) {
      // --- First-time generation: write all five rows from scratch ---

      // Row label column (A): "VAC", "RDO", "SHIFT", "ROLE", "LOCK" — one batch write for all five labels.
      scheduleSheet
        .getRange(vacationRow, WEEK_SHEET.COL_ROW_LABEL, WEEK_SHEET.ROWS_PER_EMPLOYEE, 1)
        .setValues([["VAC"], ["RDO"], ["SHIFT"], ["ROLE"], ["LOCK"]]);

      // Employee name cell (B): merged across all 5 rows, vertically centered.
      // WEEK_SHEET.ROWS_PER_EMPLOYEE is 5, so this merges the correct block height.
      const nameMergeRange = scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_EMPLOYEE_NAME, WEEK_SHEET.ROWS_PER_EMPLOYEE, 1);
      nameMergeRange.merge();
      nameMergeRange.setValue(employee.name);
      nameMergeRange.setVerticalAlignment("middle");
      nameMergeRange.setFontWeight("bold");

      // VAC row (C–I): insert all checkboxes in one call, then set all values in one batch.
      // insertCheckboxes() on a range applies to every cell in that range at once.
      const vacDayRange = scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK);
      vacDayRange.insertCheckboxes();
      vacDayRange.setValues([weekGrid[employeeIndex].map(function(cell) {
        return cell.type === "VAC";
      })]);

      // RDO row (C–I): same pattern — one insertCheckboxes call + one setValues call.
      const rdoDayRange = scheduleSheet.getRange(requestedDayOffRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK);
      rdoDayRange.insertCheckboxes();
      rdoDayRange.setValues([weekGrid[employeeIndex].map(function(cell) {
        return cell.type === "RDO";
      })]);

      // LOCK row (C–I): hidden checkboxes indicating manager cell overrides.
      // Initially all false (no locks on fresh generation). These are populated when
      // updateCellOverride() is called.
      const lockDayRange = scheduleSheet.getRange(lockRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK);
      lockDayRange.insertCheckboxes();
      lockDayRange.setValues([weekGrid[employeeIndex].map(function(_cell) {
        return false; // Fresh generation has no locks
      })]);

      // Collect lock rows to hide after the loop (batched to avoid per-employee API calls).
      lockRowsToHide.push(lockRow);
    }


    // --- SHIFT row (C–I): always written (first time or re-generation) ---
    // Build the 7-day values array in one pass, then write the whole row in one API call.
    let totalPaidHoursThisWeek = 0;
    const shiftRowValues = [weekGrid[employeeIndex].map(function(cell) {
      if (cell.type === "SHIFT") {
        totalPaidHoursThisWeek += cell.paidHours;
        return cell.displayText || cell.shiftName || "SHIFT";
      } else if (cell.type === "VAC") {
        return "VAC";
      } else if (cell.type === "RDO") {
        return "RDO";
      } else {
        return "OFF";
      }
    })];

    scheduleSheet
      .getRange(shiftRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .setValues(shiftRowValues);

    // Collect total hours for batch write after the loop (not written here).
    totalHoursCollected[employeeIndex] = totalPaidHoursThisWeek;

    // --- ROLE row (C–I): always written (roles change when shifts change) ---
    // Build the 7-day role values in one pass, write the whole row in one API call.
    const roleRowValues = [weekGrid[employeeIndex].map(function(cell) {
      // Show the employee's role on working days; em dash on all off/non-working days.
      return (cell.type === "SHIFT" && cell.role) ? cell.role : "\u2014";
    })];

    scheduleSheet
      .getRange(roleRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .setValues(roleRowValues);
  });

  // --- Batch write: total hours column (COL_TOTAL_HOURS) ---
  // Build a values array covering all employee rows (ROWS_PER_EMPLOYEE rows each).
  // Only the SHIFT row offset gets a value; other row offsets get an empty string.
  // This replaces N individual setValue() calls with one setValues() call.
  if (employeeList.length > 0) {
    const totalHoursBlockStart = WEEK_SHEET.DATA_START_ROW + (poolRowOffset * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const totalHoursBlockRows  = employeeList.length * WEEK_SHEET.ROWS_PER_EMPLOYEE;
    const totalHoursBlockValues = [];
    for (let employeeIndex = 0; employeeIndex < employeeList.length; employeeIndex++) {
      for (let rowOffset = 0; rowOffset < WEEK_SHEET.ROWS_PER_EMPLOYEE; rowOffset++) {
        totalHoursBlockValues.push(
          rowOffset === WEEK_SHEET.ROW_OFFSET_SHIFT
            ? [totalHoursCollected[employeeIndex]]
            : ['']
        );
      }
    }
    scheduleSheet
      .getRange(totalHoursBlockStart, WEEK_SHEET.COL_TOTAL_HOURS, totalHoursBlockRows, 1)
      .setValues(totalHoursBlockValues);
  }

  // --- Batch hide: lock row heights (first-time generation only) ---
  // setRowHeight() has no multi-row API, so we iterate here after all other writes
  // are done — keeping the per-employee loop free of single-cell API calls.
  lockRowsToHide.forEach(function(lockRow) {
    scheduleSheet.setRowHeight(lockRow, 1);
  });
}


/**
 * Writes the staffing summary block (REQUIRED / ACTUAL / STATUS rows) below the employee blocks.
 *
 * Supports two modes per day, driven by the staffing requirements Settings column C:
 *
 *   COUNT mode (default):
 *     REQUIRED — minimum employee count for that day.
 *     ACTUAL   — live COUNTIF formula counting shift cells (cells containing ":").
 *
 *   HOURS mode:
 *     REQUIRED — minimum total paid hours for that day (e.g., 40).
 *     ACTUAL   — computed sum of paidHours for all working employees on that day,
 *                written as a static value (recalculated on every re-generation).
 *
 * STATUS row: "OK" if actual >= required, "UNDER" otherwise. Both modes use the
 * same formula-driven STATUS row — the cell references work regardless of whether
 * the ACTUAL value is a formula result or a static number.
 *
 * @param {Sheet}  scheduleSheet        — The schedule sheet to write to.
 * @param {Array}  employeeList         — All employees (includes both pool and regular).
 * @param {Array}  weekGrid             — The generated schedule grid.
 * @param {Object} staffingRequirements — { dayName → { value, mode } } from loadStaffingRequirements().
 * @param {number} poolRowOffset        — (Optional) Number of pool member rows before regular employees.
 */
function writeStaffingSummary(scheduleSheet, employeeList, weekGrid, staffingRequirements, poolRowOffset) {
  // poolRowOffset is kept for signature compatibility but no longer used — summary rows
  // are at fixed positions and do not depend on employee count or pool section size.
  void poolRowOffset;

  // Fixed row positions — never move regardless of employee count.
  // Rows 6/7/8 are always REQUIRED/ACTUAL/STATUS; employee data starts at row 9.
  const requiredRow = WEEK_SHEET.SUMMARY_REQUIRED_ROW;
  const actualRow   = WEEK_SHEET.SUMMARY_ACTUAL_ROW;
  const statusRow   = WEEK_SHEET.SUMMARY_STATUS_ROW;

  // Row labels (column A).
  scheduleSheet.getRange(requiredRow, WEEK_SHEET.COL_ROW_LABEL).setValue("REQUIRED").setFontWeight("bold");
  scheduleSheet.getRange(actualRow,   WEEK_SHEET.COL_ROW_LABEL).setValue("ACTUAL").setFontWeight("bold");
  scheduleSheet.getRange(statusRow,   WEEK_SHEET.COL_ROW_LABEL).setValue("STATUS").setFontWeight("bold");

  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const columnNumber    = WEEK_SHEET.COL_MONDAY + dayIndex;
    const columnLetter    = columnIndexToLetter(columnNumber);
    const dayRequirement  = staffingRequirements[dayName] || { value: 0, mode: STAFFING_MODE.COUNT };
    const isHoursMode     = dayRequirement.mode === STAFFING_MODE.HOURS;

    // REQUIRED: target value — interpreted as head count or hours depending on mode.
    scheduleSheet.getRange(requiredRow, columnNumber).setValue(dayRequirement.value);

    if (isHoursMode) {
      // ACTUAL (Hours mode): sum paid hours for all employees working this day.
      // Written as a static number because paidHours live in the grid, not cells.
      // This value is refreshed on every generation/re-generation run.
      let totalHoursThisDay = 0;
      employeeList.forEach(function(_employee, employeeIndex) {
        const cell = weekGrid[employeeIndex][dayIndex];
        if (cell.type === "SHIFT") {
          totalHoursThisDay += cell.paidHours || 0;
        }
      });
      scheduleSheet.getRange(actualRow, columnNumber).setValue(totalHoursThisDay);
    } else {
      // ACTUAL (Count mode): live formula counting cells that contain a time-range
      // string (identified by the colon in "8:45 AM - 9:30 AM").
      // Role names, OFF/VAC/RDO text, and checkboxes are all excluded by "*:*".
      // Open-ended range (DATA_START_ROW:10000) automatically includes any hybrid
      // employee rows appended by a future second-pass run.
      const countIfFormula =
        "=COUNTIF(" + columnLetter + WEEK_SHEET.DATA_START_ROW + ":" + columnLetter + "10000,\"*:*\")";
      scheduleSheet.getRange(actualRow, columnNumber).setFormula(countIfFormula);
    }

    // STATUS: formula comparing actual to required — works for both modes.
    const actualCellAddress   = columnLetter + actualRow;
    const requiredCellAddress = columnLetter + requiredRow;
    const statusFormula =
      "=IF(" + actualCellAddress + ">=" + requiredCellAddress + ",\"OK\",\"UNDER\")";
    scheduleSheet.getRange(statusRow, columnNumber).setFormula(statusFormula);
  });

  // Department total hours (STATUS row, column J) — sum of all weekly per-employee totals.
  // Open-ended range automatically includes hybrid employee rows appended later.
  const totalHoursColumnLetter = columnIndexToLetter(WEEK_SHEET.COL_TOTAL_HOURS);
  const departmentTotalFormula =
    "=SUM(" + totalHoursColumnLetter + WEEK_SHEET.DATA_START_ROW + ":" + totalHoursColumnLetter + "10000)";
  scheduleSheet.getRange(statusRow, WEEK_SHEET.COL_TOTAL_HOURS)
    .setFormula(departmentTotalFormula)
    .setFontWeight("bold");
}


// ---------------------------------------------------------------------------
// Formatting Functions
// ---------------------------------------------------------------------------

/**
 * Applies cell background colors to SHIFT row cells based on the assignment type and status.
 *
 * Color mapping:
 *   Combo SHIFT → COLORS.COMBO_SHIFT (deep orange) — shift name contains "Combo"
 *   FT SHIFT    → COLORS.FT_SHIFT    (blue)
 *   PT SHIFT    → COLORS.PT_SHIFT    (green)
 *   VAC         → COLORS.VACATION    (yellow)
 *   RDO / OFF   → COLORS.DAY_OFF     (gray)
 *
 * Only SHIFT row cells receive color. VAC and RDO rows use a neutral background so the
 * shift row visually pops as the primary information row for each employee.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 */
function applyShiftColors(scheduleSheet, employeeList, weekGrid) {
  // Build a 7-value color row per employee, then write all 7 backgrounds in one API call.
  // This reduces 7 setBackground() calls per employee down to 1 setBackgrounds() call.
  employeeList.forEach(function(employee, employeeIndex) {
    const shiftRow = WEEK_SHEET.DATA_START_ROW +
      (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE) +
      WEEK_SHEET.ROW_OFFSET_SHIFT;

    const rowColors = [weekGrid[employeeIndex].map(function(cell) {
      if (cell.type === "SHIFT") {
        // Combo shifts (cross-dept handoff) get orange regardless of FT/PT status.
        // Blue for FT, green for PT/LPT otherwise.
        if (cell.shiftName && cell.shiftName.indexOf('Combo') !== -1) {
          return COLORS.COMBO_SHIFT;
        }
        return employee.status === "FT" ? COLORS.FT_SHIFT : COLORS.PT_SHIFT;
      } else if (cell.type === "VAC") {
        return COLORS.VACATION;
      } else {
        // RDO and OFF cells both use the day-off color.
        return COLORS.DAY_OFF;
      }
    })];

    scheduleSheet
      .getRange(shiftRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .setBackgrounds(rowColors);
  });
}


/**
 * Applies per-role background colors to ROLE row cells.
 *
 * Each role name (Cashier, SCO, PreScan, etc.) maps to a distinct pastel background
 * defined in the ROLE_COLORS lookup in config.js. This color coding lets supervisors
 * scan across a day column at a glance to see the role mix without reading every cell.
 *
 * Cells for non-working days (OFF, RDO, VAC) receive the generic ROLE_ROW_BG color
 * (lavender) so the row has a consistent base background even when empty.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 */
function applyRoleRowColors(scheduleSheet, employeeList, weekGrid) {
  // Build a 7-value color row per employee, then write all 7 backgrounds in one API call.
  employeeList.forEach(function(_employee, employeeIndex) {
    const roleRow = WEEK_SHEET.DATA_START_ROW +
      (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE) +
      WEEK_SHEET.ROW_OFFSET_ROLE;

    const rowColors = [weekGrid[employeeIndex].map(function(cell) {
      if (cell.type === "SHIFT" && cell.role && ROLE_COLORS[cell.role]) {
        return ROLE_COLORS[cell.role];
      }
      // Non-working days and roles not in the lookup both get the generic row color.
      return COLORS.ROLE_ROW_BG;
    })];

    scheduleSheet
      .getRange(roleRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .setBackgrounds(rowColors);
  });
}


/**
 * Highlights the employee name cell to flag hours violations:
 *   - Red  (UNDER_HOURS)  — employee is below their weekly minimum.
 *   - Orange (OVER_HOURS_FT) — FT employee is above 40 hours (overtime risk).
 *
 * Both flags are informational — they prompt the manager to review without blocking generation.
 * Common under-hours causes: too many vacation days, no valid shifts in Settings.
 * Common over-hours cause: gap resolution pulled in an already-full FT employee.
 *
 * The highlight is applied to the name cell (column B) on the SHIFT row, which is
 * the most visible row and ensures the highlight is not obscured by merged cells.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 */
function applyUnderHoursHighlight(scheduleSheet, employeeList, weekGrid, poolRowOffset) {
  poolRowOffset = poolRowOffset || 0;

  const totalEmployeeRows = employeeList.length * WEEK_SHEET.ROWS_PER_EMPLOYEE;
  if (totalEmployeeRows === 0) return;

  // Sheet row where this employee list's blocks begin (after any pool section).
  const blockStartSheetRow = WEEK_SHEET.DATA_START_ROW
    + (poolRowOffset * WEEK_SHEET.ROWS_PER_EMPLOYEE);

  // Build two sparse color arrays — one for column B (name), one for column J (total hours).
  // Each array covers every row in the employee block range; null = clear existing background.
  // Only the SHIFT row within each 5-row block receives a color; all others get null.
  // This reduces 4N individual getRange/setBackground calls to 2 total API calls.
  const nameColors  = [];
  const hoursColors = [];

  for (let employeeIndex = 0; employeeIndex < employeeList.length; employeeIndex++) {
    const employee    = employeeList[employeeIndex];
    const weeklyHours = getWeeklyHours(weekGrid, employeeIndex);
    const weeklyMin   = employee.status === 'FT'  ? HOUR_RULES.FT_MIN
                      : employee.status === 'LPT' ? HOUR_RULES.LPT_MIN
                      : HOUR_RULES.PT_MIN;

    let nameColor  = null;
    let hoursColor = null;

    if (employee.status === 'FT' && weeklyHours > HOUR_RULES.FT_MAX) {
      // Orange on the total hours cell flags overtime risk for FT employees.
      hoursColor = COLORS.OVER_HOURS_FT;
    } else if (weeklyHours < weeklyMin) {
      // Red on the name cell flags under-minimum hours.
      nameColor = COLORS.UNDER_HOURS;
    }
    // null on both clears any prior highlight when re-generating (employee now in-hours).

    for (let rowOffset = 0; rowOffset < WEEK_SHEET.ROWS_PER_EMPLOYEE; rowOffset++) {
      const isShiftRow = rowOffset === WEEK_SHEET.ROW_OFFSET_SHIFT;
      nameColors.push([isShiftRow  ? nameColor  : null]);
      hoursColors.push([isShiftRow ? hoursColor : null]);
    }
  }

  // Two API calls total regardless of roster size.
  scheduleSheet
    .getRange(blockStartSheetRow, WEEK_SHEET.COL_EMPLOYEE_NAME, totalEmployeeRows, 1)
    .setBackgrounds(nameColors);
  scheduleSheet
    .getRange(blockStartSheetRow, WEEK_SHEET.COL_TOTAL_HOURS, totalEmployeeRows, 1)
    .setBackgrounds(hoursColors);
}


/**
 * Applies conditional formatting to the STATUS row so that "OK" cells are green
 * and "UNDER" cells are red.
 *
 * IMPORTANT: This function calls clearConditionalFormatRules() before adding new rules.
 * Without this, conditional formatting rules accumulate on every re-generation, which
 * eventually causes GAS to apply the wrong rules or silently fail.
 *
 * @param {Sheet}  scheduleSheet  — The schedule sheet.
 * @param {number} employeeCount  — Used to calculate the STATUS row number.
 */
function applyStatusRowConditionalFormat(scheduleSheet) {
  const statusRow = WEEK_SHEET.SUMMARY_STATUS_ROW; // Fixed position — always row 8

  const statusRange = scheduleSheet.getRange(
    statusRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK
  );

  // Clear all existing conditional format rules on this sheet before adding new ones.
  // Accumulated rules from repeated generations can cause incorrect behavior.
  scheduleSheet.clearConditionalFormatRules();

  const okRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("OK")
    .setBackground(COLORS.SUMMARY_OK)
    .setFontColor("#FFFFFF")
    .setRanges([statusRange])
    .build();

  const underRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("UNDER")
    .setBackground(COLORS.SUMMARY_UNDER)
    .setFontColor("#FFFFFF")
    .setRanges([statusRange])
    .build();

  scheduleSheet.setConditionalFormatRules([okRule, underRule]);
}


/**
 * Applies structural formatting: column widths, row heights, frozen rows/columns, borders.
 *
 * This function is called last because it modifies the sheet's structural properties
 * rather than cell values, and it relies on the content already being written so that
 * row height calculations are accurate.
 *
 * @param {Sheet}  scheduleSheet  — The schedule sheet.
 * @param {number} employeeCount  — Number of employee blocks, used to calculate border range.
 */
function applyStructuralFormatting(scheduleSheet, employeeCount) {
  // --- Column widths ---
  scheduleSheet.setColumnWidth(WEEK_SHEET.COL_ROW_LABEL,      60);  // "VAC" / "RDO" / "SHIFT"
  scheduleSheet.setColumnWidth(WEEK_SHEET.COL_EMPLOYEE_NAME,  175); // Employee name
  // Day columns: Mon–Sun
  for (let dayColumn = WEEK_SHEET.COL_MONDAY; dayColumn <= WEEK_SHEET.COL_SUNDAY; dayColumn++) {
    scheduleSheet.setColumnWidth(dayColumn, 100);
  }
  scheduleSheet.setColumnWidth(WEEK_SHEET.COL_TOTAL_HOURS, 85); // "Total Hrs"

  // --- Freeze the header rows ---
  // Freezing rows 1–8 keeps the week label, staffing summary (REQUIRED/ACTUAL/STATUS),
  // and column headers always visible while scrolling through employee blocks.
  // Column A is NOT frozen because the week header row (row 1) has a merged cell spanning
  // all columns (A1:J1). Google Sheets does not allow freezing a column that would split
  // a merged cell — attempting to do so throws a runtime error.
  scheduleSheet.setFrozenRows(WEEK_SHEET.SUMMARY_STATUS_ROW);

  // --- Row label column (A) background ---
  // Light gray background on the label column helps visually separate it from data cells.
  const totalDataRows = employeeCount * WEEK_SHEET.ROWS_PER_EMPLOYEE;
  if (totalDataRows > 0) {
    scheduleSheet
      .getRange(WEEK_SHEET.DATA_START_ROW, WEEK_SHEET.COL_ROW_LABEL, totalDataRows, 1)
      .setBackground(COLORS.ROW_LABEL_BG)
      .setHorizontalAlignment("center")
      .setFontStyle("italic");

    // Override the ROLE label cell with lavender so it matches the ROLE data cells.
    for (let employeeIndex = 0; employeeIndex < employeeCount; employeeIndex++) {
      const roleRow = WEEK_SHEET.DATA_START_ROW +
        (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE) +
        WEEK_SHEET.ROW_OFFSET_ROLE;
      scheduleSheet
        .getRange(roleRow, WEEK_SHEET.COL_ROW_LABEL)
        .setBackground(COLORS.ROLE_ROW_BG);

      // Apply italic + center to the ROLE data cells (Mon–Sun).
      scheduleSheet
        .getRange(roleRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
        .setFontStyle("italic")
        .setHorizontalAlignment("center");
    }
  }

  // --- Borders between employee blocks ---
  // A top border on the first row of each employee block provides a clear visual
  // separator between employees, making the three-row structure easy to scan.
  for (let employeeIndex = 0; employeeIndex < employeeCount; employeeIndex++) {
    const blockStartRow = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const blockRange    = scheduleSheet.getRange(
      blockStartRow, 1, 1, WEEK_SHEET.COL_TOTAL_HOURS
    );
    blockRange.setBorder(
      true,  // top
      null, null, null, null, null,
      "#CCCCCC",
      SpreadsheetApp.BorderStyle.SOLID
    );
  }

  // --- Center-align all day columns and Total Hrs column ---
  const dataAndSummaryRowCount = (employeeCount * WEEK_SHEET.ROWS_PER_EMPLOYEE) + 5;
  scheduleSheet
    .getRange(WEEK_SHEET.DATA_START_ROW, WEEK_SHEET.COL_MONDAY, dataAndSummaryRowCount, WEEK_SHEET.DAYS_IN_WEEK + 1)
    .setHorizontalAlignment("center");

  // --- Auto fit only the day columns and total hours columns ---
  scheduleSheet.autoResizeColumns(WEEK_SHEET.COL_MONDAY, WEEK_SHEET.DAYS_IN_WEEK + 1);
}


// ---------------------------------------------------------------------------
// Multi-Department Writer
// ---------------------------------------------------------------------------

/**
 * Generates and formats one schedule sheet per department from a multi-department run.
 *
 * This is the multi-department counterpart to the single-department writeAndFormatSchedule()
 * call. It loops the Map returned by generateAllDepartmentSchedules() and delegates each
 * department's data to writeAndFormatSchedule(), which handles all content writing and
 * visual formatting for that sheet.
 *
 * Sheet names follow the pattern: Week_MM_DD_YY_DeptName
 * (e.g., "Week_04_07_26_Morning", "Week_04_07_26_Drivers")
 *
 * @param {Map}  allDeptResults — Map<deptName → { weekGrid, employeeList, staffingRequirements }>
 *                                Returned by generateAllDepartmentSchedules().
 * @param {Date} weekStartDate  — The Monday of the week being written.
 */
function writeAllDepartmentSchedules_(allDeptResults, weekStartDate) {
  allDeptResults.forEach(function(deptResult, deptKey) {
    const sheetName     = generateDeptWeekSheetName(weekStartDate, deptKey);
    const scheduleSheet = getOrCreateWeekSheet(sheetName);
    writeAndFormatSchedule(
      scheduleSheet,
      deptResult.employeeList,
      deptResult.weekGrid,
      deptResult.staffingRequirements,
      weekStartDate,
      deptResult.displayName || deptKey  // show original name in the header, not the normalized key
    );
  });
}


// ---------------------------------------------------------------------------
// Sheet Management Helpers
// ---------------------------------------------------------------------------

/**
 * Creates a new schedule sheet with the given name, or clears and returns the existing one.
 *
 * On first generation: a new sheet is inserted.
 * On re-generation (sheet already exists): the SHIFT row cells are cleared but the
 * VAC and RDO checkbox values are preserved because they represent manager decisions.
 *
 * @param {string} weekSheetName — The name for the week sheet (e.g., "Week_04_07_26").
 * @returns {Sheet} The sheet object to write to.
 */
function getOrCreateWeekSheet(weekSheetName) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  let existingSheet = workbook.getSheetByName(weekSheetName);

  if (existingSheet) {
    // Sheet already exists — clear only SHIFT row cells (every 3rd row starting at row offset 2).
    // The VAC and RDO rows are intentionally preserved because they hold manager decisions.
    return existingSheet;
  }

  // Insert a new sheet at the end of the workbook.
  return workbook.insertSheet(weekSheetName);
}


// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

/**
 * Converts a 1-indexed column number to its spreadsheet letter notation.
 *
 * For example: 1 → "A", 3 → "C", 26 → "Z", 27 → "AA".
 * This is needed for building formula strings (e.g., COUNTIF range addresses).
 *
 * @param {number} columnNumber — 1-indexed column number.
 * @returns {string} The column letter(s), e.g., "C" or "AB".
 */
function columnIndexToLetter(columnNumber) {
  let letter = "";
  let remaining = columnNumber;

  while (remaining > 0) {
    const remainder  = (remaining - 1) % 26;
    letter    = String.fromCharCode(65 + remainder) + letter;
    remaining = Math.floor((remaining - 1) / 26);
  }

  return letter;
}
