/**
 * formatter.js — Writes a generated WeekGrid to a Google Sheet and applies all visual formatting.
 * VERSION: 0.3.2
 *
 * This file is the only place in the codebase that writes to a Week schedule sheet.
 * The schedule engine (scheduleEngine.js) produces a pure JavaScript data structure (the WeekGrid).
 * This file translates that data structure into what the manager actually sees on screen.
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
 * @param {Sheet}  scheduleSheet       — The Week_MM_DD_YY sheet to write to.
 * @param {Array}  employeeList        — Employees in seniority order (from scheduleEngine.js).
 * @param {Array}  weekGrid            — The generated schedule grid (from scheduleEngine.js).
 * @param {Object} staffingRequirements — From loadStaffingRequirements().
 * @param {Date}   weekStartDate        — The Monday of the week being written.
 * @param {string} departmentName       — The department name to display in the sheet header.
 */
function writeAndFormatSchedule(scheduleSheet, employeeList, weekGrid, staffingRequirements, weekStartDate, departmentName) {
  // Write all content first, then apply formatting.
  // Interleaving content writes and formatting calls would slow down rendering
  // because GAS batches API calls — writing all values first then formatting is faster.

  writeWeekHeader(scheduleSheet, weekStartDate, departmentName);
  writeColumnHeaders(scheduleSheet);
  writeEmployeeBlocks(scheduleSheet, employeeList, weekGrid);
  writeStaffingSummary(scheduleSheet, employeeList.length, staffingRequirements);

  // Apply visual formatting after all content is written.
  applyShiftColors(scheduleSheet, employeeList, weekGrid);
  applyUnderHoursHighlight(scheduleSheet, employeeList, weekGrid);
  applyStatusRowConditionalFormat(scheduleSheet, employeeList.length);
  applyStructuralFormatting(scheduleSheet, employeeList.length);
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

  // Row 2: Generation timestamp — helps managers identify the most recent draft.
  scheduleSheet
    .getRange(WEEK_SHEET.TIMESTAMP_ROW, 1)
    .setValue("Generated: " + new Date().toLocaleString());

  // Row 3: Department name.
  scheduleSheet
    .getRange(WEEK_SHEET.DEPARTMENT_ROW, 1)
    .setValue("Department: " + departmentName);
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
 * Writes all employee data blocks (VAC row, RDO row, SHIFT row) to the schedule sheet.
 *
 * Each employee occupies three consecutive rows:
 *   Row 1 of block (VAC):   "VAC" label | employee name | checkboxes for Mon–Sun
 *   Row 2 of block (RDO):   "RDO" label | (name merged from VAC row) | checkboxes for Mon–Sun
 *   Row 3 of block (SHIFT): "SHIFT" label | (name merged) | shift text for Mon–Sun | total hours
 *
 * The employee name cell is merged across all three rows in the block and vertically centered.
 * This makes it visually clear which three rows belong to one employee.
 *
 * RE-GENERATION NOTE: On re-generation (when a manager edits a checkbox), this function
 * writes only the SHIFT row values. The VAC and RDO checkboxes are not touched because
 * they represent the manager's explicit decisions. The checkboxes are only cleared and
 * re-inserted on the first generation of a new week sheet.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet to write to.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 */
function writeEmployeeBlocks(scheduleSheet, employeeList, weekGrid) {
  // Determine if this is a first-time write or a re-generation.
  // Check by looking for content in the first employee's name cell.
  const firstEmployeeNameCell = scheduleSheet.getRange(
    WEEK_SHEET.DATA_START_ROW + WEEK_SHEET.ROW_OFFSET_VAC, WEEK_SHEET.COL_EMPLOYEE_NAME
  );
  const isFirstTimeGeneration = firstEmployeeNameCell.getValue() === "";

  employeeList.forEach(function(employee, employeeIndex) {
    const baseRow            = WEEK_SHEET.DATA_START_ROW + (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE);
    const vacationRow        = baseRow + WEEK_SHEET.ROW_OFFSET_VAC;
    const requestedDayOffRow = baseRow + WEEK_SHEET.ROW_OFFSET_RDO;
    const shiftRow           = baseRow + WEEK_SHEET.ROW_OFFSET_SHIFT;

    if (isFirstTimeGeneration) {
      // --- First-time generation: write all three rows from scratch ---

      // Row label column (A): "VAC", "RDO", "SHIFT"
      scheduleSheet.getRange(vacationRow,        WEEK_SHEET.COL_ROW_LABEL).setValue("VAC");
      scheduleSheet.getRange(requestedDayOffRow, WEEK_SHEET.COL_ROW_LABEL).setValue("RDO");
      scheduleSheet.getRange(shiftRow,           WEEK_SHEET.COL_ROW_LABEL).setValue("SHIFT");

      // Employee name cell (B): merged across all 3 rows, vertically centered.
      const nameMergeRange = scheduleSheet.getRange(vacationRow, WEEK_SHEET.COL_EMPLOYEE_NAME, WEEK_SHEET.ROWS_PER_EMPLOYEE, 1);
      nameMergeRange.merge();
      nameMergeRange.setValue(employee.name);
      nameMergeRange.setVerticalAlignment("middle");
      nameMergeRange.setFontWeight("bold");

      // VAC row (C–I): insert checkboxes, pre-checked for vacation days.
      for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
        const columnNumber = WEEK_SHEET.COL_MONDAY + dayIndex;
        const vacationCell = scheduleSheet.getRange(vacationRow, columnNumber);
        vacationCell.insertCheckboxes();
        // Pre-check the checkbox if this is a vacation day in the generated grid.
        if (weekGrid[employeeIndex][dayIndex].type === "VAC") {
          vacationCell.setValue(true);
        }
      }

      // RDO row (C–I): insert checkboxes, pre-checked for RDO days.
      for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
        const columnNumber = WEEK_SHEET.COL_MONDAY + dayIndex;
        const requestedDayOffCell = scheduleSheet.getRange(requestedDayOffRow, columnNumber);
        requestedDayOffCell.insertCheckboxes();
        if (weekGrid[employeeIndex][dayIndex].type === "RDO") {
          requestedDayOffCell.setValue(true);
        }
      }
    }

    // --- SHIFT row (C–I and J): always written (first time or re-generation) ---
    // Clear the existing SHIFT row content before writing new values.
    scheduleSheet
      .getRange(shiftRow, WEEK_SHEET.COL_MONDAY, 1, WEEK_SHEET.DAYS_IN_WEEK)
      .clearContent();

    let totalPaidHoursThisWeek = 0;

    for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
      const columnNumber = WEEK_SHEET.COL_MONDAY + dayIndex;
      const cell         = weekGrid[employeeIndex][dayIndex];
      const shiftCell    = scheduleSheet.getRange(shiftRow, columnNumber);

      if (cell.type === "SHIFT") {
        // Look up the display text (e.g., "08:00 - 16:30") from the shift timing map.
        // The shift timing map is not available in this function, so we regenerate
        // the display text from the cell's data. The formatter receives the grid
        // which stores shiftName — the full timing map lookup happens here via
        // a helper that reconstructs the display string.
        // NOTE: display text is stored on the shiftDefinition; we pass it through
        // by storing it on the grid cell at assignment time via writeShiftDisplayText().
        shiftCell.setValue(cell.displayText || cell.shiftName || "SHIFT");
        totalPaidHoursThisWeek += cell.paidHours;
      } else if (cell.type === "VAC") {
        shiftCell.setValue("VAC");
      } else if (cell.type === "RDO") {
        shiftCell.setValue("RDO");
      } else {
        shiftCell.setValue("OFF");
      }
    }

    // Total hours cell (J): shows the employee's weekly paid hours total on the SHIFT row.
    scheduleSheet.getRange(shiftRow, WEEK_SHEET.COL_TOTAL_HOURS).setValue(totalPaidHoursThisWeek);
  });
}


/**
 * Writes the staffing summary block (REQUIRED / ACTUAL / STATUS rows) below the employee blocks.
 *
 * The summary gives managers an at-a-glance view of whether each day meets minimum staffing.
 *
 * REQUIRED row: the minimum staff count from the Settings sheet for each day.
 * ACTUAL row:   a COUNTIF formula that counts shift cells on each day column.
 *               The formula uses "*:*" as its pattern — this matches any string containing
 *               a colon (e.g., "08:00 - 16:30") and therefore counts only shift assignments
 *               while correctly ignoring "OFF", "VAC", "RDO", and blank cells.
 * STATUS row:   "OK" if actual >= required, "UNDER" otherwise. Conditional formatting
 *               in applyStatusRowConditionalFormat() colors these green or red.
 *
 * @param {Sheet}  scheduleSheet        — The schedule sheet to write to.
 * @param {number} employeeCount        — Total number of employee blocks written.
 * @param {Object} staffingRequirements — Minimum staff per day of week.
 */
function writeStaffingSummary(scheduleSheet, employeeCount, staffingRequirements) {
  // The summary block starts two rows below the last employee block.
  const lastEmployeeRow   = WEEK_SHEET.DATA_START_ROW + (employeeCount * WEEK_SHEET.ROWS_PER_EMPLOYEE) - 1;
  const summaryStartRow   = lastEmployeeRow + 2;
  const requiredRow       = summaryStartRow;
  const actualRow         = summaryStartRow + 1;
  const statusRow         = summaryStartRow + 2;

  // The first data row and the last employee SHIFT row — used in COUNTIF range.
  const dataStartRow      = WEEK_SHEET.DATA_START_ROW + WEEK_SHEET.ROW_OFFSET_SHIFT;
  const lastShiftRow      = lastEmployeeRow;

  // Row labels (column A).
  scheduleSheet.getRange(requiredRow, WEEK_SHEET.COL_ROW_LABEL).setValue("REQUIRED");
  scheduleSheet.getRange(actualRow,   WEEK_SHEET.COL_ROW_LABEL).setValue("ACTUAL");
  scheduleSheet.getRange(statusRow,   WEEK_SHEET.COL_ROW_LABEL).setValue("STATUS");

  scheduleSheet.getRange(requiredRow, WEEK_SHEET.COL_ROW_LABEL).setFontWeight("bold");
  scheduleSheet.getRange(actualRow,   WEEK_SHEET.COL_ROW_LABEL).setFontWeight("bold");
  scheduleSheet.getRange(statusRow,   WEEK_SHEET.COL_ROW_LABEL).setFontWeight("bold");

  // Write values for each of the 7 day columns.
  DAY_NAMES_IN_ORDER.forEach(function(dayName, dayIndex) {
    const columnNumber          = WEEK_SHEET.COL_MONDAY + dayIndex;
    const columnLetter          = columnIndexToLetter(columnNumber);
    const minimumStaffRequired  = staffingRequirements[dayName] || 0;

    // REQUIRED: static value from staffing requirements.
    scheduleSheet.getRange(requiredRow, columnNumber).setValue(minimumStaffRequired);

    // ACTUAL: formula counting cells in this day's column that contain a time range
    // string (identified by the presence of a colon character in the cell value).
    // The range spans only SHIFT rows — every 3rd row starting at dataStartRow.
    // Using COUNTIF on the full column range is simpler and still correct because
    // only SHIFT row cells ever contain a colon-format string; labels, blanks, and
    // checkbox values are all excluded by the "*:*" pattern.
    const countIfFormula =
      "=COUNTIF(" + columnLetter + dataStartRow + ":" + columnLetter + lastShiftRow + ",\"*:*\")";
    scheduleSheet.getRange(actualRow, columnNumber).setFormula(countIfFormula);

    // STATUS: formula comparing actual to required.
    const actualCellAddress   = columnLetter + actualRow;
    const requiredCellAddress = columnLetter + requiredRow;
    const statusFormula =
      "=IF(" + actualCellAddress + ">=" + requiredCellAddress + ",\"OK\",\"UNDER\")";
    scheduleSheet.getRange(statusRow, columnNumber).setFormula(statusFormula);
  });

  // Department total hours for the week (STATUS row, column J).
  // Summing column J from the first shift row to the last employee row captures all per-employee
  // weekly totals. VAC and RDO rows leave column J empty so they contribute 0 to the sum.
  const totalHoursColumnLetter = columnIndexToLetter(WEEK_SHEET.COL_TOTAL_HOURS);
  const departmentTotalFormula =
    "=SUM(" + totalHoursColumnLetter + dataStartRow + ":" + totalHoursColumnLetter + lastShiftRow + ")";
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
 *   FT SHIFT  → COLORS.FT_SHIFT  (blue)
 *   PT SHIFT  → COLORS.PT_SHIFT  (green)
 *   VAC       → COLORS.VACATION  (yellow)
 *   RDO / OFF → COLORS.DAY_OFF   (gray)
 *
 * Only SHIFT row cells receive color. VAC and RDO rows use a neutral background so the
 * shift row visually pops as the primary information row for each employee.
 *
 * @param {Sheet} scheduleSheet — The schedule sheet.
 * @param {Array} employeeList  — Employees in seniority order.
 * @param {Array} weekGrid      — The generated schedule grid.
 */
function applyShiftColors(scheduleSheet, employeeList, weekGrid) {
  employeeList.forEach(function(employee, employeeIndex) {
    const shiftRow = WEEK_SHEET.DATA_START_ROW +
      (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE) +
      WEEK_SHEET.ROW_OFFSET_SHIFT;

    for (let dayIndex = 0; dayIndex < WEEK_SHEET.DAYS_IN_WEEK; dayIndex++) {
      const columnNumber = WEEK_SHEET.COL_MONDAY + dayIndex;
      const cell         = weekGrid[employeeIndex][dayIndex];
      const shiftCell    = scheduleSheet.getRange(shiftRow, columnNumber);

      if (cell.type === "SHIFT") {
        // Use the employee's status to choose the correct color.
        // Blue for FT, green for PT, making it easy to see the mix at a glance.
        shiftCell.setBackground(
          employee.status === "FT" ? COLORS.FT_SHIFT : COLORS.PT_SHIFT
        );
      } else if (cell.type === "VAC") {
        shiftCell.setBackground(COLORS.VACATION);
      } else {
        // RDO and OFF cells both use the day-off color.
        shiftCell.setBackground(COLORS.DAY_OFF);
      }
    }
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
function applyUnderHoursHighlight(scheduleSheet, employeeList, weekGrid) {
  employeeList.forEach(function(employee, employeeIndex) {
    const shiftRow = WEEK_SHEET.DATA_START_ROW +
      (employeeIndex * WEEK_SHEET.ROWS_PER_EMPLOYEE) +
      WEEK_SHEET.ROW_OFFSET_SHIFT;

    const weeklyHours   = getWeeklyHours(weekGrid, employeeIndex);
    const weeklyMinimum = employee.status === "FT" ? HOUR_RULES.FT_MIN : HOUR_RULES.PT_MIN;

    // Name cell highlight (column B, SHIFT row) — red when the employee is under their minimum.
    // Applied to the SHIFT row's column B, which is the visible bottom of the merged name cell.
    const nameHighlightCell        = scheduleSheet.getRange(shiftRow, WEEK_SHEET.COL_EMPLOYEE_NAME);
    // Total hours cell highlight (column J, SHIFT row) — orange when an FT employee is over 40 hrs.
    const totalHoursHighlightCell  = scheduleSheet.getRange(shiftRow, WEEK_SHEET.COL_TOTAL_HOURS);

    if (employee.status === "FT" && weeklyHours > HOUR_RULES.FT_MAX) {
      // FT employees should never exceed 40 hours — orange on the total hours number signals overtime risk.
      nameHighlightCell.setBackground(null);
      totalHoursHighlightCell.setBackground(COLORS.OVER_HOURS_FT);
    } else if (weeklyHours < weeklyMinimum) {
      nameHighlightCell.setBackground(COLORS.UNDER_HOURS);
      totalHoursHighlightCell.setBackground(null);
    } else {
      // Clear any previous highlights in case of re-generation.
      nameHighlightCell.setBackground(null);
      totalHoursHighlightCell.setBackground(null);
    }
  });
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
function applyStatusRowConditionalFormat(scheduleSheet, employeeCount) {
  const lastEmployeeRow = WEEK_SHEET.DATA_START_ROW + (employeeCount * WEEK_SHEET.ROWS_PER_EMPLOYEE) - 1;
  const statusRow       = lastEmployeeRow + 4; // +2 gap, +1 REQUIRED, +1 ACTUAL, +1 STATUS = +4

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
  // Freezing rows 1–5 keeps the week label and column headers visible while scrolling.
  // Column A is NOT frozen because the week header row (row 1) has a merged cell spanning
  // all columns (A1:J1). Google Sheets does not allow freezing a column that would split
  // a merged cell — attempting to do so throws a runtime error.
  scheduleSheet.setFrozenRows(WEEK_SHEET.COLUMN_HEADER_ROW);

  // --- Row label column (A) background ---
  // Light gray background on the label column helps visually separate it from data cells.
  const totalDataRows = employeeCount * WEEK_SHEET.ROWS_PER_EMPLOYEE;
  if (totalDataRows > 0) {
    scheduleSheet
      .getRange(WEEK_SHEET.DATA_START_ROW, WEEK_SHEET.COL_ROW_LABEL, totalDataRows, 1)
      .setBackground(COLORS.ROW_LABEL_BG)
      .setHorizontalAlignment("center")
      .setFontStyle("italic");
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
