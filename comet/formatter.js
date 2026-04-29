/**
 * formatter.js — Writes schedule data to the Week sheet in JSON format.
 * VERSION: 0.6.0
 *
 * In the new architecture, the Week sheet is pure data storage, not a visual display.
 * One row per employee; one JSON string per row in column C.
 * Managers interact entirely via the web UI; the sheet is a backend database table.
 *
 * The JSON cell payload preserves all schedule information: shift name (was lost before),
 * lock flags, role assignments, and paid hours. A single setValues() call writes all
 * employee rows at once — no per-employee batching complexity.
 */


/**
 * The single entry point: writes a complete schedule to the Week sheet in JSON format.
 *
 * @param {Sheet}  scheduleSheet       — The Week_MM_DD_YY_[Dept] sheet.
 * @param {Array}  employeeList        — Employees in seniority order (from scheduleEngine.js).
 * @param {Array}  weekGrid            — The generated schedule grid (from scheduleEngine.js).
 * @param {string} weekStartDate       — The Monday of the week (Date object).
 * @param {string} departmentName      — The department name to display in the sheet header.
 */
function writeAndFormatSchedule(scheduleSheet, employeeList, weekGrid, _staffingRequirements, weekStartDate, departmentName) {
  // Write header rows (unchanged from before)
  writeWeekHeader_(scheduleSheet, weekStartDate, departmentName);

  // Write column headers
  scheduleSheet
    .getRange(WEEK_SHEET.COLUMN_HEADER_ROW, 1, 1, 5)
    .setValues([['Name', 'Employee ID', 'Schedule', 'Total Hours', 'Stored At']])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Build the data rows: one per employee, JSON payload in col C
  const dataRows = [];
  const now = new Date().toISOString();

  employeeList.forEach(function (employee, employeeIndex) {
    const scheduleObj = {};
    let totalHours = 0;

    // Build the JSON object for all 7 days
    DAY_NAMES_IN_ORDER.forEach(function (dayName, dayIndex) {
      const cell = weekGrid[employeeIndex][dayIndex];
      scheduleObj[dayName] = {
        type: cell.type,
        shiftName: cell.shiftName || null,
        displayText: cell.displayText || null,
        paidHours: cell.paidHours || 0,
        role: cell.role || null,
        locked: cell.locked || false,
      };
      if (cell.type === 'SHIFT') totalHours += (cell.paidHours || 0);
    });

    dataRows.push([
      employee.name,
      employee.employeeId || '',
      JSON.stringify(scheduleObj),
      totalHours,
      now,
    ]);
  });

  // Write all employee rows in one batch call (huge performance win)
  if (dataRows.length > 0) {
    scheduleSheet
      .getRange(WEEK_SHEET.DATA_START_ROW, 1, dataRows.length, 5)
      .setValues(dataRows);
  }
}


/**
 * Writes the schedule sheet header rows (rows 1–4).
 *
 * @param {Sheet}  scheduleSheet  — The schedule sheet.
 * @param {Date}   weekStartDate  — The Monday of the week.
 * @param {string} departmentName — The department name.
 */
function writeWeekHeader_(scheduleSheet, weekStartDate, departmentName) {
  const weekEndDate = new Date(weekStartDate);
  weekEndDate.setDate(weekEndDate.getDate() + 6); // Sunday

  const weekLabel =
    'Week of ' +
    weekStartDate.toLocaleDateString('en-US', { month: 'long', day: 'numeric' }) +
    ' – ' +
    weekEndDate.getDate() + ', ' + weekEndDate.getFullYear();

  // Row 1: Week label
  const titleRange = scheduleSheet.getRange(WEEK_SHEET.HEADER_ROW, 1, 1, WEEK_SHEET.COL_STORED_AT);
  titleRange.merge();
  titleRange.setValue(weekLabel);
  titleRange.setFontSize(14);
  titleRange.setFontWeight('bold');
  titleRange.setBackground(COLORS.HEADER_BG);
  titleRange.setFontColor(COLORS.HEADER_TEXT);
  titleRange.setHorizontalAlignment('center');
  titleRange.setVerticalAlignment('middle');

  // Row 2: Generation timestamp
  const now = new Date();
  const timestampText =
    'Generated: ' +
    now.toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' }) +
    ' at ' +
    now.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit' });

  const timestampRange = scheduleSheet.getRange(WEEK_SHEET.TIMESTAMP_ROW, 1, 1, WEEK_SHEET.COL_STORED_AT);
  timestampRange.merge();
  timestampRange.setValue(timestampText);
  timestampRange.setFontSize(10);
  timestampRange.setFontColor('#666666');
  timestampRange.setHorizontalAlignment('left');

  // Row 3: Department name
  const deptText = 'Department: ' + departmentName;
  const deptRange = scheduleSheet.getRange(WEEK_SHEET.DEPARTMENT_ROW, 1, 1, WEEK_SHEET.COL_STORED_AT);
  deptRange.merge();
  deptRange.setValue(deptText);
  deptRange.setFontSize(11);
  deptRange.setFontWeight('bold');
  deptRange.setFontColor('#000000');
  deptRange.setHorizontalAlignment('left');

  // Row 4: Spacer
  scheduleSheet.getRange(4, 1, 1, WEEK_SHEET.COL_STORED_AT).setBackground('#EEEEEE');
}


/**
 * Reads schedule data from the JSON-format Week sheet.
 *
 * @param {Sheet} scheduleSheet — The Week sheet.
 * @returns {{ employeeRows: Array<{ name, employeeId, scheduleJson, totalHours, storedAt }> }}
 */
function readJsonSchedule(scheduleSheet) {
  const dataRange = scheduleSheet.getRange(WEEK_SHEET.DATA_START_ROW, 1, scheduleSheet.getLastRow() - WEEK_SHEET.DATA_START_ROW + 1, 5);
  const allRows = dataRange.getValues();

  const employeeRows = [];
  allRows.forEach(function (row) {
    if (!row[WEEK_SHEET.COL_NAME - 1]) return; // skip blank rows
    employeeRows.push({
      name: row[WEEK_SHEET.COL_NAME - 1],
      employeeId: row[WEEK_SHEET.COL_EMPLOYEE_ID - 1],
      scheduleJson: row[WEEK_SHEET.COL_SCHEDULE_JSON - 1],
      totalHours: row[WEEK_SHEET.COL_TOTAL_HOURS - 1],
      storedAt: row[WEEK_SHEET.COL_STORED_AT - 1],
    });
  });

  return { employeeRows: employeeRows };
}


/**
 * Detects whether a Week sheet uses the new JSON format or the legacy 5-row format.
 * Used for backward compatibility during the transition.
 *
 * @param {Sheet} scheduleSheet — The Week sheet.
 * @returns {'json'|'legacy'} Format identifier.
 */
function detectSheetFormat(scheduleSheet) {
  const headerCell = scheduleSheet.getRange(WEEK_SHEET.COLUMN_HEADER_ROW, WEEK_SHEET.COL_SCHEDULE_JSON).getValue();
  return headerCell === 'Schedule' ? 'json' : 'legacy';
}
