/**
 * calendarParser.js — Reads JSON attendance data and produces CalendarEvent objects.
 * VERSION: 0.3.0
 *
 * This file owns all logic for turning JSON attendance data (Phase 2) into
 * structured event objects that the infraction detector can work with.
 *
 * SINGLE RESPONSIBILITY:
 *   Parse JSON attendance records and normalize codes into CalendarEvent objects.
 *   All geometry and grid parsing is eliminated. The shape of CalendarEvent
 *   remains unchanged so downstream consumers (infractionDetector, etc.) need
 *   no modifications.
 *
 * OUTPUT SHAPE — CalendarEvent (unchanged):
 *   {
 *     employeeName:  string  — From the tab name or import payload
 *     employeeId:    string  — From the tab name or import payload
 *     department:    string  — From employee record
 *     hireDate:      Date    — From employee record
 *     month:         string  — Month name derived from the date
 *     date:          Date    — Parsed from ISO date string in JSON
 *     code:          string  — Normalized attendance code
 *     isInfraction:  boolean — True if the code is an infraction
 *     isIgnored:     boolean — True if the code is in IGNORE_CODES
 *     a1:            string  — "source:JSONkey" for debugging
 *   }
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Parses all calendar events from a single employee's attendance JSON tab.
 *
 * Reads the sheet, finds the row for the requested year, parses the JSON,
 * and emits one CalendarEvent per code per date.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The employee's attendance tab.
 * @param {number} year   — The calendar year to read (e.g. 2026).
 * @param {string} timeZone — The script time zone string.
 * @returns {CalendarEvent[]}
 */
function parseCalendarEventsFromJson_(sheet, year, timeZone) {
  const sheetName = sheet.getName();
  const events = [];

  // Extract employeeId from tab name: "Last, First - ID" → ID
  const idMatch = sheetName.match(/-\s*(\d+)\s*$/);
  const employeeId = idMatch ? idMatch[1] : 'unknown';

  // Read all year rows in one call
  const lastRow = sheet.getLastRow();
  if (lastRow < ATTENDANCE_JSON.DATA_START_ROW) {
    return events; // No data rows
  }

  const dataRange = sheet.getRange(ATTENDANCE_JSON.DATA_START_ROW, 1,
                                   lastRow - ATTENDANCE_JSON.DATA_START_ROW + 1, 3);
  const allRows = dataRange.getValues();

  // Find the row for this year
  let codesJson = null;
  for (let i = 0; i < allRows.length; i++) {
    if (allRows[i][ATTENDANCE_JSON.COL_YEAR - 1] === year ||
        String(allRows[i][ATTENDANCE_JSON.COL_YEAR - 1]) === String(year)) {
      codesJson = allRows[i][ATTENDANCE_JSON.COL_CODES_JSON - 1];
      break;
    }
  }

  if (!codesJson) {
    return events; // Year not found
  }

  // Parse the JSON
  let schedule = {};
  try {
    schedule = JSON.parse(codesJson);
  } catch (e) {
    Logger.log('calendarParser: Failed to parse JSON for ' + sheetName + ': ' + e.toString());
    return events;
  }

  // Get employee metadata from the Employees sheet
  const employeeRecord = getEmployeeById_(employeeId);
  if (!employeeRecord) {
    Logger.log('calendarParser: Employee not found: ' + employeeId);
    return events;
  }

  // Iterate over each date in the sparse JSON object
  for (const isoDate in schedule) {
    if (!schedule.hasOwnProperty(isoDate)) continue;

    const codesArray = schedule[isoDate];
    if (!Array.isArray(codesArray)) continue;

    const date = new Date(isoDate);
    const monthIndex = date.getMonth();
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                        'July', 'August', 'September', 'October', 'November', 'December'];
    const monthName = monthNames[monthIndex];

    codesArray.forEach(function(code) {
      if (!code) return;

      const normalizedCode = normalizeAndSplitCodes_(code)[0]; // normalizeAndSplitCodes_ already handles splitting
      if (!normalizedCode) return;

      const isIgnored = IGNORE_CODES.indexOf(normalizedCode) >= 0;
      const inInfractionList = INFRACTION_CODES.indexOf(normalizedCode) >= 0;
      const hasCodeRule = !!(CODE_RULES && CODE_RULES[normalizedCode]);
      const isInfraction = (inInfractionList || hasCodeRule) && !isIgnored;

      events.push({
        employeeName: employeeRecord.name,
        employeeId: employeeId,
        department: employeeRecord.department,
        hireDate: employeeRecord.hireDate,
        month: monthName,
        date: date,
        code: normalizedCode,
        isInfraction: isInfraction,
        isIgnored: isIgnored,
        a1: 'JSON:' + isoDate,
      });
    });
  }

  return events;
}


/**
 * Returns true if the given sheet tab looks like an individual employee
 * attendance tab (matches the EMPLOYEE_TAB_PATTERN).
 *
 * @param {string} sheetName — The tab name to test.
 * @returns {boolean}
 */
function isEmployeeTab_(sheetName) {
  return EMPLOYEE_TAB_PATTERN.test(sheetName);
}


// ---------------------------------------------------------------------------
// Helper: Employee Record Lookup
// ---------------------------------------------------------------------------

/**
 * Looks up an employee record from the master Employees sheet by ID.
 *
 * @param {string} employeeId — The employee ID to search for.
 * @returns {{ name: string, id: string, department: string, hireDate: Date }|null}
 */
function getEmployeeById_(employeeId) {
  const employeeSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(EMPLOYEES_SHEET_NAME);
  if (!employeeSheet) return null;

  const data = employeeSheet.getRange(EMPLOYEES_DATA_START_ROW, 1,
                                      employeeSheet.getLastRow() - EMPLOYEES_DATA_START_ROW + 1, 5)
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][EMPLOYEE_COLUMN.ID - 1] || '').trim();
    if (id === employeeId || id === String(employeeId)) {
      return {
        name: String(data[i][EMPLOYEE_COLUMN.NAME - 1] || '').trim(),
        id: id,
        department: String(data[i][EMPLOYEE_COLUMN.DEPARTMENT - 1] || '').trim(),
        hireDate: data[i][EMPLOYEE_COLUMN.HIRE_DATE - 1],
      };
    }
  }

  return null;
}
