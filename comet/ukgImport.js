/**
 * ukgImport.js — UKG employee data import for COMET.
 * VERSION: 0.2.3
 *
 * This file owns the server-side logic for upserting employee rows into the
 * Employees sheet from a parsed UKG CSV export.
 *
 * DIVISION OF RESPONSIBILITY:
 *   The frontend (javascript.html) handles CSV parsing and row filtering.
 *   It sends a clean array of employee objects to importEmployeesFromUkg_().
 *   This file handles only the sheet read/write — no CSV parsing happens here.
 *
 * UPSERT LOGIC:
 *   Each incoming row is matched against the Employees sheet by Employee ID.
 *   - ID found    → update Name, Hire Date, Department (Status is preserved)
 *   - ID not found → append a new row with Status = "Active"
 *
 * EMPLOYEES SHEET LAYOUT (columns from config.js EMPLOYEE_COLUMN):
 *   A — Name (Last, First)
 *   B — Employee ID
 *   C — Hire Date
 *   D — Department
 *   E — Status ("Active" or "Archived")
 *
 * RETURN VALUE:
 *   { added: number, updated: number, skipped: number }
 *
 *   "skipped" means the incoming row had a blank ID or name — those are
 *   dropped client-side before they reach this function, so skipped will
 *   normally be 0. It is included for completeness and future extensibility.
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Upserts employee rows into the Employees sheet.
 *
 * Called by api.js importFromUKG(). The rows array has already been filtered
 * by the frontend — no placeholder rows, no blank IDs.
 *
 * @param {Array<{ name: string, id: string, hireDate: string, department: string }>} rows
 * @returns {{ added: number, updated: number, skipped: number }}
 */
function importEmployeesFromUkg_(rows) {
  if (!rows || rows.length === 0) return { added: 0, updated: 0, skipped: 0 };

  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateEmployeesSheet_(workbook);

  // Read the current employee table into memory for fast ID lookup.
  // Each row maps to its 1-indexed sheet row number for in-place updates.
  const existingEmployees = readEmployeeIndex_(sheet);

  const newRows = [];
  let updated = 0;
  let skipped = 0;

  rows.forEach(incoming => {
    const id = String(incoming.id || '').trim();
    const name = String(incoming.name || '').trim();

    if (!id || !name) {
      skipped++;
      return;
    }

    if (existingEmployees.has(id)) {
      // Update existing row — preserve Status, update everything else
      const sheetRow = existingEmployees.get(id);
      sheet.getRange(sheetRow, EMPLOYEE_COLUMN.NAME).setValue(name);          // config.js
      sheet.getRange(sheetRow, EMPLOYEE_COLUMN.HIRE_DATE).setValue(incoming.hireDate || '');
      sheet.getRange(sheetRow, EMPLOYEE_COLUMN.DEPARTMENT).setValue(incoming.department || '');
      updated++;
    } else {
      // Queue new row for batch append
      newRows.push([
        name,
        id,
        incoming.hireDate || '',
        incoming.department || '',
        'Active',
      ]);
    }
  });

  // Batch-append all new rows at once to minimize API calls
  if (newRows.length > 0) {
    const firstNewRow = sheet.getLastRow() + 1;
    sheet.getRange(firstNewRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  }

  SpreadsheetApp.flush();

  // Record the timestamp of this import to the COMET Config sheet
  recordUkgImportTimestamp_(workbook);

  return { added: newRows.length, updated, skipped };
}


// ---------------------------------------------------------------------------
// Read All Employees
// ---------------------------------------------------------------------------

/**
 * Returns all employees from the Employees sheet as an array of plain objects.
 * Includes both Active and Archived employees; the caller filters as needed.
 *
 * @returns {Array<{ name: string, id: string, hireDate: string, department: string, status: string }>}
 */
function getAllEmployees_() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateEmployeesSheet_(workbook);

  const lastRow = sheet.getLastRow();
  if (lastRow < EMPLOYEES_DATA_START_ROW) return []; // config.js

  const numRows = lastRow - EMPLOYEES_DATA_START_ROW + 1;
  // Read all 13 columns (A–M) so schedule-specific fields F–M are included.
  const numCols = Math.max(5, EMPLOYEE_COLUMN.SENIORITY_RANK); // 13
  const data = sheet.getRange(EMPLOYEES_DATA_START_ROW, 1, numRows, numCols).getValues();

  return data
    .filter(row => String(row[EMPLOYEE_COLUMN.ID - 1] || '').trim() !== '')
    .map(row => {
      const hireDateRaw = row[EMPLOYEE_COLUMN.HIRE_DATE - 1];
      let hireDate = '';
      if (hireDateRaw instanceof Date && !isNaN(hireDateRaw.getTime())) {
        hireDate = Utilities.formatDate(hireDateRaw, Session.getScriptTimeZone(), 'MM/dd/yyyy');
      } else if (hireDateRaw) {
        hireDate = String(hireDateRaw).trim();
      }
      return {
        name:            String(row[EMPLOYEE_COLUMN.NAME            - 1] || '').trim(),
        id:              String(row[EMPLOYEE_COLUMN.ID              - 1] || '').trim(),
        hireDate,
        department:      String(row[EMPLOYEE_COLUMN.DEPARTMENT      - 1] || '').trim(),
        status:          String(row[EMPLOYEE_COLUMN.STATUS          - 1] || 'Active').trim(),
        ftpt:            String(row[EMPLOYEE_COLUMN.FTPT            - 1] || '').trim(),
        dayOffPrefOne:   String(row[EMPLOYEE_COLUMN.DAY_OFF_PREF_ONE - 1] || '').trim(),
        dayOffPrefTwo:   String(row[EMPLOYEE_COLUMN.DAY_OFF_PREF_TWO - 1] || '').trim(),
        preferredShift:  String(row[EMPLOYEE_COLUMN.PREFERRED_SHIFT - 1] || '').trim(),
        qualifiedShifts: String(row[EMPLOYEE_COLUMN.QUALIFIED_SHIFTS - 1] || '').trim(),
        vacationDates:   String(row[EMPLOYEE_COLUMN.VACATION_DATES  - 1] || '').trim(),
        role:            String(row[EMPLOYEE_COLUMN.ROLE            - 1] || '').trim(),
        seniorityRank:   Number(row[EMPLOYEE_COLUMN.SENIORITY_RANK  - 1] || 0),
      };
    });
}

/**
 * Returns only Active employees. Used by the infraction scanner, schedule
 * generator, and absence log to exclude archived employees from operations.
 *
 * @returns {Array<{ name, id, hireDate, department, status }>}
 */
function getActiveEmployees_() {
  return getAllEmployees_().filter(emp => emp.status === 'Active');
}

/**
 * Sets the Status column for a single employee identified by ID.
 *
 * @param {string} id — Employee ID to update.
 * @param {'Active'|'Archived'} status
 * @returns {boolean} true if the employee was found and updated, false otherwise.
 */
function setEmployeeStatus_(id, status) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateEmployeesSheet_(workbook);
  const index = readEmployeeIndex_(sheet);

  if (!index.has(id)) return false;

  const sheetRow = index.get(id);
  sheet.getRange(sheetRow, EMPLOYEE_COLUMN.STATUS).setValue(status); // config.js
  SpreadsheetApp.flush();
  return true;
}


// ---------------------------------------------------------------------------
// Employees Sheet Bootstrap
// ---------------------------------------------------------------------------

/**
 * Returns the Employees sheet, creating and formatting it if it does not
 * exist. This is the only place in COMET that creates this sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateEmployeesSheet_(workbook) {
  let sheet = workbook.getSheetByName(EMPLOYEES_SHEET_NAME); // config.js
  if (sheet) return sheet;

  sheet = workbook.insertSheet(EMPLOYEES_SHEET_NAME);
  writeEmployeesSheetHeader_(sheet);
  return sheet;
}

/**
 * Writes the header row and applies basic formatting to a new Employees sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function writeEmployeesSheetHeader_(sheet) {
  const headers = [
    'Name (Last, First)', 'Employee ID', 'Hire Date', 'Department', 'Status',
    'FT/PT', 'Day Off Pref 1', 'Day Off Pref 2', 'Preferred Shift',
    'Qualified Shifts', 'Vacation Dates', 'Role', 'Seniority Rank',
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#005DAA')
    .setFontColor('#FFFFFF');

  sheet.setColumnWidth(1,  200); // Name
  sheet.setColumnWidth(2,  110); // ID
  sheet.setColumnWidth(3,  110); // Hire Date
  sheet.setColumnWidth(4,  160); // Department
  sheet.setColumnWidth(5,   90); // Status
  sheet.setColumnWidth(6,   70); // FT/PT
  sheet.setColumnWidth(7,  120); // Day Off Pref 1
  sheet.setColumnWidth(8,  120); // Day Off Pref 2
  sheet.setColumnWidth(9,  140); // Preferred Shift
  sheet.setColumnWidth(10, 180); // Qualified Shifts
  sheet.setColumnWidth(11, 180); // Vacation Dates
  sheet.setColumnWidth(12, 160); // Role
  sheet.setColumnWidth(13, 110); // Seniority Rank

  sheet.setFrozenRows(1);
}


// ---------------------------------------------------------------------------
// Internal Helpers
// ---------------------------------------------------------------------------

/**
 * Writes schedule-specific fields (columns F–M) for a single employee.
 *
 * Only fields present in the `fields` object are written; omitted keys are
 * left unchanged so partial updates are safe.
 *
 * @param {string} id     — Employee ID.
 * @param {{
 *   ftpt?: string,
 *   dayOffPrefOne?: string,
 *   dayOffPrefTwo?: string,
 *   preferredShift?: string,
 *   qualifiedShifts?: string,
 *   vacationDates?: string,
 *   role?: string,
 * }} fields
 * @returns {boolean} true if the employee was found and updated.
 */
function updateEmployeeScheduleFields_(id, fields) {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateEmployeesSheet_(workbook);
  const index = readEmployeeIndex_(sheet);

  if (!index.has(String(id))) return false;

  const sheetRow = index.get(String(id));
  const writes = [
    [EMPLOYEE_COLUMN.FTPT,              'ftpt'],
    [EMPLOYEE_COLUMN.DAY_OFF_PREF_ONE,  'dayOffPrefOne'],
    [EMPLOYEE_COLUMN.DAY_OFF_PREF_TWO,  'dayOffPrefTwo'],
    [EMPLOYEE_COLUMN.PREFERRED_SHIFT,   'preferredShift'],
    [EMPLOYEE_COLUMN.QUALIFIED_SHIFTS,  'qualifiedShifts'],
    [EMPLOYEE_COLUMN.VACATION_DATES,    'vacationDates'],
    [EMPLOYEE_COLUMN.ROLE,              'role'],
  ];

  writes.forEach(([col, key]) => {
    if (fields[key] != null) {
      sheet.getRange(sheetRow, col).setValue(fields[key]);
    }
  });

  SpreadsheetApp.flush();
  return true;
}


/**
 * Builds a Map of { employeeId → sheetRowNumber } from the Employees sheet.
 * Used for O(1) lookup during upsert.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Map<string, number>}
 */
function readEmployeeIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  const index = new Map();

  if (lastRow < EMPLOYEES_DATA_START_ROW) return index; // config.js

  const numRows = lastRow - EMPLOYEES_DATA_START_ROW + 1;
  const idColumn = sheet
    .getRange(EMPLOYEES_DATA_START_ROW, EMPLOYEE_COLUMN.ID, numRows, 1)
    .getValues();

  idColumn.forEach((row, offset) => {
    const id = String(row[0] || '').trim();
    if (id) index.set(id, EMPLOYEES_DATA_START_ROW + offset);
  });

  return index;
}


// ---------------------------------------------------------------------------
// Config Recording
// ---------------------------------------------------------------------------

/**
 * Records the timestamp of the UKG import to the COMET Config sheet.
 * This allows the Schedule view to warn if the employee roster is stale.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} workbook
 */
function recordUkgImportTimestamp_(workbook) {
  try {
    const configSheet = workbook.getSheetByName(COMET_CONFIG_SHEET_NAME); // config.js
    if (!configSheet) return; // Config sheet not created yet (shouldn't happen in normal use)

    // Find or create the "ukgImportLastRan" row in the config sheet
    const lastRow = configSheet.getLastRow();
    let configRowFound = false;
    let configRow = 0;

    // Search rows 3+ for the key "ukgImportLastRan"
    if (lastRow >= 3) {
      const configData = configSheet.getRange(3, 1, lastRow - 2, 2).getValues();
      for (let i = 0; i < configData.length; i++) {
        if (String(configData[i][0] || '').trim() === 'ukgImportLastRan') {
          configRow = 3 + i; // Convert back to 1-indexed sheet row
          configRowFound = true;
          break;
        }
      }
    }

    // If not found, append a new row
    if (!configRowFound) {
      configRow = lastRow + 1;
    }

    // Record the current timestamp in ISO format
    const now = new Date();
    const timestamp = now.toISOString();

    configSheet.getRange(configRow, 1).setValue('ukgImportLastRan');
    configSheet.getRange(configRow, 2).setValue(timestamp);

    SpreadsheetApp.flush();
  } catch (error) {
    // Log the error but don't fail the import — the import itself succeeded
    console.warn('ukgImport: Failed to record import timestamp:', error);
  }
}
