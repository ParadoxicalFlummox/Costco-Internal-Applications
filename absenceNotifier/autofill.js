/**
 * autofill.js — Autofills Employee ID and Department when a name is entered on a call log sheet.
 * VERSION: 0.2.0
 *
 * This file handles the onEdit trigger to provide a fast data-entry experience
 * for managers logging absences. When a manager types an employee's name in
 * column B of a call log entry row, this file:
 *
 *   1. Looks up that name in the "Employee Roster" sheet (case-insensitive).
 *   2. Copies the matching employee's ID into column C of the same row.
 *   3. Copies the matching employee's home department into column D.
 *
 * If the name is cleared (the cell is set back to blank), columns C and D are
 * also cleared so stale data from a previous entry does not linger.
 *
 * If the name does not match any roster entry, C and D are left blank and a
 * warning is written to the console. No error is thrown and the manager is not
 * interrupted — they can still complete the entry manually.
 *
 * TRIGGER SETUP:
 *   The onEdit function below must be registered as an INSTALLABLE trigger,
 *   not a simple trigger, because it writes to cells other than the one being
 *   edited (cross-range writes require elevated permissions).
 *   To install: Extensions → Apps Script → Triggers → Add Trigger
 *     Function: onEdit | Event: Spreadsheet → On Edit
 *
 * EMPLOYEE ROSTER SHEET:
 *   Lives in the same call log workbook. Must have:
 *     Column A — Employee Name
 *     Column B — Employee ID
 *     Column C — Home Department
 *   Row 1 is treated as a header and is skipped during lookup.
 *   Managed manually by the store admin (future: synced from AutoScheduler Roster).
 */


// ---------------------------------------------------------------------------
// Trigger Entry Point
// ---------------------------------------------------------------------------

/**
 * Responds to cell edits on the call log sheets and triggers autofill logic.
 *
 * This function is called by Apps Script for every edit in the workbook.
 * It applies two guards before doing any work:
 *
 *   Guard 1 — Sheet guard: Only fires on sheets whose name matches the call
 *     log naming patterns ("P# W#" or "Week Ending"). Edits on the config
 *     sheet, roster sheet, or any other tab are ignored immediately.
 *
 *   Guard 2 — Column guard: Only fires when the edited cell is in column B
 *     (the Employee Name column), on a data row (row 3 or later). Edits to
 *     dates, checkboxes, times, and comments do not trigger a roster lookup.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} event — The edit event from Apps Script.
 */
function onEdit(event) {
  const editedSheet  = event.range.getSheet();
  const editedColumn = event.range.getColumn();
  const editedRow    = event.range.getRow();

  // Guard 1: Only operate on call log sheets
  if (!isCallLogSheet_(editedSheet.getName())) return;

  // Guard 2: Only operate on the Name column (column B) in data rows
  if (editedColumn !== CALL_LOG_NAME_COLUMN_NUMBER) return; // defined in config.js (= 2)
  if (editedRow    < CALL_LOG_DATA_START_ROW)       return; // defined in config.js (= 3)

  const enteredName = event.range.getValue().toString().trim();

  if (enteredName === '') {
    // Name was cleared — clear the autofilled fields to avoid stale data
    clearAutofillFields_(editedSheet, editedRow);
    return;
  }

  // Attempt to look up the employee and populate ID and Department
  const employeeData = lookupEmployeeByName_(enteredName);

  if (employeeData) {
    writeAutofillFields_(editedSheet, editedRow, employeeData);
  } else {
    // No match found — leave C and D blank (don't overwrite if they were
    // previously filled by a successful lookup that was then re-typed)
    console.warn(`autofill: No roster match found for name "${enteredName}" on row ${editedRow}.`);
  }
}


// ---------------------------------------------------------------------------
// Roster Lookup
// ---------------------------------------------------------------------------

/**
 * Looks up an employee by name in the Employee Roster sheet and returns their
 * ID and department if a match is found.
 *
 * The lookup is case-insensitive to handle minor capitalization differences
 * between how a manager types a name and how it appears in the roster. For
 * example, "john smith" will match "John Smith".
 *
 * Performance note: The roster is read in a single getValues() call and
 * scanned in memory. For a typical store roster of a few hundred employees
 * this is fast. If the roster were very large (thousands of rows), a Map
 * pre-built on first call would be more appropriate, but that optimization
 * is not needed here.
 *
 * @param {string} name — The employee name as entered by the manager (trimmed).
 * @returns {{ employeeId: string, department: string } | null}
 *   The matched employee's data, or null if no match was found.
 */
function lookupEmployeeByName_(name) {
  const workbook     = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet  = workbook.getSheetByName(ROSTER_SHEET_NAME); // defined in config.js

  if (!rosterSheet) {
    console.warn(`autofill: Roster sheet "${ROSTER_SHEET_NAME}" not found. Cannot autofill.`);
    return null;
  }

  const lastRow = rosterSheet.getLastRow();
  if (lastRow < 2) {
    // Sheet has only the header row (or is empty) — nothing to look up
    return null;
  }

  // Read all roster rows in one call: columns A (Name), B (ID), C (Department)
  // Row 1 is the header, so data starts at row 2.
  const numberOfDataRows = lastRow - 1;
  const rosterData = rosterSheet
    .getRange(2, 1, numberOfDataRows, 3)
    .getValues();

  const lowercasedName = name.toLowerCase();

  for (let rowIndex = 0; rowIndex < rosterData.length; rowIndex++) {
    const rosterName = String(rosterData[rowIndex][0] || '').trim().toLowerCase();
    if (rosterName === lowercasedName) {
      return {
        employeeId:  String(rosterData[rowIndex][1] || '').trim(),
        department:  String(rosterData[rowIndex][2] || '').trim(),
      };
    }
  }

  return null; // No matching name found in the roster
}


// ---------------------------------------------------------------------------
// Sheet Write Helpers
// ---------------------------------------------------------------------------

/**
 * Writes the autofilled Employee ID and Department to their columns on the
 * call log entry row.
 *
 * Writes are done as a single setValues() call on a two-cell range to minimize
 * the number of GAS write operations (each write call has overhead).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The call log sheet being edited.
 * @param {number} rowNumber — The 1-based row number of the entry being filled.
 * @param {{ employeeId: string, department: string }} employeeData — Data to write.
 */
function writeAutofillFields_(sheet, rowNumber, employeeData) {
  // Columns C and D are CALL_LOG_COLUMNS.EMPLOYEE_ID and DEPT (0-indexed 2 and 3)
  // which are 1-indexed columns 3 and 4.
  sheet.getRange(rowNumber, 3, 1, 2).setValues([[
    employeeData.employeeId,
    employeeData.department,
  ]]);
}

/**
 * Clears the autofilled Employee ID and Department columns on the given row.
 *
 * Called when the manager clears the name in column B, ensuring that no stale
 * ID or department data from a previous entry remains on the row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet — The call log sheet being edited.
 * @param {number} rowNumber — The 1-based row number to clear.
 */
function clearAutofillFields_(sheet, rowNumber) {
  sheet.getRange(rowNumber, 3, 1, 2).clearContent();
}


// ---------------------------------------------------------------------------
// Sheet Name Guard
// ---------------------------------------------------------------------------

/**
 * Returns true if the given sheet name is a call log sheet that should have
 * autofill applied.
 *
 * Call log sheet names follow one of two patterns:
 *   - "P# W#" (fiscal period/week): one or more digits for period, one or
 *     more digits for week, e.g. "P3 W1" or "P12 W4".
 *   - "Week Ending MM/DD/YY": the fallback format used when no FY start date
 *     is configured.
 *
 * This guard prevents onEdit from attempting roster lookups on the config
 * sheet, the roster sheet itself, or any unrelated tabs.
 *
 * @param {string} sheetName — The name of the sheet that was edited.
 * @returns {boolean} true if this sheet is a call log sheet.
 */
function isCallLogSheet_(sheetName) {
  const fiscalPattern     = /^P\d+\s+W\d+$/i;         // e.g. "P3 W1"
  const weekEndingPattern = /^Week Ending \d+\/\d+\/\d+$/i; // e.g. "Week Ending 10/19/25"
  return fiscalPattern.test(sheetName) || weekEndingPattern.test(sheetName);
}
