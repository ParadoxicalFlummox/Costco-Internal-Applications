/**
 * autofill.js — Row activation and employee autofill for call log entry rows.
 * VERSION: 0.2.7
 *
 * This file handles the onEdit trigger for two distinct actions on call log sheets:
 *
 *   ACTION 1 — ROW ACTIVATION (column A checkbox):
 *     The sheet starts with one row containing only a checkbox in column A.
 *     When a manager checks that box to begin logging a new absence:
 *       1. The checkbox is removed and today's date is stamped into column A.
 *       2. Call-Out, FMLA, and No-Show checkboxes are added to columns E, F, G
 *          of that row so the manager can record the absence type.
 *       3. A new trigger checkbox is inserted into column A of the next row,
 *          ready for the next absence entry.
 *     This keeps the sheet clean — only rows that have been deliberately opened
 *     by a manager contain data, and there is always exactly one blank row
 *     waiting at the bottom.
 *
 *   ACTION 2 — EMPLOYEE AUTOFILL (column B name entry):
 *     When a manager types an employee's name in column B:
 *       1. The name is looked up in the "Employee Roster" sheet (case-insensitive).
 *       2. The matching employee's ID is written into column C.
 *       3. The matching department is written into column D.
 *     If the name is cleared, C and D are also cleared.
 *     If no match is found, C and D are left blank without interrupting the manager.
 *
 * TRIGGER SETUP:
 *   Must be registered as an INSTALLABLE trigger (not a simple trigger) because:
 *     1. It writes to cells outside the edited cell.
 *     2. It calls SpreadsheetApp.openById() to access the external attendance
 *        controller — simple triggers are NOT permitted to access other spreadsheets.
 *   The function is named onEditHandler (not onEdit) intentionally. A function
 *   named "onEdit" runs automatically as a simple trigger, which would fire a
 *   second time alongside the installable trigger and fail on the external lookup.
 *   To install: Extensions → Apps Script → Triggers → Add Trigger
 *     Function: onEditHandler | Event: From spreadsheet → On Edit
 *
 * EMPLOYEE ROSTER SHEET:
 *   Lives in the same workbook. Must have:
 *     Column A — Employee Name
 *     Column B — Employee ID
 *     Column C — Home Department
 *   Row 1 is treated as a header and skipped during lookup.
 */


// ---------------------------------------------------------------------------
// Trigger Entry Point
// ---------------------------------------------------------------------------

/**
 * Responds to cell edits on the call log sheets and routes to the correct action.
 *
 * Two edit types are handled; all others are ignored immediately:
 *
 *   Column A (DATE / trigger checkbox), data rows:
 *     When the checkbox is checked (value === true), activateEntryRow_() is
 *     called to stamp today's date, add absence-type checkboxes, and create
 *     the next trigger row.
 *
 *   Column B (EMPLOYEE NAME), data rows:
 *     When a name is entered, lookupEmployeeByName_() is called and the
 *     result is written to columns C and D. When the name is cleared,
 *     columns C and D are cleared too.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} event — The edit event from Apps Script.
 */
function onEditHandler(event) {
  const editedSheet  = event.range.getSheet();
  const editedColumn = event.range.getColumn();
  const editedRow    = event.range.getRow();

  // Ignore edits on non-call-log sheets entirely
  if (!isCallLogSheet_(editedSheet.getName())) return;

  // Ignore edits above the data start row (header rows)
  if (editedRow < CALL_LOG_DATA_START_ROW) return;

  // --- Route: Column A — entry trigger checkbox ---
  if (editedColumn === 1) {
    // Only act when the checkbox transitions to TRUE (checked).
    // Unchecking or any other change is ignored.
    if (event.range.getValue() === true) {
      activateEntryRow_(editedSheet, editedRow);
    }
    return;
  }

  // --- Route: Column M — SEND / notify checkbox ---
  if (editedColumn === CALL_LOG_NOTIFY_COLUMN_NUMBER) { // defined in config.js (= 13)
    // Only act when the checkbox is checked. If the manager accidentally unchecks
    // a "Sent …" cell that has already been converted to plain text, this guard
    // prevents a duplicate send (the cell value would be a string, not true).
    if (event.range.getValue() === true) {
      sendEntryNotification_(editedSheet, editedRow);
    }
    return;
  }

  // --- Route: Column B — employee name autofill ---
  if (editedColumn === CALL_LOG_NAME_COLUMN_NUMBER) { // defined in config.js (= 2)
    const enteredName = event.range.getValue().toString().trim();

    if (enteredName === '') {
      // Name was cleared — remove stale autofilled data and any suggestion dropdown
      clearAutofillFields_(editedSheet, editedRow);
      return;
    }

    const matches = searchEmployeesByName_(enteredName);

    if (matches.length === 1) {
      // Exactly one match — autofill immediately, no ambiguity
      writeAutofillFields_(editedSheet, editedRow, matches[0], matches[0].displayName);

    } else if (matches.length > 1) {
      // Multiple matches — show a dropdown so the clerk can pick the right person.
      // We cap at 15 suggestions to keep the dropdown readable; if there are more
      // the search term is too short and the clerk should type more characters.
      const suggestions = matches.slice(0, 15);
      showSuggestionDropdown_(editedSheet, editedRow, suggestions);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `${matches.length} employees match "${enteredName}" — select the correct name from the dropdown.`,
        'Multiple Matches',
        8
      );

    } else {
      // No match at all — notify the clerk
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `No employees found matching "${enteredName}". Check spelling or try a partial name.`,
        'No Match Found',
        6
      );
      // Even with no roster match, link the name cell to the attendance controller
      // root so the clerk can navigate there with one click to look up the employee
      // manually. If no external ID is configured, the cell stays as plain text.
      const fallbackSpreadsheetId = readEmployeeDataSpreadsheetId_();
      if (fallbackSpreadsheetId && enteredName) {
        const nameCell    = editedSheet.getRange(editedRow, CALL_LOG_NAME_COLUMN_NUMBER);
        const escapedName = enteredName.replace(/"/g, '""');
        nameCell.clearDataValidations();
        nameCell.setFormula(
          `=HYPERLINK("https://docs.google.com/spreadsheets/d/${fallbackSpreadsheetId}","${escapedName}")`
        );
      }
    }
    return;
  }
}


// ---------------------------------------------------------------------------
// Row Activation
// ---------------------------------------------------------------------------

/**
 * Activates a call log entry row when the manager checks its column A trigger.
 *
 * A "pending" row has only a checkbox in column A — everything else is blank.
 * Activating it performs three steps:
 *
 *   1. DATE STAMP: Removes the checkbox validation from column A and writes
 *      today's date, formatted as MM/DD/YYYY. The cell transitions from a
 *      trigger into a permanent date record.
 *
 *   2. ABSENCE CHECKBOXES: Adds Call-Out, FMLA, and No-Show checkboxes to
 *      columns E, F, and G of this row so the manager can record the type.
 *      These are added now (not at sheet creation) so blank pending rows
 *      stay visually clean.
 *
 *   3. NEXT ROW: Inserts a new trigger checkbox into column A of the row
 *      immediately below, ready for the next absence entry.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet.
 * @param {number}                             rowNumber — 1-based row that was just activated.
 */
function activateEntryRow_(sheet, rowNumber) {
  // Step 1: Replace the trigger checkbox with today's date
  const dateCell = sheet.getRange(rowNumber, 1);
  dateCell.clearDataValidations(); // removes the checkbox input rule
  dateCell.setValue(new Date());
  dateCell.setNumberFormat('MM/DD/YYYY');

  // Step 2: Add absence-type checkboxes to this row (E, F, G)
  // insertAbsenceTypeCheckboxes_() is defined in sheetGenerator.js
  insertAbsenceTypeCheckboxes_(sheet, rowNumber);

  // Step 3: Add the SEND checkbox (M) for this row so the manager can fire
  // the notification immediately when they are done filling in the entry.
  // insertNotifyCheckbox_() is defined in sheetGenerator.js
  insertNotifyCheckbox_(sheet, rowNumber);

  // Step 4: Add the next trigger checkbox on the row below
  // insertEntryTriggerCheckbox_() is defined in sheetGenerator.js
  insertEntryTriggerCheckbox_(sheet, rowNumber + 1);
}


// ---------------------------------------------------------------------------
// Employee Search
// ---------------------------------------------------------------------------

/**
 * Searches for employees matching the given name and returns all candidates.
 *
 * Returns an array so the caller can handle three cases:
 *   - length === 0 → no match; show a "not found" toast
 *   - length === 1 → unambiguous match; autofill immediately
 *   - length  >  1 → ambiguous; show a suggestion dropdown
 *
 * Search strategy — token matching:
 *   The search term is split into whitespace/comma-separated tokens. A row
 *   matches if EVERY token appears as a substring somewhere in the employee's
 *   full name (first or last). This means:
 *     "Tony Le"  → tokens ["tony","le"] → matches "Tony Le", "Anthony Le",
 *                  "Tony Leonard", but NOT "Tony Smith"
 *     "Le"       → matches every employee whose name contains "le"
 *     "Smith J"  → tokens ["smith","j"] → matches "John Smith", "Jane Smith"
 *
 *   An exact full-name match is always prioritized: if any candidate's
 *   displayName exactly equals the search term, only that candidate is
 *   returned. This ensures that selecting from the suggestion dropdown
 *   (which writes the exact displayName back to the cell) always resolves
 *   to a single result on the follow-up edit.
 *
 * Source order:
 *   1. External attendance controller ("Employee Details" sheet) — preferred.
 *   2. Local "Employee Roster" sheet — fallback if no external ID is set or
 *      the external lookup fails.
 *
 * @param {string} name — The text the clerk typed in column B (trimmed).
 * @returns {Array<{ employeeId: string, department: string, displayName: string }>}
 */
function searchEmployeesByName_(name) {
  const externalSpreadsheetId = readEmployeeDataSpreadsheetId_();

  let candidates = [];

  if (externalSpreadsheetId) {
    candidates = searchEmployeesInExternalSheet_(name, externalSpreadsheetId);
  }

  // Fall back to local roster if external is not configured or returned nothing
  if (candidates.length === 0) {
    candidates = searchEmployeesInLocalRoster_(name);
  }

  // If an exact displayName match exists among the candidates, return only that
  // one. This is what makes dropdown selection work: the clerk picks "Tony Le",
  // the cell value becomes "Tony Le", onEdit fires again, and this prioritization
  // collapses the multiple-match set back down to one.
  const searchLower = name.toLowerCase();
  const exactMatches = candidates.filter(
    candidate => candidate.displayName.toLowerCase() === searchLower
  );
  if (exactMatches.length === 1) return exactMatches;

  return candidates;
}

/**
 * Reads the master employee data spreadsheet ID from the Absence Config sheet.
 *
 * Returns null (without throwing) if the config sheet does not exist, the cell
 * is blank, or the value is not a non-empty string. The caller treats null as
 * "use the local roster fallback."
 *
 * @returns {string|null} The spreadsheet ID string, or null if not configured.
 */
function readEmployeeDataSpreadsheetId_() {
  try {
    const workbook    = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = workbook.getSheetByName(CONFIG_SHEET_NAME); // defined in config.js
    if (!configSheet) return null;

    const value = configSheet.getRange(EMPLOYEE_DATA_SPREADSHEET_ID_CELL).getValue(); // "B3"
    if (!value || typeof value !== 'string') return null;

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : null;
  } catch (error) {
    console.warn('autofill: Could not read employee data spreadsheet ID —', error);
    return null;
  }
}

/**
 * Searches the external attendance controller's "Employee Details" sheet for
 * all employees whose name contains all of the typed tokens.
 *
 * Opens the attendance controller by spreadsheet ID and scans the
 * EXTERNAL_EMPLOYEE_SHEET_NAME tab. The name is split across LAST_NAME (B)
 * and FIRST_NAME (C) columns, so the full name searched against is the
 * concatenation of both.
 *
 * Returns an empty array (not null) on any failure so callers always get an
 * array back and the clerk is never blocked by a lookup error.
 *
 * @param {string} name          — The text the clerk typed (trimmed).
 * @param {string} spreadsheetId — The Google Spreadsheet ID of the attendance controller.
 * @returns {Array<{ employeeId, department, displayName }>}
 */
function searchEmployeesInExternalSheet_(name, spreadsheetId) {
  let externalSpreadsheet;
  try {
    externalSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    // Surface the failure as a toast so the payroll clerk knows the lookup is
    // broken and can alert a manager to check the spreadsheet ID in cell B3 of
    // the "Absence Config" sheet.
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Could not access the attendance controller. Verify the spreadsheet ID in "Absence Config" cell B3. Using local roster as fallback.`,
      'Lookup Error',
      10
    );
    console.warn(
      `autofill: Could not open attendance controller (ID: ${spreadsheetId}). ` +
      `Error: ${error.message}`
    );
    return [];
  }

  const employeeSheet = externalSpreadsheet.getSheetByName(EXTERNAL_EMPLOYEE_SHEET_NAME);
  if (!employeeSheet) {
    console.warn(
      `autofill: Sheet "${EXTERNAL_EMPLOYEE_SHEET_NAME}" not found in attendance controller.`
    );
    return [];
  }

  const allRows = employeeSheet.getDataRange().getValues();
  if (allRows.length <= 1) return [];

  const columns = EXTERNAL_EMPLOYEE_COLUMNS;
  const tokens  = tokenize_(name);
  const results = [];

  for (let rowIndex = 1; rowIndex < allRows.length; rowIndex++) {
    const row       = allRows[rowIndex];
    const lastName  = String(row[columns.LAST_NAME]  || '').trim();
    const firstName = String(row[columns.FIRST_NAME] || '').trim();

    if (!lastName && !firstName) continue;

    if (tokensMatchName_(tokens, firstName, lastName)) {
      results.push({
        employeeId:  String(row[columns.EMPLOYEE_ID] || '').trim(),
        department:  String(row[columns.DEPT]        || '').trim(),
        displayName: firstName ? `${firstName} ${lastName}` : lastName,
        // Carry first/last separately so setNameHyperlink_() can build the
        // exact tab name ("Last, First - ID") to resolve the direct GID link.
        lastName:    lastName,
        firstName:   firstName,
      });
    }
  }

  return results;
}

/**
 * Searches the local "Employee Roster" sheet for all employees whose name
 * contains all of the typed tokens.
 *
 * The local roster stores the full name in a single column A, so token
 * matching is done against that combined string.
 *
 * @param {string} name — The text the clerk typed (trimmed).
 * @returns {Array<{ employeeId, department, displayName }>}
 */
function searchEmployeesInLocalRoster_(name) {
  const workbook    = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = workbook.getSheetByName(ROSTER_SHEET_NAME);

  if (!rosterSheet) {
    console.warn(`autofill: Local roster sheet "${ROSTER_SHEET_NAME}" not found.`);
    return [];
  }

  const lastRow = rosterSheet.getLastRow();
  if (lastRow < 2) return [];

  const rosterData = rosterSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const tokens     = tokenize_(name);
  const results    = [];

  for (let rowIndex = 0; rowIndex < rosterData.length; rowIndex++) {
    const fullName = String(rosterData[rowIndex][0] || '').trim();
    if (!fullName) continue;

    // Local roster has a single combined name — treat it as firstName="" lastName=fullName
    // so tokensMatchName_ searches the full string.
    if (tokensMatchName_(tokens, fullName, '')) {
      results.push({
        employeeId:  String(rosterData[rowIndex][1] || '').trim(),
        department:  String(rosterData[rowIndex][2] || '').trim(),
        displayName: fullName,
      });
    }
  }

  return results;
}

/**
 * Splits a search string into lowercase tokens, stripping commas.
 * "Tony Le" → ["tony", "le"]
 * "Smith, J" → ["smith", "j"]
 *
 * @param {string} searchTerm
 * @returns {string[]}
 */
function tokenize_(searchTerm) {
  return searchTerm
    .toLowerCase()
    .replace(/,/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
}

/**
 * Returns true if every token appears as a substring in the employee's
 * first name, last name, or the combined "first last" string.
 *
 * This is what makes partial search work:
 *   tokens=["tony","le"], firstName="Tony", lastName="Le"
 *     → "tony" in "tony" ✓, "le" in "le" ✓ → MATCH
 *   tokens=["tony","le"], firstName="Anthony", lastName="Le"
 *     → "tony" in "anthony" ✓ (substring), "le" in "le" ✓ → MATCH
 *   tokens=["tony","le"], firstName="Tony", lastName="Smith"
 *     → "tony" in "tony" ✓, "le" in "smith" ✗ → NO MATCH
 *
 * @param {string[]} tokens    — Lowercased search tokens.
 * @param {string}   firstName — Employee first name (may be empty for local roster).
 * @param {string}   lastName  — Employee last name.
 * @returns {boolean}
 */
function tokensMatchName_(tokens, firstName, lastName) {
  const first    = firstName.toLowerCase();
  const last     = lastName.toLowerCase();
  const combined = `${first} ${last}`.trim();

  return tokens.every(token =>
    first.includes(token) || last.includes(token) || combined.includes(token)
  );
}


// ---------------------------------------------------------------------------
// Sheet Write Helpers
// ---------------------------------------------------------------------------

/**
 * Writes the autofilled Employee ID and Department to their columns on the
 * call log entry row, and converts the employee name in column B into a
 * clickable hyperlink pointing to the master employee spreadsheet.
 *
 * The name hyperlink is set using a HYPERLINK formula so the cell displays
 * the employee's name as normal text but is clickable. Payroll clerks can
 * click the name to open the master spreadsheet directly without leaving
 * the call log. If no external spreadsheet ID is configured, the name is
 * left as plain text and only the ID and department are written.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet       — The call log sheet being edited.
 * @param {number}                             rowNumber   — The 1-based row of the entry.
 * @param {{ employeeId: string, department: string }} employeeData — Data to write.
 * @param {string}                             employeeName — The name as entered, used for the hyperlink label.
 */
function writeAutofillFields_(sheet, rowNumber, employeeData, employeeName) {
  // Write Employee ID (column C) and Department (column D) in one call
  sheet.getRange(rowNumber, 3, 1, 2).setValues([[
    employeeData.employeeId,
    employeeData.department,
  ]]);

  // Remove any suggestion dropdown validation that may be on the name cell before
  // writing the HYPERLINK formula. If showSuggestionDropdown_() was called for
  // this row, column B still has requireValueInList validation. Clearing it first
  // ensures setFormula() is not competing with an active validation rule.
  sheet.getRange(rowNumber, CALL_LOG_NAME_COLUMN_NUMBER).clearDataValidations();

  // Convert the name in column B into a hyperlink pointing directly to the
  // employee's own tab in the attendance controller.
  const spreadsheetId = readEmployeeDataSpreadsheetId_();
  if (spreadsheetId && employeeName) {
    setNameHyperlink_(
      sheet,
      rowNumber,
      employeeName,
      spreadsheetId,
      employeeData.lastName   || '',
      employeeData.firstName  || '',
      employeeData.employeeId || ''
    );
  }
}

/**
 * Sets the employee name cell (column B) as a HYPERLINK formula that links
 * directly to the employee's own sheet tab inside the attendance controller.
 *
 * The employee's tab name is constructed from EMPLOYEE_TAB_NAME_FORMAT in
 * config.js (e.g. "Le, Tony - 12345"). If a sheet with that name is found
 * in the attendance controller, the URL includes its GID so the clerk lands
 * directly on that tab:
 *   https://docs.google.com/spreadsheets/d/{ID}/edit#gid={sheetGID}
 *
 * If the tab cannot be found (employee has no individual sheet yet, or the
 * naming convention differs), the link falls back to the spreadsheet root
 * so the clerk still gets to the right workbook with one click.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet         — The call log sheet.
 * @param {number}                             rowNumber     — 1-based row of the entry.
 * @param {string}                             employeeName  — Display name for the link label.
 * @param {string}                             spreadsheetId — Attendance controller spreadsheet ID.
 * @param {string}                             lastName      — Employee last name (for tab name).
 * @param {string}                             firstName     — Employee first name (for tab name).
 * @param {string}                             employeeId    — Employee number (for tab name).
 */
function setNameHyperlink_(sheet, rowNumber, employeeName, spreadsheetId, lastName, firstName, employeeId) {
  // Build the expected tab name using the format defined in config.js.
  // e.g. EMPLOYEE_TAB_NAME_FORMAT = "{LAST}, {FIRST} - {ID}" → "Le, Tony - 12345"
  const expectedTabName = EMPLOYEE_TAB_NAME_FORMAT
    .replace('{LAST}',  lastName  || '')
    .replace('{FIRST}', firstName || '')
    .replace('{ID}',    employeeId || '');

  // Attempt to resolve the GID of the employee's tab.
  // GAS caches SpreadsheetApp.openById() within a script execution, so this
  // does not incur a second network round-trip if the search already opened it.
  let url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
  try {
    const externalWorkbook = SpreadsheetApp.openById(spreadsheetId);
    const employeeTab      = externalWorkbook.getSheetByName(expectedTabName);

    if (employeeTab) {
      // Direct link to the employee's specific tab
      url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${employeeTab.getSheetId()}`;
    } else {
      console.log(
        `autofill: Tab "${expectedTabName}" not found in attendance controller. ` +
        `Linking to spreadsheet root instead.`
      );
    }
  } catch (error) {
    console.warn(`autofill: Could not resolve employee tab GID — ${error.message}`);
  }

  // In Sheets formula syntax, a literal " inside a string is escaped as ""
  const escapedName = employeeName.replace(/"/g, '""');

  sheet.getRange(rowNumber, CALL_LOG_NAME_COLUMN_NUMBER)
    .setFormula(`=HYPERLINK("${url}","${escapedName}")`);
}

/**
 * Sets a dropdown on the name cell (column B) containing the display names of
 * all suggestion candidates, so the clerk can pick the correct person without
 * retyping.
 *
 * After the clerk selects a name from the dropdown, onEdit fires again with
 * that exact displayName as the cell value. searchEmployeesByName_() will find
 * an exact match, autofill runs, and the dropdown validation is replaced by the
 * HYPERLINK formula — so the clerk never sees the dropdown again for that row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet       — The call log sheet.
 * @param {number}                             rowNumber   — 1-based row to set the dropdown on.
 * @param {Array<{ displayName: string }>}     candidates  — The matching employees to list.
 */
function showSuggestionDropdown_(sheet, rowNumber, candidates) {
  const nameList   = candidates.map(candidate => candidate.displayName);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(nameList, true) // true = show dropdown arrow
    .setAllowInvalid(true)              // allow the clerk to still type freely
    .build();

  sheet.getRange(rowNumber, CALL_LOG_NAME_COLUMN_NUMBER)
    .setDataValidation(validation);
}

/**
 * Clears the autofilled fields when the manager deletes the name in column B.
 *
 * Clears Employee ID (C), Department (D), any HYPERLINK formula, and any
 * suggestion dropdown validation from column B so the cell is a plain empty
 * input ready for a new name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet being edited.
 * @param {number}                             rowNumber — The 1-based row number to clear.
 */
function clearAutofillFields_(sheet, rowNumber) {
  const nameCell = sheet.getRange(rowNumber, CALL_LOG_NAME_COLUMN_NUMBER);
  nameCell.setFormula('');           // remove any HYPERLINK formula
  nameCell.clearDataValidations();   // remove any suggestion dropdown

  // Clear Employee ID and Department
  sheet.getRange(rowNumber, 3, 1, 2).clearContent();
}


// ---------------------------------------------------------------------------
// Manual Notification Send
// ---------------------------------------------------------------------------

/**
 * Sends an immediate absence notification for a single call log row.
 *
 * Called when the manager checks the SEND checkbox in column M. This bypasses
 * the 15-minute time-driven window so the department manager is notified the
 * moment the call log entry is complete — no delay, no waiting for the next
 * scheduled trigger run.
 *
 * After a successful send the checkbox is replaced with "Sent HH:MM AM" so:
 *   1. The row is visually confirmed as notified at a glance.
 *   2. The time-driven trigger in digestEngine.js will skip this row because
 *      the NOTIFY cell is now a string, not a boolean.
 *
 * If the row is missing required data (employee name, at least one absence type),
 * the checkbox is reverted to unchecked and a toast explains what is missing.
 * On any sending failure the checkbox is also reverted and an error toast is shown.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet     — The call log sheet.
 * @param {number}                             rowNumber — 1-based row whose SEND was checked.
 */
function sendEntryNotification_(sheet, rowNumber) {
  const timeZone = Session.getScriptTimeZone();
  const columns  = CALL_LOG_COLUMNS;
  const notifyCell = sheet.getRange(rowNumber, CALL_LOG_NOTIFY_COLUMN_NUMBER);

  // Read all columns for this row in a single call
  const rowData = sheet.getRange(rowNumber, 1, 1, columns.TOTAL_COLUMNS_TO_READ).getValues()[0];

  // --- Validate: employee name is required ---
  const employeeName = String(rowData[columns.NAME] || '').trim();
  if (!employeeName) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Enter an employee name in column B before sending.',
      'Missing Name',
      5
    );
    notifyCell.setValue(false);
    return;
  }

  // --- Validate: at least one absence type must be selected ---
  const isCallout = coerceToBool_(rowData[columns.IS_CALLOUT]); // defined in dataIngestion.js
  const isFmla    = coerceToBool_(rowData[columns.IS_FMLA]);
  const isNoShow  = coerceToBool_(rowData[columns.IS_NOSHOW]);

  if (!isCallout && !isFmla && !isNoShow) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Check at least one absence type (Call-Out, FMLA, or No-Show) before sending.',
      'Missing Absence Type',
      5
    );
    notifyCell.setValue(false);
    return;
  }

  // --- Build absence reason ---
  let absenceReason = 'Unknown';
  if (isCallout)     absenceReason = 'Call-Out';
  else if (isFmla)   absenceReason = 'FMLA';
  else if (isNoShow) absenceReason = 'No-Show';

  // --- Resolve the call time ---
  // Use TIME_CALLED (column H) if it parses successfully; fall back to now.
  // A narrow 1-minute window anchored to the call time is passed to the email
  // builder so the "Window" line in the email body shows a meaningful time.
  const now          = new Date();
  const narrowWindow = { start: new Date(now.getTime() - 60000), end: now };
  let   calledAt     = now;

  const timeRaw = rowData[columns.TIME_CALLED];
  if (timeRaw) {
    const parsed = parseTimeToMilliseconds_(timeRaw, narrowWindow); // defined in timeWindow.js
    if (parsed) calledAt = new Date(parsed);
  }

  // --- Assemble the AbsenceRecord ---
  const record = {
    rowNumber:       rowNumber,
    employeeName:    employeeName,
    employeeId:      String(rowData[columns.EMPLOYEE_ID]     || 'Unknown').trim(),
    isAbsence:       true,
    absenceReason:   absenceReason,
    department:      String(rowData[columns.DEPT]            || '').trim(),
    employeeComment: String(rowData[columns.COMMENT]         || '').trim(),
    scheduledShift:  String(rowData[columns.SCHEDULED_SHIFT] || '').trim(),
    calledAt:        calledAt,
  };

  const emailWindow = { start: calledAt, end: now };

  // --- Send and stamp ---
  try {
    sendDepartmentDigests_([record], emailWindow, timeZone); // defined in notifier.js
    const sentTime = Utilities.formatDate(now, timeZone, 'h:mm a');
    notifyCell.clearDataValidations().setValue(`Sent ${sentTime}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Notification sent for ${employeeName}.`,
      'Sent',
      4
    );
  } catch (error) {
    console.error(`autofill: Failed to send notification for row ${rowNumber} — ${error.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Failed to send notification. Check the Apps Script logs.',
      'Send Error',
      8
    );
    notifyCell.setValue(false); // revert so the manager can try again
  }
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


// ---------------------------------------------------------------------------
// Diagnostics (run manually from Apps Script editor to troubleshoot lookups)
// ---------------------------------------------------------------------------

/**
 * Diagnostic helper — run this function directly from the Apps Script editor
 * (not via a trigger) to verify the employee lookup configuration.
 *
 * What it checks and logs:
 *   1. Whether the spreadsheet ID in "Absence Config" B3 is set.
 *   2. Whether the attendance controller spreadsheet can be opened.
 *   3. Every sheet name in that workbook (so you can spot an exact-name mismatch).
 *   4. Whether the expected EXTERNAL_EMPLOYEE_SHEET_NAME sheet exists.
 *   5. The first 5 data rows of that sheet (so you can confirm column order).
 *   6. A live search for the name you supply in TEST_SEARCH_NAME below.
 *
 * HOW TO USE:
 *   1. Open the Apps Script editor (Extensions → Apps Script).
 *   2. Change TEST_SEARCH_NAME to a real employee name from the roster.
 *   3. Select this function in the function dropdown and click Run.
 *   4. Open View → Logs (or Cmd+Enter) to see the output.
 */
function debugEmployeeLookup() {
  const TEST_SEARCH_NAME = 'Smith'; // ← change to any real last name from your roster

  // ── Step 1: Read spreadsheet ID ──────────────────────────────────────────
  const spreadsheetId = readEmployeeDataSpreadsheetId_();
  if (!spreadsheetId) {
    console.error('DEBUG: No spreadsheet ID found in "Absence Config" cell B3. Aborting.');
    return;
  }
  console.log(`DEBUG: Spreadsheet ID = "${spreadsheetId}"`);

  // ── Step 2: Open the workbook ─────────────────────────────────────────────
  let wb;
  try {
    wb = SpreadsheetApp.openById(spreadsheetId);
    console.log(`DEBUG: Opened workbook "${wb.getName()}" successfully.`);
  } catch (err) {
    console.error(`DEBUG: Failed to open spreadsheet — ${err.message}`);
    return;
  }

  // ── Step 3: List every sheet name ─────────────────────────────────────────
  const allSheets = wb.getSheets();
  console.log(`DEBUG: ${allSheets.length} sheet(s) found in workbook:`);
  allSheets.forEach((s, i) => console.log(`  [${i}] "${s.getName()}"`));

  // ── Step 4: Find the expected employee sheet ──────────────────────────────
  const employeeSheet = wb.getSheetByName(EXTERNAL_EMPLOYEE_SHEET_NAME);
  if (!employeeSheet) {
    console.error(
      `DEBUG: Sheet "${EXTERNAL_EMPLOYEE_SHEET_NAME}" NOT FOUND. ` +
      `Check the sheet name above and update EXTERNAL_EMPLOYEE_SHEET_NAME in config.js to match exactly.`
    );
    return;
  }
  console.log(`DEBUG: Found sheet "${EXTERNAL_EMPLOYEE_SHEET_NAME}". Rows: ${employeeSheet.getLastRow()}`);

  // ── Step 5: Show the first 5 data rows to confirm column layout ───────────
  const previewRows = Math.min(6, employeeSheet.getLastRow()); // row 1 = header + up to 5 data rows
  if (previewRows >= 1) {
    const preview = employeeSheet.getRange(1, 1, previewRows, 6).getValues();
    console.log('DEBUG: First rows of employee sheet (columns A–F):');
    preview.forEach((row, i) => console.log(`  Row ${i + 1}: ${JSON.stringify(row)}`));
  }

  // ── Step 6: Run the token search ─────────────────────────────────────────
  console.log(`DEBUG: Searching for "${TEST_SEARCH_NAME}"...`);
  const results = searchEmployeesInExternalSheet_(TEST_SEARCH_NAME, spreadsheetId);
  if (results.length === 0) {
    console.warn(`DEBUG: Search returned 0 results for "${TEST_SEARCH_NAME}".`);
  } else {
    console.log(`DEBUG: ${results.length} result(s) found:`);
    results.forEach((r, i) =>
      console.log(`  [${i}] displayName="${r.displayName}" | id="${r.employeeId}" | dept="${r.department}"`)
    );
  }
}
