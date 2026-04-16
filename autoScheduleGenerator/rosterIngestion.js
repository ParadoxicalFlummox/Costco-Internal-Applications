/**
 * rosterIngestion.js — Syncs employees from an external master spreadsheet into the local Roster sheet.
 * VERSION 1.2.0
 *
 * This file handles the entire "bring employee data into the workbook" workflow:
 *   1. The manager opens the Ingestion sheet and types a Google Spreadsheet ID.
 *   2. The script reads that spreadsheet and populates a department dropdown.
 *   3. The manager selects a department and triggers a sync.
 *   4. The script compares the source employees against the current Roster,
 *      adds new ones, skips duplicates (matched by Employee ID), and reports the result.
 *
 * SEPARATION OF CONCERNS:
 * Each step above is its own function. The orchestrating function syncRosterFromSource()
 * calls these functions in sequence but contains no logic of its own. This means if
 * the sync produces wrong data, the bug can be traced to exactly one function:
 *   - Wrong employees returned  → fetchEmployeesFromSource()
 *   - Wrong duplicate detection → deduplicateAgainstRoster()
 *   - Wrong data written        → writeNewEmployeesToRoster()
 *   - Wrong inputs read         → readIngestionInputs()
 */


/**
 * Creates the Ingestion sheet with its layout if it does not already exist.
 *
 * This function is called once during first-run setup. It creates the labels,
 * input cells, and instructional text that guide the manager through the
 * roster sync process. After setup, the manager fills in the spreadsheet ID
 * and clicks "Sync Roster" from the Schedule Admin menu.
 */
function setupIngestionSheet() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  let ingestionSheet = workbook.getSheetByName(SHEET_NAMES.INGESTION);

  // If the sheet already exists and the manager has entered a source spreadsheet ID,
  // skip setup entirely. clearContents() would destroy the ID and the department
  // selection they have already configured, which is the most painful accidental loss.
  // An empty or placeholder value means setup has not been completed and is safe to run.
  if (ingestionSheet) {
    const existingId = ingestionSheet
      .getRange(INGESTION_CELL.SOURCE_SPREADSHEET_ID)
      .getValue()
      .toString()
      .trim();
    const isConfigured = existingId !== "" &&
      existingId !== "— enter spreadsheet ID first —" &&
      !existingId.startsWith("←");
    if (isConfigured) {
      return;
    }
  }

  if (!ingestionSheet) {
    ingestionSheet = workbook.insertSheet(SHEET_NAMES.INGESTION);
  }

  // Clear any existing content. At this point we know B3 is empty or a placeholder,
  // so no real configuration is being lost.
  ingestionSheet.clearContents();
  ingestionSheet.clearFormats();

  // --- Row 1: Title header ---
  const titleCell = ingestionSheet.getRange("A1:B1");
  titleCell.merge();
  titleCell.setValue("Schedule Roster Sync");
  titleCell.setFontWeight("bold");
  titleCell.setFontSize(14);
  titleCell.setBackground(COLORS.HEADER_BG);
  titleCell.setFontColor(COLORS.HEADER_TEXT);

  // --- Row 3: Source spreadsheet ID input ---
  ingestionSheet.getRange("A3").setValue("Source Spreadsheet ID:");
  ingestionSheet.getRange("A3").setFontWeight("bold");
  ingestionSheet.getRange("B3").setValue("").setBackground("#FFFFFF");
  ingestionSheet.getRange("B3").setNote(
    "Paste the ID from the URL of your master employee spreadsheet.\n" +
    "The ID is the long string of characters between /d/ and /edit in the URL.\n" +
    "Example: https://docs.google.com/spreadsheets/d/[THIS_IS_THE_ID]/edit"
  );

  // --- Row 4: Department dropdown (initially empty; populated when ID is entered) ---
  ingestionSheet.getRange("A4").setValue("Department:");
  ingestionSheet.getRange("A4").setFontWeight("bold");
  ingestionSheet.getRange("B4").setValue("— enter spreadsheet ID first —");
  ingestionSheet.getRange("B4").setFontColor("#999999");

  // --- Row 6: Instructions for triggering the sync ---
  ingestionSheet.getRange("A6").setValue(
    "After selecting a department, use the Schedule Admin menu → Sync Roster."
  );
  ingestionSheet.getRange("A6").setFontStyle("italic");
  ingestionSheet.getRange("A6:B6").merge();

  // --- Rows 8–10: Result area (script writes here after sync) ---
  ingestionSheet.getRange("A8").setValue("Last sync status:");
  ingestionSheet.getRange("A9").setValue("Employees added:");
  ingestionSheet.getRange("A10").setValue("Employees skipped (already on roster):");
  ingestionSheet.getRange("A8:A10").setFontWeight("bold");

  // Set column widths so labels and values are readable.
  ingestionSheet.setColumnWidth(1, 280);
  ingestionSheet.setColumnWidth(2, 380);
}


/**
 * Reads the source spreadsheet ID and selected department from the Ingestion sheet.
 *
 * This function is the sole reader of the Ingestion sheet's input cells. Separating
 * the read from the sync logic means that if the wrong value is being used for the
 * spreadsheet ID or department, this is the only function to check.
 *
 * @returns {{ sourceSpreadsheetId: string, departmentName: string }}
 */
function readIngestionInputs() {
  const ingestionSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.INGESTION);

  if (!ingestionSheet) {
    throw new Error(
      "Ingestion sheet not found. Run \"Setup Sheets\" from the Schedule Admin menu."
    );
  }

  const sourceSpreadsheetId = ingestionSheet
    .getRange(INGESTION_CELL.SOURCE_SPREADSHEET_ID)
    .getValue()
    .toString()
    .trim();

  const departmentName = ingestionSheet
    .getRange(INGESTION_CELL.DEPARTMENT)
    .getValue()
    .toString()
    .trim();

  if (!sourceSpreadsheetId || sourceSpreadsheetId === "") {
    throw new Error(
      "No source spreadsheet ID found in Ingestion cell " +
      INGESTION_CELL.SOURCE_SPREADSHEET_ID + ". Enter the spreadsheet ID before syncing."
    );
  }

  if (!departmentName || departmentName === "" || departmentName === "— enter spreadsheet ID first —") {
    throw new Error(
      "No department selected. Enter the spreadsheet ID first, then choose a department from the dropdown."
    );
  }

  return {
    sourceSpreadsheetId: sourceSpreadsheetId,
    departmentName: departmentName,
  };
}


/**
 * Reads the source spreadsheet and returns all employees matching the given department.
 *
 * This function opens an external Google Spreadsheet by ID and reads employee data
 * from its first sheet. It expects the source sheet to have the following columns:
 *   Column A — Employee Name
 *   Column B — Employee ID
 *   Column F — Hire Date
 *   Column C — Department
 *
 * If the source spreadsheet cannot be accessed (wrong ID, insufficient permissions),
 * this function throws a descriptive error that is caught by the orchestrator and
 * displayed to the manager.
 *
 * @param {string} sourceSpreadsheetId — The Google Spreadsheet ID of the master employee list.
 * @param {string} departmentName      — The department name to filter by (exact match).
 * @returns {Array<{name: string, employeeId: string, hireDate: Date}>}
 *   An array of employee objects matching the department.
 */
function fetchEmployeesFromSource(sourceSpreadsheetId, departmentName) {
  let sourceSpreadsheet;

  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  } catch (error) {
    throw new Error(
      "Could not open the source spreadsheet (ID: " + sourceSpreadsheetId + "). " +
      "Check that the ID is correct and that this script has been granted access. " +
      "Technical detail: " + error.message
    );
  }

  // Read from the first sheet of the source spreadsheet.
  // If the master list uses a specific named sheet, this can be updated to
  // sourceSpreadsheet.getSheetByName("YourSheetNameHere").
  const sourceSheet = sourceSpreadsheet.getSheets()[0];
  const allRows = sourceSheet.getDataRange().getValues();

  if (allRows.length <= 1) {
    // The source sheet only has a header row (or is empty) — nothing to sync.
    return [];
  }

  // Skip row 0 (the header row) and filter by department.
  const matchingEmployees = [];

  allRows.slice(1).forEach(function (row) {
    const employeeName = row[0];
    const employeeId = row[1];
    const hireDate = row[5];
    const department = row[2];

    // Skip rows where the department column is blank or does not match.
    if (!department || department.toString().trim() !== departmentName) {
      return;
    }

    // Skip rows with missing required fields and log them so the manager
    // knows to fix the source data.
    if (!employeeName || !employeeId) {
      Logger.log(
        "WARNING: Skipping a row in the source spreadsheet because Employee Name or " +
        "Employee ID is blank. Row data: " + JSON.stringify(row)
      );
      return;
    }

    matchingEmployees.push({
      name: employeeName.toString().trim(),
      employeeId: employeeId.toString().trim(),
      hireDate: hireDate instanceof Date ? hireDate : new Date(hireDate),
    });
  });

  return matchingEmployees;
}


/**
 * Compares a list of source employees against the existing Roster and separates them into
 * employees to add and employees to skip (because they are already on the Roster).
 *
 * Deduplication is performed by Employee ID, not by name. Employee IDs are stable
 * identifiers — names can have typos or change, but IDs do not. This prevents an
 * employee from being added twice if their name is formatted slightly differently
 * in the source spreadsheet.
 *
 * @param {Array<{name: string, employeeId: string, hireDate: Date}>} sourceEmployees
 *   The employees fetched from the source spreadsheet.
 * @returns {{ employeesToAdd: Array, skippedEmployees: Array }}
 */
function deduplicateAgainstRoster(sourceEmployees) {
  const existingEmployeeIds = getExistingRosterEmployeeIds();

  const employeesToAdd = [];
  const skippedEmployees = [];

  sourceEmployees.forEach(function (employee) {
    if (existingEmployeeIds.has(employee.employeeId)) {
      // This employee is already on the Roster — skip them to avoid creating a duplicate row.
      skippedEmployees.push(employee);
    } else {
      employeesToAdd.push(employee);
    }
  });

  return {
    employeesToAdd: employeesToAdd,
    skippedEmployees: skippedEmployees,
  };
}


/**
 * Reads the Employee ID column from the Roster sheet and returns the values as a Set.
 *
 * A Set is used (rather than an array) so that membership checks during deduplication
 * are O(1) lookups instead of O(n) array scans. For large rosters this is a meaningful
 * performance difference.
 *
 * @returns {Set<string>} The Employee IDs currently on the Roster sheet.
 */
function getExistingRosterEmployeeIds() {
  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    return new Set();
  }

  const lastRow = rosterSheet.getLastRow();

  if (lastRow < ROSTER_DATA_START_ROW) {
    // The roster only has a header row (or is empty) — no existing IDs.
    return new Set();
  }

  // Read only the Employee ID column to avoid loading the entire roster into memory.
  const employeeIdValues = rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.EMPLOYEE_ID, lastRow - ROSTER_DATA_START_ROW + 1, 1)
    .getValues();

  const existingIds = new Set();

  employeeIdValues.forEach(function (row) {
    const employeeId = row[0];
    if (employeeId && employeeId.toString().trim() !== "") {
      existingIds.add(employeeId.toString().trim());
    }
  });

  return existingIds;
}


/**
 * Writes a list of new employee records to the Roster sheet, one row per employee.
 *
 * Each new employee is written with default values:
 *   - Status: "PT" (part-time) — the manager should update this manually for FT employees.
 *   - All preference fields: blank.
 *   - Seniority rank: 0 (placeholder; refreshAllSeniorityRanks() must be called after this).
 *
 * This function performs only writes — it does not read the sheet, calculate seniority,
 * or apply validation. Those concerns belong to other functions.
 *
 * @param {Array<{name: string, employeeId: string, hireDate: Date}>} newEmployees
 *   The employees to write to the Roster sheet.
 */
function writeNewEmployeesToRoster(newEmployees) {
  if (newEmployees.length === 0) {
    return;
  }

  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    throw new Error("Roster sheet not found. Run \"Setup Sheets\" from the Schedule Admin menu.");
  }

  // Determine the first empty row below any existing data.
  const firstEmptyRow = Math.max(rosterSheet.getLastRow() + 1, ROSTER_DATA_START_ROW);

  // Build all rows in memory first, then write them in a single setValues() call.
  // This replaces N × 13 individual setValue() calls with one batch write.
  const rowsToWrite = newEmployees.map(function(employee) {
    return buildEmployeeRosterRow_(employee);
  });

  if (rowsToWrite.length > 0) {
    rosterSheet
      .getRange(firstEmptyRow, 1, rowsToWrite.length, ROSTER_COLUMN.PRIMARY_ROLE)
      .setValues(rowsToWrite);
  }
}


/**
 * Builds and returns a single Roster row array for one employee with default values.
 *
 * Returns a 1D array whose indices correspond to ROSTER_COLUMN positions (1-indexed columns
 * mapped to 0-indexed array positions). The caller (writeNewEmployeesToRoster) collects these
 * arrays and writes them all in one setValues() batch call.
 *
 * @param {{name: string, employeeId: string, hireDate: Date}} employee — The employee to write.
 * @returns {Array} A 13-element array matching columns A–M of the Roster sheet.
 */
function buildEmployeeRosterRow_(employee) {
  // Validate the hire date before writing. An invalid date would produce a
  // nonsensical seniority rank and is better caught here with a clear warning.
  const hireDateValue = employee.hireDate instanceof Date && !isNaN(employee.hireDate.getTime())
    ? employee.hireDate
    : null;

  if (!hireDateValue) {
    Logger.log(
      "WARNING: Employee \"" + employee.name + "\" (ID: " + employee.employeeId + ") " +
      "has an invalid or missing hire date. A placeholder hire date of today will be used. " +
      "Update column C for this employee in the Roster sheet after sync."
    );
  }

  // Array indices are (ROSTER_COLUMN value - 1) since ROSTER_COLUMN is 1-indexed.
  // Columns D–I (preferences, qualified shifts, vacation dates) are left blank —
  // the manager fills these in after sync.
  return [
    employee.name,           // A — NAME
    employee.employeeId,     // B — EMPLOYEE_ID
    hireDateValue || new Date(), // C — HIRE_DATE
    "PT",                    // D — STATUS (default to PT; manager updates to FT if needed)
    "",                      // E — DAY_OFF_PREF_ONE
    "",                      // F — DAY_OFF_PREF_TWO
    "",                      // G — PREFERRED_SHIFT
    "",                      // H — QUALIFIED_SHIFTS
    "",                      // I — VACATION_DATES
    0,                       // J — SENIORITY_RANK (placeholder; refreshAllSeniorityRanks() recalculates)
    "",                      // K — DEPARTMENT
    "",                      // L — QUALIFIED_DEPARTMENTS
    "",                      // M — PRIMARY_ROLE
  ];
}


/**
 * Reads the source spreadsheet and writes a department dropdown to Ingestion cell B4.
 *
 * This function is triggered by onEdit() when the manager changes the value of
 * Ingestion cell B3 (the spreadsheet ID). It opens the source spreadsheet, reads
 * the unique department values from column C, and replaces the B4 cell's data
 * validation with a dropdown containing those values.
 *
 * If the spreadsheet cannot be opened (bad ID, no access), the dropdown is cleared
 * and a friendly error message is written to B4 instead of crashing.
 *
 * @param {string} sourceSpreadsheetId — The Google Spreadsheet ID entered in Ingestion B3.
 */
function populateDepartmentDropdown(sourceSpreadsheetId) {
  const ingestionSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.INGESTION);

  const departmentCell = ingestionSheet.getRange(INGESTION_CELL.DEPARTMENT);

  // Clear the existing dropdown and reset the cell appearance while we fetch data.
  departmentCell.clearDataValidations();
  departmentCell.setValue("Loading departments...");
  departmentCell.setFontColor("#999999");

  let sourceSpreadsheet;

  try {
    sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId.trim());
  } catch (error) {
    departmentCell.setValue("Error: Could not open spreadsheet. Check the ID.");
    departmentCell.setFontColor("#CC0000");
    Logger.log("populateDepartmentDropdown error: " + error.message);
    return;
  }

  const sourceSheet = sourceSpreadsheet.getSheets()[0];
  const allRows = sourceSheet.getDataRange().getValues();

  // Collect unique, non-blank department names from column C (index 2).
  // This matches the column layout used by fetchEmployeesFromSource():
  //   Column A (0) — Employee Name
  //   Column B (1) — Employee ID
  //   Column C (2) — Department
  //   Column F (5) — Hire Date
  const uniqueDepartments = new Set();

  allRows.slice(1).forEach(function (row) {
    const department = row[2];
    if (department && department.toString().trim() !== "") {
      uniqueDepartments.add(department.toString().trim());
    }
  });

  const departmentList = Array.from(uniqueDepartments).sort();

  if (departmentList.length === 0) {
    departmentCell.setValue("No departments found in column C of source sheet.");
    departmentCell.setFontColor("#CC0000");
    return;
  }

  // Write the dropdown using data validation. The manager picks one value from the list.
  const dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(departmentList, true)
    .setAllowInvalid(false)
    .build();

  departmentCell.setDataValidation(dropdownRule);
  departmentCell.setValue(departmentList[0]);
  departmentCell.setFontColor("#000000");
}


/**
 * Writes the sync result summary to the Ingestion sheet status cells.
 *
 * This is separated from the orchestrator so that the result-writing concern
 * is isolated. If the status cells show wrong information, this is the only
 * function to check.
 *
 * @param {{ employeesAdded: number, employeesSkipped: number, statusMessage: string }} syncResult
 */
function writeSyncResultToIngestionSheet(syncResult) {
  const ingestionSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.INGESTION);

  if (!ingestionSheet) {
    return;
  }

  ingestionSheet.getRange(INGESTION_CELL.SYNC_STATUS).setValue(syncResult.statusMessage);
  ingestionSheet.getRange(INGESTION_CELL.EMPLOYEES_ADDED).setValue(syncResult.employeesAdded);
  ingestionSheet.getRange(INGESTION_CELL.EMPLOYEES_SKIPPED).setValue(syncResult.employeesSkipped);
}


/**
 * Applies data validation dropdowns to all data rows in the Roster sheet.
 *
 * This function sets up the dropdowns that guide managers when editing the Roster:
 *   - Column D (Status): "FT" or "PT"
 *   - Columns E & F (Day Off Preferences): the seven day names
 *   - Column G (Preferred Shift): shift names read from the Settings sheet
 *
 * It is called after every sync and after refreshing shift names via the menu.
 * Running it multiple times is safe — it replaces any existing validation rules.
 *
 * @param {Sheet} rosterSheet — The Roster sheet object.
 */
function applyRosterValidation(rosterSheet) {
  const lastRow = rosterSheet.getLastRow();

  if (lastRow < ROSTER_DATA_START_ROW) {
    // No data rows to validate yet.
    return;
  }

  const dataRowCount = lastRow - ROSTER_DATA_START_ROW + 1;

  // --- Status dropdown: "FT" or "PT" ---
  const statusValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["FT", "PT"], true)
    .setAllowInvalid(false)
    .setHelpText("Select FT for full-time (40 hrs/week) or PT for part-time (24–35 hrs/week).")
    .build();

  rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.STATUS, dataRowCount, 1)
    .setDataValidation(statusValidationRule);

  // --- Day off preference dropdowns: Monday through Sunday (or blank) ---
  const dayNameValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["", ...DAY_NAMES_IN_ORDER], true)
    .setAllowInvalid(false)
    .setHelpText("Select a preferred day off, or leave blank if no preference.")
    .build();

  rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.DAY_OFF_PREF_ONE, dataRowCount, 1)
    .setDataValidation(dayNameValidationRule);

  rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.DAY_OFF_PREF_TWO, dataRowCount, 1)
    .setDataValidation(dayNameValidationRule);

  // --- Preferred shift dropdown: populated from Settings shift names ---
  const shiftNames = readShiftNamesFromSettings();

  if (shiftNames.length > 0) {
    const shiftValidationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["", ...shiftNames], true)
      .setAllowInvalid(false)
      .setHelpText("Select the employee's preferred shift. Must match a shift name in the Settings sheet.")
      .build();

    rosterSheet
      .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.PREFERRED_SHIFT, dataRowCount, 1)
      .setDataValidation(shiftValidationRule);
  }

  // The Qualified Shifts column (H) uses free text with a cell note listing valid values,
  // because it is a comma-separated list and GAS dropdown validation does not support
  // multi-select. Add a note to the header cell to guide data entry.
  const qualifiedShiftsHeaderNote =
    "Enter comma-separated shift names this employee is trained to work.\n" +
    "Valid shift names: " + shiftNames.join(", ") + "\n" +
    "Example: Morning, Mid, Closing";

  rosterSheet
    .getRange(ROSTER_DATA_START_ROW, ROSTER_COLUMN.QUALIFIED_SHIFTS, dataRowCount, 1)
    .setNote(qualifiedShiftsHeaderNote);
}


/**
 * Orchestrates the full roster sync: reads inputs → fetches source → deduplicates → writes → validates.
 *
 * This function contains no logic of its own. It calls the other functions in the correct
 * order, passes their outputs as inputs to the next function, and assembles the result.
 * If anything goes wrong, the error thrown by the failing worker function will include a
 * descriptive message that points directly at the problem.
 *
 * @returns {{ employeesAdded: number, employeesSkipped: number, statusMessage: string }}
 */
function syncRosterFromSource() {
  // Step 1: Read what the manager has entered on the Ingestion sheet.
  const inputs = readIngestionInputs();

  // Step 2: Fetch matching employees from the external source spreadsheet.
  const sourceEmployees = fetchEmployeesFromSource(
    inputs.sourceSpreadsheetId,
    inputs.departmentName
  );

  if (sourceEmployees.length === 0) {
    const emptyResult = {
      employeesAdded: 0,
      employeesSkipped: 0,
      statusMessage: "No employees found for department \"" + inputs.departmentName + "\". " +
        "Check that the department name matches exactly.",
    };
    writeSyncResultToIngestionSheet(emptyResult);
    return emptyResult;
  }

  // Step 3: Separate new employees from those already on the Roster.
  const deduplicationResult = deduplicateAgainstRoster(sourceEmployees);

  // Step 4: Write only the new employees to the Roster sheet.
  writeNewEmployeesToRoster(deduplicationResult.employeesToAdd);

  // Step 5: Recalculate seniority ranks for all Roster rows now that new employees
  // have been added. New employees were written with a placeholder rank of 0.
  if (deduplicationResult.employeesToAdd.length > 0) {
    refreshAllSeniorityRanks();

    // Step 6: Reapply validation dropdowns to cover the newly added rows.
    const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);
    applyRosterValidation(rosterSheet);
  }

  // Step 7: Write the result summary back to the Ingestion sheet so the manager
  // can see what happened without needing to check the Roster sheet directly.
  const syncResult = {
    employeesAdded: deduplicationResult.employeesToAdd.length,
    employeesSkipped: deduplicationResult.skippedEmployees.length,
    statusMessage: "Sync complete on " + new Date().toLocaleString(),
  };

  writeSyncResultToIngestionSheet(syncResult);

  return syncResult;
}


/**
 * Removes vacation dates from the Roster that fall within the generated week.
 *
 * Called after a week is successfully written to the schedule sheet. Vacation dates
 * that land on a day within [weekStartDate, weekStartDate + 6] are removed from
 * the employee's Vacation Dates cell, since the schedule has already applied them.
 * Dates outside that range are left untouched so they will be available when the
 * manager generates future weeks.
 *
 * Two rules protect against accidental data loss:
 *   1. Only dates we can positively parse are removed. A garbled or unrecognized date
 *      string is left in the cell rather than silently deleted.
 *   2. This function is only called from orchestrateSingleWeekGeneration(), not from
 *      resolveEntireWeek(). Vacation dates are never pruned on a checkbox re-calculation.
 *
 * This function calls parseVacationDateString() and parseVacationDateStrings() from
 * scheduleEngine.js. In GAS all files share the global scope, so cross-file calls
 * like this work without any import statement.
 *
 * @param {Date} weekStartDate — The Monday of the week that was just generated.
 */
function removeProcessedVacationDates(weekStartDate) {
  const rosterSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.ROSTER);

  if (!rosterSheet) {
    return;
  }

  const lastRow = rosterSheet.getLastRow();
  if (lastRow < ROSTER_DATA_START_ROW) {
    return;
  }

  const rowCount = lastRow - ROSTER_DATA_START_ROW + 1;

  // Read only the Vacation Dates column — no need to load the entire Roster.
  const vacationRange = rosterSheet.getRange(
    ROSTER_DATA_START_ROW, ROSTER_COLUMN.VACATION_DATES, rowCount, 1
  );
  const cellValues = vacationRange.getValues();

  // The week runs Monday through Sunday. Build the Sunday boundary for the range check.
  const weekEndDate = new Date(weekStartDate);
  weekEndDate.setDate(weekStartDate.getDate() + 6);
  weekEndDate.setHours(23, 59, 59, 999);

  let anyRowChanged = false;

  const updatedValues = cellValues.map(function (row) {
    const rawCellValue = row[0];

    // parseVacationDateStrings handles empty/null cells and returns [] — no special-casing needed.
    const dateStrings = parseVacationDateStrings(rawCellValue);

    if (dateStrings.length === 0) {
      return row; // Nothing to clean in this cell.
    }

    const remainingDates = dateStrings.filter(function (dateString) {
      // Pass weekStartDate so the parser can infer the correct year for MM/DD shorthand.
      // For example, "4/14" becomes April 14 of weekStartDate's year.
      const parsedDate = parseVacationDateString(dateString, weekStartDate);

      // Cannot parse → keep it. Do not silently delete entries the manager intentionally entered.
      if (!parsedDate) {
        return true;
      }

      // Keep dates that fall outside the generated week's Monday–Sunday range.
      return parsedDate < weekStartDate || parsedDate > weekEndDate;
    });

    if (remainingDates.length < dateStrings.length) {
      anyRowChanged = true;
      // Rejoin with ", " to match the format the manager would have entered.
      return [remainingDates.join(", ")];
    }

    return row;
  });

  // Batch-write only if at least one cell actually changed — avoids a needless
  // sheet write (and the associated quota cost) when no vacation dates were due.
  if (anyRowChanged) {
    vacationRange.setValues(updatedValues);
  }
}
