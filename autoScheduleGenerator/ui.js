/**
 * ui.js — Google Apps Script triggers, custom menu, and generation orchestrators.
 * VERSION: 1.2.0
 *
 * This file is the entry point for all user-initiated actions. It contains:
 *   - onOpen():  Creates the "Schedule Admin" menu when the spreadsheet opens.
 *   - onEdit():  Routes edit events (checkbox changes, ID entry) to the correct handler.
 *   - Menu handlers: Functions triggered by menu items.
 *   - Orchestrators: Thin functions that coordinate the engine and formatter.
 *   - setupAllSheets(): First-run initialization.
 *
 * DESIGN RULE:
 * Every function in this file is either a pure router (onEdit) or a thin orchestrator
 * that calls one or more worker functions from other files. No scheduling logic, no
 * sheet reads beyond what is needed for routing, no formatting logic lives here.
 * If a bug occurs in generation or formatting, the trace will not lead to this file.
 *
 * GAS TRIGGER NOTES:
 * - onOpen() and onEdit() are simple triggers — they run automatically with limited
 *   permissions and cannot access external services.
 * - Menu handler functions (menuSync*, menuGenerate*, menuSetup*) are called by the
 *   user from the menu and run with the user's full permissions.
 * - The onEdit() function must complete quickly to avoid blocking the user's typing.
 *   Any slow operation (like schedule regeneration) is triggered asynchronously by
 *   showing a toast first, then running the operation.
 */


// ---------------------------------------------------------------------------
// GAS Simple Triggers
// ---------------------------------------------------------------------------

/**
 * Creates the "Schedule Admin" custom menu in the Google Sheets menu bar.
 *
 * This function runs automatically every time the spreadsheet is opened.
 * It uses the simple trigger signature onOpen() which GAS recognizes automatically
 * — no manual trigger installation is required.
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("Schedule Admin", [
      { name: "Sync Roster", functionName: "menuSyncRoster" },
      { name: "Refresh Seniority", functionName: "menuRefreshSeniority" },
      { name: "Sync Shift Dropdowns", functionName: "menuSyncShiftDropdowns" },
      { name: "Load Departments", functionName: "menuLoadDepartments" },
      null, // Separator line
      { name: "GENERATE (1 week)", functionName: "menuGenerate1Week" },
      { name: "GENERATE (3 weeks)", functionName: "menuGenerate3Weeks" },
      null, // Separator line
      { name: "Setup Department Settings Tab", functionName: "menuSetupDeptSettings" },
      { name: "Setup Sheets (First Run Only)", functionName: "menuSetupAllSheets" },
    ]);
}


/**
 * Routes edit events from the user to the appropriate handler.
 *
 * This function is a pure router — it inspects the edit event and dispatches to
 * another function. All logic for what to do lives in the dispatched function.
 *
 * Routed events:
 *   1. Ingestion sheet, cell B3 (Source Spreadsheet ID changed):
 *      → populateDepartmentDropdown() — refreshes the department dropdown
 *
 *   2. Roster sheet, column D (Status changed for an employee):
 *      → recalculateSeniorityRankForRow() — updates that employee's seniority rank
 *
 *   3. Week_* sheet, VAC or RDO row, columns C–I:
 *      → resolveEntireWeek() — re-runs Phases 1–3 and re-formats the sheet
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} editEvent — The edit event object from GAS.
 */
function onEdit(editEvent) {
  const editedSheet = editEvent.source.getActiveSheet();
  const editedRange = editEvent.range;
  const sheetName = editedSheet.getName();

  // --- Route 1: Ingestion sheet — Source Spreadsheet ID changed ---
  if (sheetName === SHEET_NAMES.INGESTION) {
    const editedCellAddress = editedRange.getA1Notation();
    if (editedCellAddress === INGESTION_CELL.SOURCE_SPREADSHEET_ID) {
      const newSpreadsheetId = editedRange.getValue().toString().trim();
      if (newSpreadsheetId !== "") {
        // Cannot call populateDepartmentDropdown here - onEdit is a simple trigger
        // and does not have the Oauth scope needed to open external spreadsheets.
        // Prompt the manager to use the menu action instead.
        const ingestionSheet = editEvent.source.getSheetByName(SHEET_NAMES.INGESTION);
        if (ingestionSheet) {
          ingestionSheet.getRange(INGESTION_CELL.DEPARTMENT).setValue("← click 'Load Departments from Source' in Schedule Admin Menu");
        }
      }
    }
    return;
  }

  // --- Route 2: Roster sheet — Status column changed ---
  if (sheetName === SHEET_NAMES.ROSTER) {
    if (editedRange.getColumn() === ROSTER_COLUMN.STATUS &&
      editedRange.getRow() >= ROSTER_DATA_START_ROW) {
      recalculateSeniorityRankForRow(editedRange.getRow());
    }
    return;
  }

  // --- Route 3: Week schedule sheet — VAC or RDO checkbox changed ---
  // Week sheets are named "Week_MM_DD_YY" — check for the "Week_" prefix.
  if (sheetName.startsWith("Week_")) {
    const editedRow = editedRange.getRow();
    const editedColumn = editedRange.getColumn();

    // Only react to edits in the day columns (Monday through Sunday).
    if (editedColumn < WEEK_SHEET.COL_MONDAY || editedColumn > WEEK_SHEET.COL_SUNDAY) {
      return;
    }

    // Only react to edits in the employee data rows (not header or summary rows).
    if (editedRow < WEEK_SHEET.DATA_START_ROW) {
      return;
    }

    // Determine whether the edited row is a VAC or RDO row by checking its offset
    // within the three-row employee block.
    // Block structure: (row - DATA_START_ROW) % 3 === 0 → VAC, === 1 → RDO, === 2 → SHIFT
    const rowOffsetWithinBlock = (editedRow - WEEK_SHEET.DATA_START_ROW) % WEEK_SHEET.ROWS_PER_EMPLOYEE;
    const isVacationRow = rowOffsetWithinBlock === WEEK_SHEET.ROW_OFFSET_VAC;
    const isRequestedDayOffRow = rowOffsetWithinBlock === WEEK_SHEET.ROW_OFFSET_RDO;

    if (isVacationRow || isRequestedDayOffRow) {
      // Show a brief toast to acknowledge the edit before the recalculation begins.
      // GAS toasts appear immediately and do not block execution.
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Re-calculating schedule...", "Schedule Admin", 3
      );
      resolveEntireWeek(editedSheet);
    }
  }
}


// ---------------------------------------------------------------------------
// Menu Handlers
// ---------------------------------------------------------------------------

/**
 * Triggered by: Schedule Admin → Sync Roster
 *
 * Reads the source spreadsheet ID and department from the Ingestion sheet,
 * fetches matching employees, deduplicates against the Roster, and writes new rows.
 */
function menuSyncRoster() {
  try {
    const syncResult = syncRosterFromSource();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      syncResult.employeesAdded + " employee(s) added. " +
      syncResult.employeesSkipped + " skipped (already on roster).",
      "Sync Complete",
      6
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert("Sync Failed", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Triggered by: Schedule Admin → Refresh Seniority
 *
 * Recalculates seniority ranks for every employee on the Roster sheet.
 * Useful after manually editing hire dates or employment status in bulk.
 */
function menuRefreshSeniority() {
  try {
    refreshAllSeniorityRanks();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Seniority ranks updated for all employees.", "Refresh Complete", 4
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert("Refresh Failed", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Triggered by: Schedule Admin → Sync Shift Dropdowns
 *
 * Re-reads the Settings sheet and updates the Preferred Shift dropdown on the Roster
 * to reflect any shifts that have been added, removed, or renamed.
 */
function menuSyncShiftDropdowns() {
  try {
    const rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.ROSTER);
    if (!rosterSheet) {
      throw new Error("Roster sheet not found.");
    }
    applyRosterValidation(rosterSheet);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Shift dropdowns updated from Settings sheet.", "Sync Complete", 4
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert("Sync Failed", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Triggered by: Schedule Admin → Load Departments from Source
 *
 * Reads the source spreadsheet ID from Ingestion B3 and populates the department
 * dropdown in B4. Must be a menu-triggered function — onEdit() cannot open external
 * spreadsheets because simple triggers lack the required OAuth authorization scope.
 */
function menuLoadDepartments() {
  const ingestionSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.INGESTION);

  if (!ingestionSheet) {
    SpreadsheetApp.getUi().alert("Load Failed", "Ingestion sheet not found. Run Setup Sheets First.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const sourceSpreadsheetId = ingestionSheet
    .getRange(INGESTION_CELL.SOURCE_SPREADSHEET_ID)
    .getValue()
    .toString()
    .trim();

  if (!sourceSpreadsheetId) {
    SpreadsheetApp.getUi().alert("Load Failed", "Enter a source spreadsheet ID in cell B3 first.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  try {
    populateDepartmentDropdown(sourceSpreadsheetId);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Department dropdown updated from source spreadsheet.", "Load Complete", 4
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert("Load Failed", error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}




/**
 * Triggered by: Schedule Admin → Setup Sheets (First Run Only)
 *
 * Creates the Ingestion, Roster, and Settings sheets with default layouts,
 * headers, and data validation. Safe to run more than once — existing content
 * is preserved and only missing sheets are created.
 */
function menuSetupAllSheets() {
  setupAllSheets();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Setup complete. Fill in the Settings sheet and sync your roster to get started.",
    "Setup Complete",
    8
  );
}


/**
 * Triggered by: Schedule Admin → GENERATE (1 week)
 */
function menuGenerate1Week() {
  menuGenerateAllDeptsForWeekCount_(1);
}


/**
 * Triggered by: Schedule Admin → GENERATE (3 weeks)
 */
function menuGenerate3Weeks() {
  menuGenerateAllDeptsForWeekCount_(3);
}


/**
 * Shared implementation for both generation menu items.
 *
 * Guards that the Departments tab exists before opening the date picker.
 * If the tab is missing, prompts the manager to run Setup Sheets first —
 * Setup Sheets now auto-creates the Departments tab with a template.
 *
 * @param {number} weekCount — 1 or 3.
 */
function menuGenerateAllDeptsForWeekCount_(weekCount) {
  const departmentsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.DEPARTMENTS);

  if (!departmentsSheet) {
    SpreadsheetApp.getUi().alert(
      "Departments Tab Missing",
      "No \"" + SHEET_NAMES.DEPARTMENTS + "\" tab was found.\n\n" +
      "Run Schedule Admin \u2192 Setup Sheets (First Run Only) to create it, " +
      "then add your departments and use Setup Department Settings Tab to create per-department Settings tabs.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  showAllDeptsDatePickerDialog(weekCount);
}


/**
 * Triggered by: Schedule Admin → Setup Department Settings Tab
 *
 * Prompts the manager for a department name and creates a Settings_{dept} tab
 * pre-populated with the standard shift template. Safe to run multiple times —
 * if the tab already exists, the manager is notified rather than overwriting it.
 */
function menuSetupDeptSettings() {
  const userInterface = SpreadsheetApp.getUi();
  const response = userInterface.prompt(
    "Setup Department Settings Tab",
    "Enter the department name (e.g., Morning, Drivers/Merch, Full Time Cashier):",
    userInterface.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== userInterface.Button.OK) {
    return;
  }

  const departmentName = response.getResponseText().trim();
  if (!departmentName) {
    userInterface.alert("Setup Cancelled", "Department name cannot be blank.", userInterface.ButtonSet.OK);
    return;
  }

  const tabName = SETTINGS_SHEET_PREFIX + departmentName;
  const workbook = SpreadsheetApp.getActiveSpreadsheet();

  if (workbook.getSheetByName(tabName)) {
    userInterface.alert(
      "Tab Already Exists",
      "\"" + tabName + "\" already exists. No changes were made.\n\n" +
      "Edit that tab directly to update shift definitions or staffing requirements.",
      userInterface.ButtonSet.OK
    );
    return;
  }

  const newSettingsSheet = workbook.insertSheet(tabName);
  setupSettingsSheetTemplate(newSettingsSheet);

  workbook.toast(
    "\"" + tabName + "\" created with a default shift template.\n" +
    "Edit the shift definitions and staffing requirements to match this department.\n\n" +
    "Then add \"" + departmentName + "\" to the Departments tab (col A) with \"" + tabName + "\" in col B and TRUE in col C.",
    "Setup Complete",
    10
  );
}


// ---------------------------------------------------------------------------
// Generation Orchestrators
// ---------------------------------------------------------------------------

/**
 * Shows a date picker dialog for schedule generation.
 *
 * @param {number} weekCount — 1 for single week, 3 for three consecutive weeks.
 */
function showAllDeptsDatePickerDialog(weekCount) {
  const currentMonday = getMondayOfCurrentWeek();
  const defaultDateValue = Utilities.formatDate(
    currentMonday, Session.getScriptTimeZone(), "yyyy-MM-dd"
  );

  const htmlOutput = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: "Google Sans", Arial, sans-serif; font-size: 13px; color: #202124; margin: 0; padding: 20px 24px 16px; }
          label { display: block; font-weight: 500; margin-bottom: 8px; }
          input[type="date"] { width: 100%; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; box-sizing: border-box; outline: none; transition: border-color 0.15s; }
          input[type="date"]:focus { border-color: #1a73e8; }
          .hint { font-size: 11px; color: #5f6368; margin-top: 5px; margin-bottom: 16px; }
          .button-row { display: flex; justify-content: flex-end; gap: 8px; margin-top: 4px; }
          button { padding: 8px 18px; font-size: 13px; font-family: inherit; border-radius: 4px; cursor: pointer; border: 1px solid transparent; }
          #cancelBtn { background: #fff; color: #1a73e8; border-color: #dadce0; }
          #cancelBtn:hover { background: #f1f3f4; }
          #okBtn { background: #1a73e8; color: #fff; }
          #okBtn:hover { background: #1765cc; }
          #okBtn:disabled { background: #dadce0; color: #80868b; cursor: default; }
          #statusMsg { font-size: 12px; color: #d93025; min-height: 16px; margin-top: 10px; }
          #statusMsg.info { color: #5f6368; }
        </style>
      </head>
      <body>
        <label for="dateInput">Starting Monday:</label>
        <input type="date" id="dateInput" value="${defaultDateValue}">
        <p class="hint">Generates one schedule tab per active department${weekCount === 3 ? " for 3 consecutive weeks" : ""}. If you pick a non-Monday, the schedule will snap to that week's Monday.</p>
        <div id="statusMsg"></div>
        <div class="button-row">
          <button type="button" id="cancelBtn" onclick="google.script.host.close()">Cancel</button>
          <button type="button" id="okBtn" onclick="submitDate()">Generate</button>
        </div>
        <script>
          function submitDate() {
            const dateValue = document.getElementById('dateInput').value;
            if (!dateValue) { showStatus('Please select a date before generating.', false); return; }
            document.getElementById('okBtn').disabled = true;
            document.getElementById('okBtn').textContent = 'Generating\u2026';
            showStatus('Starting \u2014 this may take a moment.', true);
            google.script.run
              .withSuccessHandler(function() { google.script.host.close(); })
              .withFailureHandler(function(error) {
                document.getElementById('okBtn').disabled = false;
                document.getElementById('okBtn').textContent = 'Generate';
                showStatus('Error: ' + error.message, false);
              })
              .receiveAllDeptsDateAndGenerate(dateValue, ${weekCount});
          }
          function showStatus(message, isInfo) {
            const element = document.getElementById('statusMsg');
            element.textContent = message;
            element.className = isInfo ? 'info' : '';
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(360)
  .setHeight(240);

  const dialogTitle = weekCount === 3 ? "Generate (3 Weeks)" : "Generate (1 Week)";
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, dialogTitle);
}


/**
 * Called by the date picker dialog via google.script.run.
 *
 * Parses and snaps the date to Monday, then generates all active departments for
 * each requested week. Produces weekCount × deptCount schedule sheets in one run.
 *
 * @param {string} dateString — A date string in "YYYY-MM-DD" format from the date picker.
 * @param {number} weekCount  — 1 for a single week, 3 for three consecutive weeks.
 */
function receiveAllDeptsDateAndGenerate(dateString, weekCount) {
  const parts = dateString.split("-");
  const year  = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day   = parseInt(parts[2], 10);

  const firstMonday = new Date(year, month, day);
  firstMonday.setHours(0, 0, 0, 0);

  // Snap to Monday.
  const dayOfWeek = firstMonday.getDay();
  if (dayOfWeek !== 1) {
    const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    firstMonday.setDate(firstMonday.getDate() - daysToSubtract);
  }

  // Run preflight once before any week is generated.
  // Blocking errors throw and are caught by the dialog's withFailureHandler.
  // Warnings are shown in a confirm dialog — manager can proceed or cancel.
  const preflightWarnings = runPreGenerationPreflight(firstMonday);
  if (preflightWarnings.length > 0) {
    const warningText =
      "The following issues were found. Generation will still run, but these employees " +
      "may be scheduled incorrectly or skipped entirely.\n\n" +
      preflightWarnings.map(function(w, i) { return (i + 1) + ". " + w; }).join("\n\n") +
      "\n\nClick OK to proceed anyway, or Cancel to fix these issues first.";
    const response = SpreadsheetApp.getUi().alert(
      "Pre-flight Warnings (" + preflightWarnings.length + ")",
      warningText,
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response !== SpreadsheetApp.getUi().Button.OK) {
      return; // Manager chose to fix issues before generating.
    }
  }

  let totalSheetsGenerated = 0;

  for (let weekOffset = 0; weekOffset < weekCount; weekOffset++) {
    const weekStartDate = new Date(firstMonday);
    weekStartDate.setDate(firstMonday.getDate() + (weekOffset * 7));

    const allDeptResults = generateAllDepartmentSchedules(weekStartDate);

    if (!allDeptResults || allDeptResults.size === 0) {
      throw new Error(
        "No departments were generated for week starting " + weekStartDate.toDateString() + ". " +
        "Check that the Departments tab has at least one active row (column C = TRUE) " +
        "and that each department's employees have a matching Department value on the Roster."
      );
    }

    writeAllDepartmentSchedules_(allDeptResults, weekStartDate);
    totalSheetsGenerated += allDeptResults.size;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    totalSheetsGenerated + " schedule sheet" + (totalSheetsGenerated === 1 ? "" : "s") + " generated" +
    (weekCount > 1 ? " across " + weekCount + " weeks" : "") + ".",
    "Generation Complete",
    8
  );
}


/**
 * Re-runs the schedule generation phases (1–3) for an existing Week sheet and re-formats it.
 *
 * Called by onEdit() when a manager checks or unchecks a VAC or RDO checkbox.
 * Reads the current checkbox state from the sheet (which represents the manager's
 * decisions), rebuilds the grid with those decisions locked in, re-runs Phase 1–3,
 * then re-writes the SHIFT rows and re-applies formatting.
 *
 * This function does NOT re-run Phase 0 (it does not reload the roster from scratch)
 * because the manager may have manually modified checkboxes and we want to respect
 * the current state of the sheet as the source of truth for VAC/RDO decisions.
 *
 * @param {Sheet} weekSheet — The Week_MM_DD_YY sheet being edited.
 */
function resolveEntireWeek(weekSheet) {
  try {
    // Load the roster in seniority order — needed to map employee indices to grid rows.
    const employeeList = loadRosterSortedBySeniority();

    if (employeeList.length === 0) {
      return;
    }

    // Read the manager's current VAC/RDO checkbox decisions from the sheet.
    const checkboxGrid = readCheckboxStateFromSheet(weekSheet, employeeList.length);

    // Parse the week start date from the sheet name (format: "Week_MM_DD_YY").
    const weekStartDate = parseWeekStartDateFromSheetName(weekSheet.getName());

    if (!weekStartDate) {
      Logger.log(
        "WARNING: resolveEntireWeek could not parse the week start date from sheet name: " +
        weekSheet.getName()
      );
      return;
    }

    // Determine which department this sheet belongs to by parsing the dept name from
    // the sheet name suffix (format: "Week_MM_DD_YY_DeptName"). Then look up that
    // department's Settings tab so the correct shifts and staffing requirements are used.
    // Falls back to the base Settings tab if the sheet name has no dept suffix or if
    // the department is not found in the Departments tab.
    const sheetName = weekSheet.getName();
    const deptNameFromSheet = parseDeptNameFromWeekSheetName_(sheetName);
    const deptSettings = loadSettingsForResolve_(deptNameFromSheet);

    const shiftTimingMap      = deptSettings.shiftTimingMap;
    const staffingRequirements = deptSettings.staffingRequirements;

    // Initialize a fresh grid that reflects the manager's checkbox decisions.
    // The checkbox grid already has VAC cells locked (locked = true), so they
    // will be respected by Phases 1–3 the same as in a fresh generation.
    const freshGrid = checkboxGrid;

    // Re-run Phases 1–3. Phase 0 is skipped because the checkbox state from the
    // sheet is already the starting point.
    runPhaseOnePreferenceAssignment(
      freshGrid, employeeList, shiftTimingMap, staffingRequirements, weekStartDate
    );
    runPhaseTwoHourEnforcement(freshGrid, employeeList, shiftTimingMap);
    runPhaseThreeGapResolution(freshGrid, employeeList, shiftTimingMap, staffingRequirements);

    // Department name for the header comes from the sheet name suffix (already parsed above).
    const departmentName = deptNameFromSheet || getDepartmentNameForHeader();

    writeAndFormatSchedule(
      weekSheet,
      employeeList,
      freshGrid,
      staffingRequirements,
      weekStartDate,
      departmentName
    );

  } catch (error) {
    Logger.log("resolveEntireWeek error: " + error.message + "\n" + error.stack);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Re-calculation failed: " + error.message,
      "Error",
      8
    );
  }
}


// ---------------------------------------------------------------------------
// First-Run Setup
// ---------------------------------------------------------------------------

/**
 * Creates and configures all sheets required by the tool.
 *
 * This function is idempotent — it is safe to run multiple times. Existing sheets
 * are not deleted or overwritten; only missing sheets are created. Validation
 * dropdowns are re-applied on every run to pick up any Settings sheet changes.
 */
function setupAllSheets() {
  const workbook = SpreadsheetApp.getActiveSpreadsheet();

  // --- Ingestion sheet ---
  setupIngestionSheet();

  // --- Roster sheet ---
  let rosterSheet = workbook.getSheetByName(SHEET_NAMES.ROSTER);
  if (!rosterSheet) {
    rosterSheet = workbook.insertSheet(SHEET_NAMES.ROSTER);
  }
  setupRosterSheetHeaders(rosterSheet);

  // --- Departments tab ---
  // Created automatically so managers never have to build it manually.
  // If it already exists, leave it untouched (manager may have customised it).
  if (!workbook.getSheetByName(SHEET_NAMES.DEPARTMENTS)) {
    setupDepartmentsSheetTemplate_(workbook.insertSheet(SHEET_NAMES.DEPARTMENTS));
  }
}


/**
 * Writes the header row to the Roster sheet and applies column formatting.
 *
 * @param {Sheet} rosterSheet — The Roster sheet object.
 */
function setupRosterSheetHeaders(rosterSheet) {
  // If the "Name" header already exists, the Roster has been set up before.
  // Skip to avoid resetting column widths the manager may have adjusted and to
  // prevent the header row from being re-written over any data that has crept into row 1.
  const existingFirstHeader = rosterSheet
    .getRange(1, ROSTER_COLUMN.NAME)
    .getValue()
    .toString()
    .trim();
  if (existingFirstHeader === "Name") {
    return;
  }

  const headerValues = [[
    "Name",
    "Employee ID",
    "Hire Date",
    "Status (FT/PT)",
    "Day Off Pref 1",
    "Day Off Pref 2",
    "Preferred Shift",
    "Qualified Shifts",
    "Vacation Dates",
    "Seniority Rank",
    "Department",
    "Qualified Departments",
    "Primary Role",
  ]];

  const headerRange = rosterSheet.getRange(1, 1, 1, headerValues[0].length);
  headerRange.setValues(headerValues);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(COLORS.HEADER_BG);
  headerRange.setFontColor(COLORS.HEADER_TEXT);

  // Freeze the header row so it stays visible when scrolling.
  rosterSheet.setFrozenRows(1);

  // Column widths for readability.
  rosterSheet.setColumnWidth(ROSTER_COLUMN.NAME, 160);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.EMPLOYEE_ID, 110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.HIRE_DATE, 100);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.STATUS, 90);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.DAY_OFF_PREF_ONE, 110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.DAY_OFF_PREF_TWO, 110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.PREFERRED_SHIFT, 120);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.QUALIFIED_SHIFTS, 200);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.VACATION_DATES, 200);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.SENIORITY_RANK, 120);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.DEPARTMENT, 140);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.QUALIFIED_DEPARTMENTS, 200);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.PRIMARY_ROLE, 120);

  // Add a note to the Seniority Rank column header explaining it is script-managed.
  rosterSheet.getRange(1, ROSTER_COLUMN.SENIORITY_RANK).setNote(
    "This column is calculated and managed by the script.\n" +
    "Do not edit values here manually — they will be overwritten on the next Refresh Seniority run.\n\n" +
    "Higher number = more senior. Full-time employees always outrank part-time employees hired on the same date."
  );

  rosterSheet.getRange(1, ROSTER_COLUMN.DEPARTMENT).setNote(
    "The department this employee belongs to.\n" +
    "Must match a department name in column A of the Departments tab.\n" +
    "Capitalization and spaces are normalized automatically, so 'Full Time Cashier',\n" +
    "'full time cashier', and 'full_time_cashier' all resolve to the same department.\n" +
    "Example: Morning, Drivers/Merch, Full Time Cashier"
  );

  rosterSheet.getRange(1, ROSTER_COLUMN.QUALIFIED_DEPARTMENTS).setNote(
    "Comma-separated list of other departments this employee can float to when needed.\n" +
    "Leave blank if the employee works in their primary department only.\n" +
    "Example: Front End Assistants, Floater/Con"
  );

  rosterSheet.getRange(1, ROSTER_COLUMN.PRIMARY_ROLE).setNote(
    "The role this employee is assigned on days they work.\n" +
    "Shown in the ROLE row of the generated schedule.\n" +
    "Example: Cashier, SCO, PreScan, Carts, Assist, Go Backs, Liquor"
  );
}


/**
 * Writes the Settings sheet template with example shift definitions and staffing requirements.
 *
 * This gives managers a working starting point they can customize for their department.
 * The template includes one FT and one PT variant for three shift types (Morning, Mid, Closing).
 *
 * @param {Sheet} settingsSheet — The Settings sheet object.
 */
/**
 * Writes the Departments tab template with headers and an example row.
 *
 * Column layout:
 *   A — Department Name (must match employee Roster column K)
 *   B — Settings Tab Name (e.g., "Settings_Morning")
 *   C — Active (TRUE/FALSE)
 *   D — Header Accent Color (hex, optional)
 *
 * @param {Sheet} deptsSheet — The freshly-created Departments sheet.
 */
function setupDepartmentsSheetTemplate_(deptsSheet) {
  const headers = [["Department Name", "Settings Tab Name", "Active", "Accent Color (optional)"]];
  const exampleRow = [["My Department", "Settings_MyDepartment", "TRUE", ""]];

  const headerRange = deptsSheet.getRange("A1:D1");
  headerRange.setValues(headers);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(COLORS.HEADER_BG);
  headerRange.setFontColor(COLORS.HEADER_TEXT);

  deptsSheet.getRange("A2:D2").setValues(exampleRow);
  deptsSheet.getRange("A2").setFontStyle("italic").setFontColor("#999999");
  deptsSheet.getRange("B2").setFontStyle("italic").setFontColor("#999999");

  deptsSheet.setColumnWidth(1, 180);
  deptsSheet.setColumnWidth(2, 200);
  deptsSheet.setColumnWidth(3, 80);
  deptsSheet.setColumnWidth(4, 180);

  deptsSheet.getRange("A1").setNote(
    "Department name — must match the value in the Department column (col K) of the Roster sheet exactly."
  );
  deptsSheet.getRange("B1").setNote(
    "The name of the Settings tab for this department.\n" +
    "Use Schedule Admin → Setup Department Settings Tab to create it."
  );
  deptsSheet.getRange("C1").setNote(
    "Set to TRUE to include this department in generation, FALSE to skip it."
  );

  deptsSheet.setFrozenRows(1);
}


function setupSettingsSheetTemplate(settingsSheet) {
  // Only write the template if the sheet appears to be empty.
  if (settingsSheet.getLastRow() > 0) {
    return;
  }

  // --- Table 1: Staffing requirements (columns A–C) ---
  // Column C controls whether the target is a head count ("Count") or total paid hours ("Hours").
  // Leave column C blank or write "Count" for count mode (the default).
  const staffingHeaders = [["Day", "Target Value", "Type (Count / Hours)"]];
  const staffingData = [
    ["Monday",    6, "Count"],
    ["Tuesday",   6, "Count"],
    ["Wednesday", 6, "Count"],
    ["Thursday",  6, "Count"],
    ["Friday",    6, "Count"],
    ["Saturday",  4, "Count"],
    ["Sunday",    4, "Count"],
  ];

  settingsSheet.getRange("A1:C1").setValues(staffingHeaders).setFontWeight("bold");
  settingsSheet.getRange("A2:C8").setValues(staffingData);

  // Add a dropdown on the Type column so managers can easily switch modes.
  const modeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Count", "Hours"], true)
    .setAllowInvalid(false)
    .setHelpText("Count = minimum employees scheduled. Hours = minimum total paid hours across all employees.")
    .build();
  settingsSheet.getRange("C2:C8").setDataValidation(modeValidation);

  // --- Table 2: Shift definitions (columns D–I) ---
  // Shift design rationale:
  //   "Early"   covers 4:00 AM open through the morning rush — needed for stocking and setup.
  //   "Morning" is the standard day shift, overlapping Early for a mid-morning coverage boost.
  //   "Mid"     bridges the gap between morning and evening, covering the busiest midday window.
  //   "Closing" runs through the evening into close, ending before or at the day's coverage cutoff.
  //
  // FT shifts: 8 paid hours + 30-minute unpaid lunch = 8.5-hour clock block.
  // PT shifts: 5 paid hours, available with or without a 30-minute unpaid lunch ("+"-suffix variant).
  // The coverage window for Saturday ends at 10:00 PM and Sunday at 9:00 PM, so Closing|FT
  // on those days ends at 10:00 PM and 9:00 PM respectively — add those as weekend-specific
  // rows if the department runs different closing times on weekends.
  const shiftHeaders = [["Shift Name", "Status (FT/PT)", "Start Time", "End Time", "Paid Hours", "Has Lunch (TRUE/FALSE)"]];
  const shiftData = [
    // --- Full-Time Shifts (8 paid hrs + 30 min unpaid lunch = 8.5 hr block) ---
    ["Early", "FT", "4:00 AM", "12:30 PM", 8.0, true],   // 4:00 AM – 12:30 PM  open shift; covers early-morning store setup
    ["Morning", "FT", "6:00 AM", "2:30 PM", 8.0, true],   // 6:00 AM –  2:30 PM  standard day shift
    ["Mid", "FT", "10:00 AM", "6:30 PM", 8.0, true],   // 10:00 AM –  6:30 PM bridges morning and evening coverage
    ["Closing", "FT", "1:30 PM", "10:00 PM", 8.0, true],   // 1:30 PM – 10:00 PM  evening through close (aligns with Sat 10 PM cutoff)

    // --- Part-Time Shifts, No Lunch (5 paid hrs = 5 hr block) ---
    ["Early", "PT", "4:00 AM", "9:00 AM", 5.0, false],  // 4:00 AM –  9:00 AM  early coverage, no lunch needed at this length
    ["Morning", "PT", "6:00 AM", "11:00 AM", 5.0, false],  // 6:00 AM – 11:00 AM
    ["Mid", "PT", "10:00 AM", "3:00 PM", 5.0, false],  // 10:00 AM –  3:00 PM
    ["Closing", "PT", "5:00 PM", "10:00 PM", 5.0, false],  // 5:00 PM – 10:00 PM  aligns with Saturday close

    // --- Part-Time Shifts, With Lunch (5 paid hrs + 30 min unpaid = 5.5 hr block) ---
    // The "+" suffix distinguishes the lunch variant. Both "Morning" and "Morning+" are
    // valid entries in an employee's Qualified Shifts column.
    ["Early+", "PT", "4:00 AM", "9:30 AM", 5.0, true],   // 4:00 AM –  9:30 AM
    ["Morning+", "PT", "6:00 AM", "11:30 AM", 5.0, true],   // 6:00 AM – 11:30 AM
    ["Mid+", "PT", "10:00 AM", "3:30 PM", 5.0, true],   // 10:00 AM –  3:30 PM
    ["Closing+", "PT", "4:30 PM", "10:00 PM", 5.0, true],   // 4:30 PM – 10:00 PM
  ];

  settingsSheet.getRange("D1:I1").setValues(shiftHeaders).setFontWeight("bold");
  settingsSheet.getRange("D2:I13").setValues(shiftData);

  // Style both header rows.
  settingsSheet.getRange("A1:C1").setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT);
  settingsSheet.getRange("D1:I1").setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT);

  // Format the Start Time and End Time columns as time values so GAS reads them correctly.
  settingsSheet.getRange("F2:G50").setNumberFormat("h:mm AM/PM");
}


// ---------------------------------------------------------------------------
// Utility Functions
// ---------------------------------------------------------------------------

/**
 * Returns the Monday of the current calendar week as a Date object.
 *
 * GAS uses the local system timezone for Date objects. The Sunday/Monday boundary
 * is handled by checking the day-of-week and subtracting the appropriate number of days.
 *
 * @returns {Date} The Monday of the current week (time set to 00:00:00).
 */
function getMondayOfCurrentWeek() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // Calculate how many days to subtract to reach Monday.
  // If today is Sunday (0), go back 6 days. If Monday (1), go back 0. Etc.
  const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

  const monday = new Date(today);
  monday.setDate(today.getDate() - daysToSubtract);
  monday.setHours(0, 0, 0, 0);

  return monday;
}


/**
 * Reads the department name from the Ingestion sheet for use in schedule headers.
 *
 * Returns "Unknown Department" if the Ingestion sheet is not set up or the department
 * cell is blank, so that the schedule header always shows something useful.
 *
 * @returns {string} The department name.
 */
function getDepartmentNameForHeader() {
  const ingestionSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.INGESTION);

  if (!ingestionSheet) {
    return "Unknown Department";
  }

  const departmentValue = ingestionSheet
    .getRange(INGESTION_CELL.DEPARTMENT)
    .getValue()
    .toString()
    .trim();

  if (!departmentValue ||
    departmentValue === "" ||
    departmentValue === "— enter spreadsheet ID first —" ||
    departmentValue.startsWith("←") ||
    departmentValue.startsWith("Error:")) {
    return "Unknown Department";
  }

  return departmentValue;
}


/**
 * Parses the week start date (the Monday) from a Week sheet's name.
 *
 * Sheet names follow the format "Week_MM_DD_YY" (e.g., "Week_04_07_26" for April 7, 2026).
 * This function reverses the name generation logic from generateWeekSheetName().
 *
 * @param {string} sheetName — The name of a Week schedule sheet.
 * @returns {Date|null} The Monday date, or null if the name does not match the expected format.
 */
function parseWeekStartDateFromSheetName(sheetName) {
  // Match "Week_MM_DD_YY" with an optional "_DeptName" suffix (multi-dept format).
  const namePattern = /^Week_(\d{2})_(\d{2})_(\d{2})(?:_.*)?$/;
  const match = sheetName.match(namePattern);

  if (!match) {
    return null;
  }

  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);
  const shortYear = parseInt(match[3], 10);

  // Convert two-digit year to four-digit year.
  // Years 00–99 map to 2000–2099. Adjust this logic if the tool is still in use in 2100.
  const fullYear = 2000 + shortYear;

  const parsedDate = new Date(fullYear, month - 1, day);
  parsedDate.setHours(0, 0, 0, 0);

  return isNaN(parsedDate.getTime()) ? null : parsedDate;
}


/**
 * Extracts the department name from a week sheet name that includes a dept suffix.
 *
 * "Week_04_07_26_Morning"          → "Morning"
 * "Week_04_07_26_Drivers/Morning Merch" → "Drivers/Morning Merch"
 * "Week_04_07_26"                  → null (no dept suffix)
 *
 * @param {string} sheetName
 * @returns {string|null}
 */
function parseDeptNameFromWeekSheetName_(sheetName) {
  // Sheet name format: Week_MM_DD_YY_DeptName  (dept name follows the 4th underscore)
  const match = sheetName.match(/^Week_\d{2}_\d{2}_\d{2}_(.+)$/);
  return match ? match[1] : null;
}


/**
 * Loads shift timing map and staffing requirements for a re-calculation triggered by
 * a VAC/RDO checkbox edit on an existing week sheet.
 *
 * Looks up the department by name in the Departments tab. If not found (or no dept name),
 * falls back to the base Settings tab so single-dept setups continue to work.
 *
 * @param {string|null} departmentName — Parsed from the sheet name, or null.
 * @returns {{ shiftTimingMap: Object, staffingRequirements: Object }}
 */
function loadSettingsForResolve_(departmentName) {
  if (departmentName) {
    // Normalize the name parsed from the sheet suffix so it matches the normalized
    // keys stored in the settings map by readDepartmentList_() / loadAllDepartmentSettings().
    const normalizedName  = normalizeDeptName_(departmentName);
    const allDeptSettings = loadAllDepartmentSettings();
    if (allDeptSettings && allDeptSettings.has(normalizedName)) {
      const settings = allDeptSettings.get(normalizedName);
      return {
        shiftTimingMap:       settings.shiftTimingMap,
        staffingRequirements: settings.staffingRequirements,
      };
    }
  }

  // Fallback: base Settings tab (single-dept mode or dept not found in map).
  return {
    shiftTimingMap:       buildShiftTimingMap(),
    staffingRequirements: loadStaffingRequirements(),
  };
}


