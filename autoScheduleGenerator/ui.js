/**
 * ui.js — Google Apps Script triggers, custom menu, and generation orchestrators.
 * VERSION 0.3.0
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
      { name: "Sync Roster",          functionName: "menuSyncRoster" },
      { name: "Refresh Seniority",    functionName: "menuRefreshSeniority" },
      { name: "Sync Shift Dropdowns", functionName: "menuSyncShiftDropdowns" },
      null, // Separator line
      { name: "GENERATE SCHEDULE DRAFT (3 weeks)", functionName: "menuGenerateScheduleDraft" },
      null, // Separator line
      { name: "Setup Sheets (First Run Only)",      functionName: "menuSetupAllSheets" },
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
  const sheetName   = editedSheet.getName();

  // --- Route 1: Ingestion sheet — Source Spreadsheet ID changed ---
  if (sheetName === SHEET_NAMES.INGESTION) {
    const editedCellAddress = editedRange.getA1Notation();
    if (editedCellAddress === INGESTION_CELL.SOURCE_SPREADSHEET_ID) {
      const newSpreadsheetId = editedRange.getValue().toString().trim();
      if (newSpreadsheetId !== "") {
        populateDepartmentDropdown(newSpreadsheetId);
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
    const editedRow    = editedRange.getRow();
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
    const isVacationRow        = rowOffsetWithinBlock === WEEK_SHEET.ROW_OFFSET_VAC;
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
 * Triggered by: Schedule Admin → GENERATE SCHEDULE DRAFT (3 weeks)
 *
 * Asks the manager to confirm the starting Monday, then generates three consecutive
 * weekly schedule sheets: the starting week, the following week, and the week after.
 *
 * This is the main entry point for schedule generation.
 */
function menuGenerateScheduleDraft() {
  const startingMonday = promptForStartingWeek();

  if (!startingMonday) {
    // The manager cancelled the dialog or provided an invalid date — do nothing.
    return;
  }

  try {
    orchestrateMultiWeekGeneration(startingMonday);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Generation Failed",
      error.message + "\n\nCheck the Execution Log (Extensions → Apps Script → Executions) for details.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
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


// ---------------------------------------------------------------------------
// Generation Orchestrators
// ---------------------------------------------------------------------------

/**
 * Shows a dialog asking the manager to confirm the Monday of the starting week.
 *
 * Defaults to the current Monday so the manager can simply click OK for the most
 * common case (generating starting this week). If the manager enters a different
 * date, it is parsed and validated.
 *
 * This function handles only the UI interaction — it does not call the engine.
 *
 * @returns {Date|null} The confirmed Monday date, or null if the manager cancelled
 *   or entered an invalid date.
 */
function promptForStartingWeek() {
  const currentMonday = getMondayOfCurrentWeek();
  const defaultDateString = currentMonday.toLocaleDateString("en-US", {
    month: "2-digit", day: "2-digit", year: "numeric"
  });

  const userInterface = SpreadsheetApp.getUi();

  const response = userInterface.prompt(
    "Generate Schedule Draft",
    "Enter the Monday of the starting week (MM/DD/YYYY).\n" +
    "Three weekly sheets will be generated: this week, next week, and the week after.\n\n" +
    "Starting Monday:",
    userInterface.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== userInterface.Button.OK) {
    // Manager clicked Cancel — abort without generating.
    return null;
  }

  const enteredText = response.getResponseText().trim();

  // If the manager left the field blank, use the current Monday as the default.
  if (enteredText === "" || enteredText === defaultDateString) {
    return currentMonday;
  }

  const parsedDate = new Date(enteredText);

  if (isNaN(parsedDate.getTime())) {
    userInterface.alert(
      "Invalid Date",
      "\"" + enteredText + "\" could not be interpreted as a date. " +
      "Please use MM/DD/YYYY format (e.g., 04/07/2026).",
      userInterface.ButtonSet.OK
    );
    return null;
  }

  // Validate that the entered date is a Monday. If not, find the Monday of that week.
  const dayOfWeek = parsedDate.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  if (dayOfWeek !== 1) {
    // Adjust to the Monday of the entered date's week.
    const daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    parsedDate.setDate(parsedDate.getDate() - daysToSubtract);

    userInterface.alert(
      "Date Adjusted",
      "The date you entered is not a Monday. The schedule will start on " +
      parsedDate.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" }) +
      " instead.",
      userInterface.ButtonSet.OK
    );
  }

  return parsedDate;
}


/**
 * Generates and formats three consecutive weekly schedule sheets starting from the given Monday.
 *
 * Each week is generated completely independently — the grid for week 2 is built fresh
 * from the Roster sheet, not derived from week 1's output. This means:
 *   - Vacation dates are checked against each week's date range independently.
 *   - RDO preferences are re-evaluated each week from scratch.
 *   - Hour rules are enforced per-week, not across weeks.
 *   - Managers can re-generate a single week without affecting the others.
 *
 * @param {Date} startingMondayDate — The Monday of the first week to generate.
 */
function orchestrateMultiWeekGeneration(startingMondayDate) {
  const generatedSheetNames = [];

  // Generate three consecutive weeks by incrementing the start date by 7 days each time.
  for (let weekOffset = 0; weekOffset < 3; weekOffset++) {
    const weekStartDate = new Date(startingMondayDate);
    weekStartDate.setDate(startingMondayDate.getDate() + (weekOffset * 7));

    const sheetName = orchestrateSingleWeekGeneration(weekStartDate);
    generatedSheetNames.push(sheetName);
  }

  // Show a completion toast listing all three generated sheet names.
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Three schedule drafts generated:\n" + generatedSheetNames.join("\n"),
    "Generation Complete",
    10
  );
}


/**
 * Generates and formats a single weekly schedule sheet for the given Monday.
 *
 * This function coordinates the engine and formatter:
 *   1. Calls generateWeeklySchedule() to run the 4-phase algorithm.
 *   2. Gets or creates the target sheet.
 *   3. Calls writeAndFormatSchedule() to write the grid to the sheet.
 *
 * It contains no scheduling or formatting logic — those live in scheduleEngine.js
 * and formatter.js respectively.
 *
 * @param {Date} weekStartDate — The Monday of the week to generate.
 * @returns {string} The name of the generated sheet (e.g., "Week_04_07_26").
 */
function orchestrateSingleWeekGeneration(weekStartDate) {
  // Read the department name from the Ingestion sheet to display in the schedule header.
  const departmentName = getDepartmentNameForHeader();

  // Run the 4-phase generation algorithm. This produces a pure JavaScript data
  // structure (the WeekGrid) without touching any sheet.
  const generationResult = generateWeeklySchedule(weekStartDate);

  // Get the target sheet (create it if it does not exist, reuse if it does).
  const scheduleSheet = getOrCreateWeekSheet(generationResult.weekSheetName);

  // Load staffing requirements for writing the summary footer and for passing to the formatter.
  const staffingRequirements = loadStaffingRequirements();

  // Write the grid to the sheet and apply all visual formatting.
  writeAndFormatSchedule(
    scheduleSheet,
    generationResult.employeeList,
    generationResult.weekGrid,
    staffingRequirements,
    weekStartDate,
    departmentName
  );

  return generationResult.weekSheetName;
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

    // Load settings — needed by the engine phases.
    const shiftTimingMap       = buildShiftTimingMap();
    const staffingRequirements = loadStaffingRequirements();

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

    // Re-write only the SHIFT rows and formatting.
    const departmentName = getDepartmentNameForHeader();

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

  // --- Settings sheet ---
  let settingsSheet = workbook.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = workbook.insertSheet(SHEET_NAMES.SETTINGS);
  }
  setupSettingsSheetTemplate(settingsSheet);
}


/**
 * Writes the header row to the Roster sheet and applies column formatting.
 *
 * @param {Sheet} rosterSheet — The Roster sheet object.
 */
function setupRosterSheetHeaders(rosterSheet) {
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
    "Seniority Rank"
  ]];

  const headerRange = rosterSheet.getRange(1, 1, 1, headerValues[0].length);
  headerRange.setValues(headerValues);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(COLORS.HEADER_BG);
  headerRange.setFontColor(COLORS.HEADER_TEXT);

  // Freeze the header row so it stays visible when scrolling.
  rosterSheet.setFrozenRows(1);

  // Column widths for readability.
  rosterSheet.setColumnWidth(ROSTER_COLUMN.NAME,             160);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.EMPLOYEE_ID,      110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.HIRE_DATE,        100);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.STATUS,            90);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.DAY_OFF_PREF_ONE, 110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.DAY_OFF_PREF_TWO, 110);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.PREFERRED_SHIFT,  120);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.QUALIFIED_SHIFTS, 200);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.VACATION_DATES,   200);
  rosterSheet.setColumnWidth(ROSTER_COLUMN.SENIORITY_RANK,   120);

  // Add a note to the Seniority Rank column header explaining it is script-managed.
  rosterSheet.getRange(1, ROSTER_COLUMN.SENIORITY_RANK).setNote(
    "This column is calculated and managed by the script.\n" +
    "Do not edit values here manually — they will be overwritten on the next Refresh Seniority run.\n\n" +
    "Higher number = more senior. Full-time employees always outrank part-time employees hired on the same date."
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
function setupSettingsSheetTemplate(settingsSheet) {
  // Only write the template if the sheet appears to be empty.
  if (settingsSheet.getLastRow() > 0) {
    return;
  }

  // --- Table 1: Staffing requirements (columns A–B) ---
  const staffingHeaders = [["Day", "Min Staff Required"]];
  const staffingData = [
    ["Monday",    6],
    ["Tuesday",   6],
    ["Wednesday", 6],
    ["Thursday",  6],
    ["Friday",    6],
    ["Saturday",  4],
    ["Sunday",    4],
  ];

  settingsSheet.getRange("A1:B1").setValues(staffingHeaders).setFontWeight("bold");
  settingsSheet.getRange("A2:B8").setValues(staffingData);

  // --- Table 2: Shift definitions (columns D–I) ---
  const shiftHeaders = [["Shift Name", "Status (FT/PT)", "Start Time", "End Time", "Paid Hours", "Has Lunch (TRUE/FALSE)"]];
  const shiftData = [
    // Each FT shift: 8.5-hour block (8 paid + 0.5 unpaid lunch)
    ["Morning",  "FT", "8:00 AM",   "4:30 PM",   8.0, true],
    ["Mid",      "FT", "10:00 AM",  "6:30 PM",   8.0, true],
    ["Closing",  "FT", "2:30 PM",   "11:00 PM",  8.0, true],
    // PT shifts: 5-hour paid, no lunch
    ["Morning",  "PT", "8:00 AM",   "1:00 PM",   5.0, false],
    ["Mid",      "PT", "10:00 AM",  "3:00 PM",   5.0, false],
    ["Closing",  "PT", "5:00 PM",   "10:00 PM",  5.0, false],
    // PT shifts with lunch: 5.5-hour block (5 paid + 0.5 unpaid lunch)
    ["Morning+", "PT", "8:00 AM",   "1:30 PM",   5.0, true],
    ["Mid+",     "PT", "10:00 AM",  "3:30 PM",   5.0, true],
  ];

  settingsSheet.getRange("D1:I1").setValues(shiftHeaders).setFontWeight("bold");
  settingsSheet.getRange("D2:I9").setValues(shiftData);

  // Style both header rows.
  settingsSheet.getRange("A1:B1").setBackground(COLORS.HEADER_BG).setFontColor(COLORS.HEADER_TEXT);
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
  const today      = new Date();
  const dayOfWeek  = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

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
      departmentValue === "— enter spreadsheet ID first —") {
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
  // Match the pattern "Week_MM_DD_YY" with exactly this format.
  const namePattern = /^Week_(\d{2})_(\d{2})_(\d{2})$/;
  const match       = sheetName.match(namePattern);

  if (!match) {
    return null;
  }

  const month     = parseInt(match[1], 10);
  const day       = parseInt(match[2], 10);
  const shortYear = parseInt(match[3], 10);

  // Convert two-digit year to four-digit year.
  // Years 00–99 map to 2000–2099. Adjust this logic if the tool is still in use in 2100.
  const fullYear = 2000 + shortYear;

  const parsedDate = new Date(fullYear, month - 1, day);
  parsedDate.setHours(0, 0, 0, 0);

  return isNaN(parsedDate.getTime()) ? null : parsedDate;
}
