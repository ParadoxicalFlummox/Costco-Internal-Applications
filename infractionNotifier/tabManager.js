/**
 * tabManager.js — Employee tab and spreadsheet creation from the Employee Details roster.
 * VERSION: 0.1.0
 *
 * This file owns all functions for generating individual employee attendance
 * controller tabs or standalone spreadsheets from the "(Employee Details)" roster sheet.
 *
 * TWO CREATION MODES:
 *
 *   TAB MODE (buildtabs / buildtabs2):
 *     Creates a new sheet tab inside the current workbook for each unprocessed
 *     employee. Uses either the color-coded or plain template sheet as the source.
 *     Best for a single consolidated workbook where all employees share one file.
 *
 *   INDIVIDUAL SPREADSHEET MODE (buildSheetColor / buildSheetNoColor):
 *     Creates a separate Google Spreadsheet file for each unprocessed employee
 *     and saves them all into a Drive folder named "2026 Attendance Controllers".
 *     Best when each employee's controller is managed independently or shared
 *     with different people.
 *
 * IDEMPOTENCY:
 *   Both modes read column F of "(Employee Details)" for a "Completed" flag.
 *   Rows already marked Completed are skipped, so re-running will not create
 *   duplicate tabs or files for employees who were already processed.
 *
 * TEMPLATE SHEETS (must exist in the workbook):
 *   "(Color-coded Electronic)" — color-coded attendance grid template
 *   "(Electronic - No Color)"  — plain attendance grid template
 *
 * EMPLOYEE DETAILS SHEET LAYOUT (row 1 = headers, data starts row 2):
 *   A — Employee Number (ID)
 *   B — Last Name
 *   C — First Name
 *   D — Hire Date
 *   E — Department
 *   F — Completed flag (written by this script after tab/file creation)
 */


// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const TAB_MANAGER_SOURCE_SHEET      = '(Employee Details)';
const TAB_MANAGER_TEMPLATE_COLOR    = '(Color-coded Electronic)';
const TAB_MANAGER_TEMPLATE_NO_COLOR = '(Electronic - No Color)';
const TAB_MANAGER_DRIVE_FOLDER      = '2026 Attendance Controllers';


// ---------------------------------------------------------------------------
// Tab Mode — Create tabs inside this workbook
// ---------------------------------------------------------------------------

/**
 * Creates one color-coded attendance tab per unprocessed employee inside the
 * current workbook, using "(Color-coded Electronic)" as the template.
 *
 * Each new tab is named "Last, First - EmployeeID" and pre-populated with
 * the employee's name, ID, hire date, and department in the standard cells.
 */
function buildtabs() {
  buildTabsFromTemplate_(TAB_MANAGER_TEMPLATE_COLOR);
}

/**
 * Creates one plain (no color) attendance tab per unprocessed employee inside
 * the current workbook, using "(Electronic - No Color)" as the template.
 */
function buildtabs2() {
  buildTabsFromTemplate_(TAB_MANAGER_TEMPLATE_NO_COLOR);
}

/**
 * Core tab creation logic shared by buildtabs and buildtabs2.
 *
 * Reads all rows from "(Employee Details)", skips any row where column F is
 * already "Completed", creates a sheet tab from the given template for each
 * remaining row, populates the standard metadata cells, and marks the row
 * as Completed to prevent re-creation on future runs.
 *
 * @param {string} templateName — The name of the template sheet tab to copy.
 */
function buildTabsFromTemplate_(templateName) {
  const workbook     = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet  = workbook.getSheetByName(TAB_MANAGER_SOURCE_SHEET);
  const templateSheet = workbook.getSheetByName(templateName);

  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${TAB_MANAGER_SOURCE_SHEET}" not found.`);
    return;
  }
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert(`Template sheet "${templateName}" not found.`);
    return;
  }

  const rows      = rosterSheet.getRange('A2:F').getValues();
  let   created   = 0;
  let   skipped   = 0;

  rows.forEach((row, index) => {
    const employeeId = row[0];
    const lastName   = row[1];
    const firstName  = row[2];
    const hireDate   = row[3];
    const department = row[4];
    const completed  = row[5];

    if (!employeeId || !lastName || !firstName) return; // blank row
    if (completed === 'Completed') { skipped++; return; }

    // Tab name follows "Last, First - ID" convention used throughout the system
    const tabName = `${lastName}, ${firstName} - ${employeeId}`;

    // Skip if a tab with this name already exists (safety guard)
    if (workbook.getSheetByName(tabName)) {
      rosterSheet.getRange(index + 2, 6).setValue('Completed');
      skipped++;
      return;
    }

    const newSheet = workbook.insertSheet(tabName, { template: templateSheet });

    // Populate the standard metadata cells
    newSheet.getRange('X1').setValue(`${lastName} ${firstName}`);
    newSheet.getRange('X3').setValue(employeeId);
    newSheet.getRange('AD3').setValue(hireDate);
    newSheet.getRange('R3').setValue(department);

    // Mark as completed to prevent re-creation on future runs
    rosterSheet.getRange(index + 2, 6).setValue('Completed');
    created++;
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Created ${created} tab(s). ${skipped} already completed.`,
    'Tab Creation Done',
    6
  );
  console.log(`tabManager: buildTabsFromTemplate_ (${templateName}) — created ${created}, skipped ${skipped}.`);
}


// ---------------------------------------------------------------------------
// Individual Spreadsheet Mode — Create separate files per employee
// ---------------------------------------------------------------------------

/**
 * Creates one color-coded standalone spreadsheet per unprocessed employee and
 * saves each to the "2026 Attendance Controllers" Drive folder.
 */
function buildSheetColor() {
  buildIndividualSheets_(TAB_MANAGER_TEMPLATE_COLOR);
}

/**
 * Creates one plain (no color) standalone spreadsheet per unprocessed employee
 * and saves each to the "2026 Attendance Controllers" Drive folder.
 */
function buildSheetNoColor() {
  buildIndividualSheets_(TAB_MANAGER_TEMPLATE_NO_COLOR);
}

/**
 * Core individual spreadsheet creation logic shared by buildSheetColor and
 * buildSheetNoColor.
 *
 * For each unprocessed employee:
 *   1. Creates a new Google Spreadsheet named "Last, First EmployeeID - 2026 Attendance".
 *   2. Copies the template sheet into it and deletes the default blank Sheet1.
 *   3. Populates the standard metadata cells.
 *   4. Moves the file into the Drive folder (creates the folder if it doesn't exist).
 *   5. Marks the "(Employee Details)" row as Completed.
 *
 * @param {string} templateName — The name of the template sheet tab to copy.
 */
function buildIndividualSheets_(templateName) {
  const workbook      = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet   = workbook.getSheetByName(TAB_MANAGER_SOURCE_SHEET);
  const templateSheet = workbook.getSheetByName(templateName);

  if (!rosterSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${TAB_MANAGER_SOURCE_SHEET}" not found.`);
    return;
  }
  if (!templateSheet) {
    SpreadsheetApp.getUi().alert(`Template sheet "${templateName}" not found.`);
    return;
  }

  // Resolve the target Drive folder, creating it if needed
  const targetFolder = getOrCreateDriveFolder_(TAB_MANAGER_DRIVE_FOLDER);

  const lastRow  = rosterSheet.getLastRow();
  const rows     = rosterSheet.getRange(`A2:F${lastRow}`).getValues();
  let   created  = 0;
  let   skipped  = 0;

  rows.forEach((row, index) => {
    const employeeId = row[0];
    const lastName   = row[1];
    const firstName  = row[2];
    const hireDate   = row[3];
    const department = row[4];
    const completed  = row[5];

    if (!employeeId || !lastName || !firstName) return;
    if (completed === 'Completed') { skipped++; return; }

    const fileName  = `${lastName}, ${firstName} ${employeeId} - 2026 Attendance`;
    const tabName   = `${lastName}, ${firstName} - ${employeeId}`;

    // Create the new spreadsheet
    const newSpreadsheet  = SpreadsheetApp.create(fileName);
    const blankSheet      = newSpreadsheet.getActiveSheet(); // default Sheet1

    // Copy template into the new workbook
    templateSheet.copyTo(newSpreadsheet);
    const copiedSheet = newSpreadsheet.getSheets()[1]; // index 1 = the just-copied sheet

    // Populate metadata cells
    copiedSheet.getRange('X1').setValue(`${lastName} ${firstName}`);
    copiedSheet.getRange('X3').setValue(employeeId);
    copiedSheet.getRange('AD3').setValue(hireDate);
    copiedSheet.getRange('R3').setValue(department);

    // Name the sheet tab consistently with the in-workbook tab convention
    copiedSheet.setName(tabName);

    // Remove the default blank Sheet1
    newSpreadsheet.deleteSheet(blankSheet);

    // Move the file into the target folder and out of Drive root
    const file = DriveApp.getFileById(newSpreadsheet.getId());
    targetFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    // Mark as completed
    rosterSheet.getRange(index + 2, 6).setValue('Completed');
    created++;

    console.log(`tabManager: Created spreadsheet "${fileName}".`);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Created ${created} spreadsheet(s) in "${TAB_MANAGER_DRIVE_FOLDER}". ${skipped} already completed.`,
    'Sheet Creation Done',
    8
  );
  console.log(`tabManager: buildIndividualSheets_ (${templateName}) — created ${created}, skipped ${skipped}.`);
}


// ---------------------------------------------------------------------------
// Sort Tabs
// ---------------------------------------------------------------------------

/**
 * Sorts all sheet tabs in the workbook alphabetically by name.
 *
 * Hidden sheets (those whose names start with "(") sort to the top
 * alphabetically and remain in place. Visible employee tabs end up sorted
 * "Last, First - ID" order automatically.
 */
function sortTabs() {
  const workbook   = SpreadsheetApp.getActiveSpreadsheet();
  const sheets     = workbook.getSheets();
  const sheetNames = sheets.map(s => s.getName()).sort();

  sheetNames.forEach((name, position) => {
    workbook.setActiveSheet(workbook.getSheetByName(name));
    workbook.moveActiveSheet(position + 1);
  });

  console.log(`tabManager: Sorted ${sheets.length} tab(s) alphabetically.`);
}


// ---------------------------------------------------------------------------
// Drive Utility
// ---------------------------------------------------------------------------

/**
 * Returns the Drive folder with the given name, creating it if it does not
 * yet exist. If multiple folders share the same name, the first one found
 * is returned.
 *
 * @param {string} folderName
 * @returns {GoogleAppsScript.Drive.Folder}
 */
function getOrCreateDriveFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  console.log(`tabManager: Creating Drive folder "${folderName}".`);
  return DriveApp.createFolder(folderName);
}
