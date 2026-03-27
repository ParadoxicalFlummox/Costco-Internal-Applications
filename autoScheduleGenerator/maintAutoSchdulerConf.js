/**
 *       ___        _              _       _         ___     _           _      _         
 *      / __|___ __| |_ __ ___    /_\ _  _| |_ ___  / __| __| |_  ___ __| |_  _| |___ _ _ 
 *     | (__/ _ (_-<  _/ _/ _ \  / _ \ || |  _/ _ \ \__ \/ _| ' \/ -_) _` | || | / -_) '_|
 *      \___\___/__/\__\__\___/ /_/_\_\_,_|\__\___/ |___/\__|_||_\___\__,_|\_,_|_\___|_|                                         
 *
 * Version 0.0.4
 * Built by: Adam Roy
 * /


/* Global constants for column mapping */

const COLUMN_NAME      = 0; // COLUMN A
const COLUMN_ID        = 1; // COLUMN B
const COLUMN_DEPT      = 2; // COLUMN C
const COLUMN_HIRE_DATE = 5; // COLUMN F

/**
 * Automatically runs whenever a cell is edited.
 */
function onEdit(event) {
    const editedRange  = event.range;
    const editedSheet  = editedRange.getSheet();
    const sheetName    = editedSheet.getName();
    const editedColumn = editedRange.getColumn();
    const editedRow    = editedRange.getRow();

    // Logic for edits to the CONFIG sheet (employment status changes)
    if (editedSheet.getName() === "CONFIG" && editedColumn === 4 && editedRow > 1) {
        calculateEmployeeSeniority(editedRow);
        return;
    }

    // Logic for weekly schedule edits (vacation and or preference changes)
    if (sheetName.startsWith("Week_") && editedColumn >= 3 && editedColumn <= 9 && editedRow >= 6) {
        const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
        const dayName = dayNames[editedColumn - 3];

        // trigger the conflict resolver for the day that was edited
        resolveScheduleConflicts(editedSheet, dayName, editedColumn);
    }
}

function calculateWeekLabel() {
    const todayDate = new Date();
    const dayOfWeekIndex = todayDate.getDay();
    const daysToSubtract = (dayOfWeekIndex === 0) ? 6 : dayOfWeekIndex - 1;
    const weekBeginningDate = new Date(todayDate);
    weekBeginningDate.setDate(todayDate.getDate() - daysToSubtract);

    // component formatting
    const calendarMonth = weekBeginningDate.getMonth() + 1;
    const calendarDay = weekBeginningDate.getDate();
    const shortYear = String(weekBeginningDate.getFullYear()).slice(2);

    // Ensure leading zeros look optimal
    const formattedMonth = (calendarMonth < 10) ? "0" + calendarMonth : calendarMonth;
    const formattedDay = (calendarDay < 10) ? "0" + calendarDay : calendarDay;
    const weekLabel = "Week_" + formattedMonth + "_" + formattedDay + "_" + shortYear;

    return weekLabel;
}

/* 
 * Syncs staff from the master sheet into CONFIG.
 * This data is persistent meaning it wont be wiped out when a new schedule is generated
 */

function synchronizeRoster() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterDataSheet = activeSpreadsheet.getSheetByName("Input"); // this references an internal sheet for testing
    
    /* switch to using these to pull from the master employee list
    const externalFileId = "the_file_id"
    const activeSpreadsheet = SpreadsheetApp.openById(externalFileId);
    const masterDataSheet = masterDataFile.getSheetByName("name_of_list_tab")
    */
    const configurationSheet = activeSpreadsheet.getSheetByName("CONFIG");

    // Get the data from both sheets
    const masterValues = masterDataSheet.getDataRange().getValues();
    const configurationValues = configurationSheet.getDataRange().getValues();

    // Create a list of employee IDs currently in CONFIG to avoid making duplicates
    const existingConfigIds =  configurationValues.map(function(row) {
        return row[1].toString();
    });

    const targetDepartment = "Maintenance"; // This will eventually be mapped to a dropdown so it can be used for any department

    masterValues.forEach(function(currentRow, rowIndex) {
        if (rowIndex === 0) return; // Skips the title headers

        const employeeName     = currentRow[COLUMN_NAME];
        const employeeId       = currentRow[COLUMN_ID];
        const employeeDept     = currentRow[COLUMN_DEPT];
        const employeeHireDate = currentRow[COLUMN_HIRE_DATE];

        if (employeeDept === targetDepartment && !existingConfigIds.includes(employeeId)) {
            const lastNameInCol = configurationSheet.getRange("A:A").getValues().filter(String);
            const nextRow = lastNameInCol.length + 1;
            // Append: [Name, ID, Hire Date, Default Status, Pref 1, Pref 2]
            configurationSheet.getRange(nextRow, 1, 1, 6).setValues([[
                employeeName,
                employeeId,
                employeeHireDate,
                "PT",
                "",
                ""
            ]]);
            Logger.log("Synchronized new employee: " + employeeName);
            
            // now calculate the employee seniority weighting
            calculateEmployeeSeniority(nextRow);
        }
    });

    applyDataValidationToConfig();
}

/**
 * Adds dropdown menus to the CONFIG sheet for employment status and days off.
 */
function applyDataValidationToConfig(){
    const configurationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const lastRow = configurationSheet.getLastRow();
    if (lastRow < 2) return; // protects the script from crashing if the spreadsheet is empty

    // FT/PT dropdown (Column D)
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(["FT", "PT"]).build();
    configurationSheet.getRange(2, 4, lastRow - 1).setDataValidation(statusRule);

    // Days off dropdown (columns E and F)
    // ** may use a set of buttons in the future if its easier to read **
    const daysRule = SpreadsheetApp.newDataValidation().requireValueInList(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]).build();
    configurationSheet.getRange(2, 5, lastRow - 1, 2).setDataValidation(daysRule);
}
/**
 * Helper function that calculates employee seniority for a specific row
 */
function calculateEmployeeSeniority(rowNumber) {
    const configurationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
    const rowData = configurationSheet.getRange(rowNumber, 1, 1, 4).getValues()[0];

    const hireDateValue = rowData[2];
    const status = rowData[3];

    if (!hireDateValue) return;

    // Ensure that we are working with a valid date object
    const hireDate = new Date(hireDateValue);

    const statusMultiplier =  (status === "FT") ? 200000000 : 100000000;
    const referenceDate = new Date("2050-01-01").getTime();

    // Calculate days, if this is invalid it results in NaN.
    const seniorityDays = Math.floor((referenceDate - hireDate.getTime()) / (1000 * 60 * 60 * 24));

    if (isNaN(seniorityDays)) {
        Logger.log("Warning: Invalid date found at row " + rowNumber);
        return;
    }

    const finalRank = statusMultiplier + seniorityDays;

    // Write the result to column G
    configurationSheet.getRange(rowNumber, 7).setValue(finalRank);
}

/**
 * The clickable generate button and the logic it executes.
 * Cloning the template sheet, renaming it and the populating the schedule.
 */
function generateWeeklySchedule() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configurationSheet = activeSpreadsheet.getSheetByName("CONFIG")
    const templateSheet = activeSpreadsheet.getSheetByName("Grid Template"); // this can be named whatever you want the template sheet to be named

    // Create a unique label for the week
    const weekLabel = calculateWeekLabel();

    // Clean existing table if re-running script
    let newScheduleSheet = activeSpreadsheet.getSheetByName(weekLabel);
    if (newScheduleSheet) {
        activeSpreadsheet.deleteSheet(newScheduleSheet);
    }
    newScheduleSheet = templateSheet.copyTo(activeSpreadsheet).setName(weekLabel);
    newScheduleSheet.showSheet();

    // Pull and sort the roster by seniority rank
    const rosterData = configurationSheet.getRange(2, 1, configurationSheet.getLastRow() - 1, 7).getValues();
    rosterData.sort((a, b) => b[6] - a[6]);

    let destinationRow = 6;
    const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    rosterData.forEach(function(employee) {
        const employeeName = employee[0];
        const pref1 = employee[4];
        const pref2 = employee[5];

        // Name injection and merge rows for a cleaner look
        newScheduleSheet.getRange(destinationRow, 2).setValue(employeeName);
        newScheduleSheet.getRange(destinationRow, 2, 2, 1).merge().setVerticalAlignment("center");

        // Insert the checkboxes for both vacation (top row) and preferred days off (bottom row)
        const checkboxRange = newScheduleSheet.getRange(destinationRow, 3, 2, 7);
        checkboxRange.insertCheckboxes();

        // Set the initial preferences from CONFIG
        [pref1, pref2].forEach(pref => {
            if(pref) {
                const dayIndex = dayNames.indexOf(pref);
                if (dayIndex !== -1) {
                    newScheduleSheet.getRange(destinationRow + 1, 3 + dayIndex).setValue(true);
                }
            }
        });
        destinationRow += 2;
    });

    // Finalize by adding summary and run the initial conflict resolution function for each day
    addSummary(newScheduleSheet, destinationRow -1);

    for (let col = 3; col <= 9; col++) {
        resolveScheduleConflicts(newScheduleSheet, dayNames[col - 3], col);
    }
    activeSpreadsheet.setActiveSheet(newScheduleSheet)
}

/**
 * The Seniority Engine: Grants or denies "Off" requests based on staffing floors.
 */
function resolveScheduleConflicts(scheduleSheet, dayName, columnIndex) {
  const minRequired = getMinStaffRequired(dayName);
  const lastRow = scheduleSheet.getLastRow();
  
  // Find only rows where employees are listed (B6 down to Summary)
  // We determine the last employee row by looking for the last merged name in B
  const employeeRowsCount = (scheduleSheet.getRange("B6:B" + lastRow).getValues().filter(String).length) * 2;
  const maxAllowedOff = (employeeRowsCount / 2) - minRequired;

  let currentOffCount = 0;
  let preferenceRows = [];

  // Scan Column for Vacation (Unchangeable) and Preferences (Changeable)
  for (let row = 6; row < 6 + employeeRowsCount; row += 2) {
    const isVacation = scheduleSheet.getRange(row, columnIndex).getValue();
    const wantsOff = scheduleSheet.getRange(row + 1, columnIndex).getValue();

    if (isVacation) {
      currentOffCount++; // Vacation takes a "slot" first
    } else if (wantsOff) {
      preferenceRows.push(row + 1); // Save the row index for later processing
    }
  }

  // Resolve based on Rank (already sorted by row order)
  preferenceRows.forEach(prefRow => {
    if (currentOffCount < maxAllowedOff) {
      currentOffCount++;
      scheduleSheet.getRange(prefRow, columnIndex).setBackground(null).setNote(null);
    } else {
      // BUMPED: Too many requests, junior employee must work
      scheduleSheet.getRange(prefRow, columnIndex).setValue(false).setBackground("#fff2cc").setNote("Staffing Limit: Must Work");
    }
  });
}

/**
 * Adds the OK/UNDER status block at the bottom of the weekly sheet.
 */
function addSummary(scheduleSheet, lastEmployeeRow) {
  const summaryRow = lastEmployeeRow + 2;
  const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  
  scheduleSheet.getRange(summaryRow, 2, 3, 1).setValues([["REQUIRED"], ["ACTUAL"], ["STATUS"]]).setFontWeight("bold");

  for (let i = 0; i < 7; i++) {
    const colIndex = 3 + i;
    const colLetter = String.fromCharCode(67 + i);
    const minRequired = getMinStaffRequired(dayNames[i]);

    scheduleSheet.getRange(summaryRow, colIndex).setValue(minRequired);
    
    // Formula: Total Employees - Vacation Checks - Preference Checks
    const totalStaff = (lastEmployeeRow - 5) / 2;
    const formula = `=${totalStaff} - COUNTIF(${colLetter}6:${colLetter}${lastEmployeeRow-1}, TRUE) - COUNTIF(${colLetter}7:${colLetter}${lastEmployeeRow}, TRUE)`;
    scheduleSheet.getRange(summaryRow + 1, colIndex).setFormula(formula);

    // Status visual alert
    const statusFormula = `=IF(${colLetter}${summaryRow + 1} >= ${colLetter}${summaryRow}, "OK", "UNDER")`;
    scheduleSheet.getRange(summaryRow + 2, colIndex).setFormula(statusFormula);
  }
}

/**
 * Helper: Fetches min staff from K1:L8 in CONFIG.
 */
function getMinStaffRequired(dayName) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CONFIG");
  const data = configSheet.getRange("K2:L8").getValues();
  for (let row of data) {
    if (row[0] === dayName) return row[1];
  }
  return 6; // Safety fallback
}