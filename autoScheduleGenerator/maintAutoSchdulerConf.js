/**
 *       ___        _              _       _         ___     _           _      _         
 *      / __|___ __| |_ __ ___    /_\ _  _| |_ ___  / __| __| |_  ___ __| |_  _| |___ _ _ 
 *     | (__/ _ (_-<  _/ _/ _ \  / _ \ || |  _/ _ \ \__ \/ _| ' \/ -_) _` | || | / -_) '_|
 *      \___\___/__/\__\__\___/ /_/_\_\_,_|\__\___/ |___/\__|_||_\___\__,_|\_,_|_\___|_|  
 *      / __|___ _ _  / _(_)__ _  / __| __ _ _(_)_ __| |_                                 
 *     | (__/ _ \ ' \|  _| / _` | \__ \/ _| '_| | '_ \  _|                                
 *      \___\___/_||_|_| |_\__, | |___/\__|_| |_| .__/\__|                                
 *                         |___/                |_|                                       
 *
 * Version 0.0.3
 * Built by: Adam Roy
 * /


/* Global constants for column mapping */

const COLUMN_NAME      = 0; // COLUMN A
const COLUMN_ID        = 1; // COLUMN B
const COLUMN_DEPT      = 2; // COLUMN C
const COLUMN_HIRE_DATE = 5; // COLUMN F

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
 * Automatically runs whenever a cell is edited.
 * if an employee status changes it updates the weighting math unless overridden
 */
function onEdit(event) {
    const editedRange = event.range;
    const editedSheet = editedRange.getSheet();
    const editedColumn = editedRange.getColumn();
    const editedRow = editedRange.getRow();

    // only run this if we are on the CONFIG sheet and editing column 4 (employment status)
    if (editedSheet.getName() === "CONFIG" && editedColumn === 4 && editedRow > 1) {
        calculateEmployeeSeniority(editedRow);
    }
}