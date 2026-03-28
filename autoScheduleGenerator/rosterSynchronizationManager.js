/**
 *       _____             _                                _            _____        _                _         _             
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |            
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __ 
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |   
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|   
 *      _____              _                 _____                       __  __                                                
 *     |  __ \            | |               / ____|                     |  \/  |                                               
 *     | |__) | ___   ___ | |_  ___  _ __  | (___   _   _  _ __    ___  | \  / |  __ _  _ __    __ _   __ _   ___  _ __        
 *     |  _  / / _ \ / __|| __|/ _ \| '__|  \___ \ | | | || '_ \  / __| | |\/| | / _` || '_ \  / _` | / _` | / _ \| '__|       
 *     | | \ \| (_) |\__ \| |_|  __/| |     ____) || |_| || | | || (__  | |  | || (_| || | | || (_| || (_| ||  __/| |          
 *     |_|  \_\\___/ |___/ \__|\___||_|    |_____/  \__, ||_| |_| \___| |_|  |_| \__,_||_| |_| \__,_| \__, | \___||_|          
 *                                                   __/ |                                             __/ |                   
 *                                                  |___/                                             |___/                    
 * Built by: Adam Roy
 * Version 0.0.11
 * /

/**
 * Pulls new employees from the master 'Input' sheet into the CONFIG sheet.
 */
function synchronizeEmployeeRoster() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const masterDataSheet   = activeSpreadsheet.getSheetByName("Input"); 
    const configurationSheet = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);

    const masterDataValues = masterDataSheet.getDataRange().getValues();
    const existingEmployeeIds = configurationSheet.getRange("B:B").getValues().flat().map(String);

    masterDataValues.forEach((currentRow, rowIndex) => {
        if (rowIndex === 0) return; // Skip Header
        
        const employeeIdFromMaster = currentRow[MASTER_COLUMN_ID].toString();
        const departmentFromMaster = currentRow[MASTER_COLUMN_DEPARTMENT];

        if (departmentFromMaster === TARGET_DEPARTMENT_NAME && !existingEmployeeIds.includes(employeeIdFromMaster)) {
            const nextAvailableRow = configurationSheet.getLastRow() + 1;
            
            const newEmployeeData = [[
                currentRow[MASTER_COLUMN_NAME],
                employeeIdFromMaster,
                currentRow[MASTER_COLUMN_HIRE_DATE],
                "PT", // Default Status
                "",   // Preference 1
                ""    // Preference 2
            ]];

            configurationSheet.getRange(nextAvailableRow, 1, 1, 6).setValues(newEmployeeData);
            calculateEmployeeSeniority(nextAvailableRow);
        }
    });
    applyDataValidationRules();
}

/**
 * Calculates the seniority rank for a specific row.
 */
function calculateEmployeeSeniority(rowNumber) {
    const configurationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURATION_SHEET_NAME);
    const rowData = configurationSheet.getRange(rowNumber, 1, 1, 4).getValues()[0];
    
    const hireDateValue    = new Date(rowData[2]);
    const employmentStatus = rowData[3];

    if (isNaN(hireDateValue.getTime())) return;

    const statusMultiplier = (employmentStatus === "FT") ? 200000000 : 100000000;
    const referenceDateValue = new Date("2050-01-01").getTime();
    const seniorityDaysCount = Math.floor((referenceDateValue - hireDateValue.getTime()) / (1000 * 60 * 60 * 24));

    const finalSeniorityRank = statusMultiplier + seniorityDaysCount;
    configurationSheet.getRange(rowNumber, COLUMN_INDEX_SENIORITY_RANK + 1).setValue(finalSeniorityRank);
}

/**
 * Applies dropdown menus (Data Validation) to the CONFIG sheet columns.
 * This ensures managers can only pick valid Statuses and Days.
 */
function applyDataValidationRules() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configurationSheet = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);
    const lastRowInConfiguration = configurationSheet.getLastRow();
    
    // Protect the script from running on an empty sheet
    if (lastRowInConfiguration < 2) return; 

    // 1. Define the FT/PT Dropdown (Column D / Index 3)
    const statusValidationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["FT", "PT"], true)
        .setHelpText("Please select FT or PT.")
        .build();
    configurationSheet.getRange(2, 4, lastRowInConfiguration - 1).setDataValidation(statusValidationRule);

    // 2. Define the Days Off Dropdown (Columns E & F / Index 4 & 5)
    const dayNamesList = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
    const daysValidationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(dayNamesList, true)
        .build();
    configurationSheet.getRange(2, 5, lastRowInConfiguration - 1, 2).setDataValidation(daysValidationRule);
    
    // 3. Define the Shift Preference Dropdown (Column H / Index 7)
    // This matches the "Shift Name" in your SETTINGS sheet
    const shiftNamesList = ["Morning", "Mid", "Night"];
    const shiftValidationRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(shiftNamesList, true)
        .build();
    configurationSheet.getRange(2, 8, lastRowInConfiguration - 1).setDataValidation(shiftValidationRule);

    Logger.log("Data validation rules successfully applied to CONFIG sheet.");
}