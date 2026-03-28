/**
 *       _____             _                                _            _____        _                _         _             
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |            
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __ 
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |   
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|   
 *       _____        _    _    _                      __  __                                                                  
 *      / ____|      | |  | |  (_)                    |  \/  |                                                                 
 *     | (___    ___ | |_ | |_  _  _ __    __ _  ___  | \  / |  __ _  _ __    __ _   __ _   ___  _ __                          
 *      \___ \  / _ \| __|| __|| || '_ \  / _` |/ __| | |\/| | / _` || '_ \  / _` | / _` | / _ \| '__|                         
 *      ____) ||  __/| |_ | |_ | || | | || (_| |\__ \ | |  | || (_| || | | || (_| || (_| ||  __/| |                            
 *     |_____/  \___| \__| \__||_||_| |_| \__, ||___/ |_|  |_| \__,_||_| |_| \__,_| \__, | \___||_|                            
 *                                         __/ |                                     __/ |                                     
 *                                        |___/                                     |___/                                      
 * Built by: Adam Roy
 * Version 0.0.11
 * /

/**
 * Fetches the minimum staffing requirement for a specific day from the SETTINGS sheet.
 */
function getMinimumStaffRequiredForDay(selectedDayName) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet     = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    // Read the Staffing Table (A2:B8)
    const staffingRequirementData = settingsSheet.getRange("A2:B8").getValues();
    
    for (let rowIndex = 0; rowIndex < staffingRequirementData.length; rowIndex++) {
        const dayInTable = staffingRequirementData[rowIndex][0];
        const staffCount = staffingRequirementData[rowIndex][1];
        
        if (dayInTable === selectedDayName) {
            return staffCount;
        }
    }
    return 6; // Safety fallback if day is not found
}

/**
 * Looks up the specific Start and End times based on Shift Preference and Employment Status.
 */
function getShiftTimingFromSettings(requestedShiftName, employmentStatus) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet     = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    // Read the Shift Definition Table (D2:G10)
    const shiftDefinitionData = settingsSheet.getRange("D2:H10").getValues();
    
    for (let rowIndex = 0; rowIndex < shiftDefinitionData.length; rowIndex++) {
        const tableShiftName = shiftDefinitionData[rowIndex][0].toString().trim();
        const tableStatus    = shiftDefinitionData[rowIndex][1].toString().trim();
        
        if (tableShiftName === requestedShiftName && tableStatus === employmentStatus) {
            const startTime = shiftDefinitionData[rowIndex][2];
            const endTime   = shiftDefinitionData[rowIndex][3];
            const hours     = shiftDefinitionData[rowIndex][4];

            //Safety check, checks if a cell is empty if it is dont return a 12:00 AM
            if (!startTime || !endTime) return {text: "OFF", hours: 0};
            
            // Use HH:mm for 24-hour time to match spreadsheet default
            const timeString = Utilities.formatDate(new Date(start), "GMT", TIME_FORMAT_STRING) + 
                               " - " + 
                               Utilities.formatDate(new Date(end), "GMT", TIME_FORMAT_STRING);
            
            return {text: timeString, hours: hours};
        }
    }
    return { text: "OFF", hours: 0 };
}