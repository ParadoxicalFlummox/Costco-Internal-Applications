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
 * Version 0.0.15
 * /

/**
 * Fetches the minimum staffing requirement for a specific day from the SETTINGS sheet.
 */
function getMinimumStaffRequiredForDay(selectedDayName) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet     = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    // Read the Staffing Table (A2:B8)
    const staffingRequirementData = settingsSheet.getRange("A2:B8").getValues();
    
    for (let i = 0; i < staffingRequirementData.length; i++) {
        const dayInTable = staffingRequirementData[i][0];
        const staffCount = staffingRequirementData[i][1];
        
        if (dayInTable === selectedDayName) {
            return staffCount;
        }
    }
    return 6; // Safety fallback if day is not found
}

/**
 * Looks up the specific Start and End times based on Shift Preference and Employment Status.
 */
function getShiftTimingFromSettings(requestedShift, employmentStatus) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet     = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    
    // Read the Shift Definition Table (D2:G10)
    const shiftData = settingsSheet.getRange("D2:H10").getValues();

    // Normalize data for comparison
    const searchShift = requestedShift.toString().trim();
    const searchStatus = employmentStatus.toString().trim();
    
    for (let i = 0; i < shiftData.length; i++) {
        const tableShift  = shiftData[i][0].toString().trim();
        const tableStatus = shiftData[i][1].toString().trim();
        
        if (tableShift === searchShift && tableStatus === searchStatus) {
            const startTime = shiftData[i][2];
            const endTime   = shiftData[i][3];
            const hours     = shiftData[i][4];

            //Safety check, checks if a cell is empty if it is dont return a 12:00 AM
            if (!startTime || !endTime) return {text: "OFF", hours: 0};
            
            // Use HH:mm for 24-hour time to match spreadsheet default
            const timeString = Utilities.formatDate(new Date(startTime), "GMT", TIME_FORMAT_STRING) + 
                               " - " + 
                               Utilities.formatDate(new Date(endTime), "GMT", TIME_FORMAT_STRING);
            
            return {text: timeString, hours: hours};
        }
    }
    Logger.log("WARNING: No shift match found for " + searchShift + " / " + searchStatus);
    return { text: "OFF", hours: 0 };
}