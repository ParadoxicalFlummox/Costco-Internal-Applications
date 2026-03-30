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
 * Version 0.0.16
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
    const shiftData = settingsSheet.getRange("D2:H10").getValues();

    let fallbackShift = null;
    
    for (let i = 0; i < shiftData.length; i++) {
        const [shiftName, status, start, end, hours] = shiftData[i];
        if (!shiftName) continue;

        // Exact match case: if a manager chooses a specific shift
        if (shiftName.toString().trim() === requestedShift.toString().trim() && status === employmentStatus) {
            return formatShiftObject(start, end, hours);
        }

        // Window match case: if a manager chooses a specific shift window
        if (shiftName.toString().trim() == requestedShift.toString().trim() && !fallbackShift) {
            fallbackShift = formatShiftObject(start, end, hours);
        }
    }

    return fallbackShift || {text: "OFF", hours: 0};
}

// K.I.S.S. helper code
function formatShiftObject(start, end, hours) {
    if (!start || !end) return {text: "OFF", hours: 0};
    const timeStr = Utilities.formatDate(new Date(start), TIME_ZONE, TIME_FORMAT_STRING) + " - " +
                    Utilities.formatDate(new Date(end), TIME_ZONE, TIME_FORMAT_STRING);
    return {text: timeStr, hours: Number(hours) };
}
