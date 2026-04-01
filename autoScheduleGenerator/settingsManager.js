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
 * Branch: shift-window-with-minimums
 * Version 0.2.1
 * /

/**
 * Fetches the minimum staffing requirement for a specific day from the SETTINGS sheet.
 */
function getMinimumStaffRequiredForDay(selectedDayName) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);

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
function getShiftTimingFromSettings(requestedValue, employmentStatus) {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = activeSpreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    const shiftData = settingsSheet.getRange("D2:H20").getValues();

    // Search for EXACT match first (Name + Status)
    for (let i = 0; i < shiftData.length; i++) {
        const [name, status, start, end, hours] = shiftData[i];
        if (name.toString().trim() === requestedValue.toString().trim() &&
            status.toString().trim() === employmentStatus.toString().trim()) {
            return formatShiftObject(start, end, hours);
        }
    }

    // FALLBACK: No match found — "NO SHIFT" flags a misconfigured preference in CONFIG or SETTINGS
    return { text: "NO SHIFT", hours: 0, status: employmentStatus };
}

/**
 * Reads all shift definitions from SETTINGS D2:H20 and returns a map keyed by "ShiftName|Status".
 * Avoids repeated sheet reads during resolution — call once and pass the result through all phases.
 */
function buildShiftTimingMap() {
    const settingsSheet  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    const shiftData      = settingsSheet.getRange("D2:H20").getValues();
    const shiftTimingMap = {};

    shiftData.forEach(shiftRow => {
        const [shiftName, employmentStatus, startTime, endTime, hours] = shiftRow;
        if (!shiftName || !employmentStatus || !startTime || !endTime) return;

        const startDate    = new Date(startTime);
        const endDate      = new Date(endTime);
        const startMinutes = (startDate.getHours() * 60) + startDate.getMinutes();
        const endMinutes   = (endDate.getHours() * 60) + endDate.getMinutes();
        const mapKey       = shiftName.toString().trim() + "|" + employmentStatus.toString().trim();

        shiftTimingMap[mapKey] = {
            startMinutes: startMinutes,
            endMinutes:   endMinutes,
            hours:        Number(hours),
            text:         Utilities.formatDate(startDate, TIME_ZONE, TIME_FORMAT_STRING) + " - " +
                          Utilities.formatDate(endDate, TIME_ZONE, TIME_FORMAT_STRING)
        };
    });

    return shiftTimingMap;
}

/**
 * Returns the deduplicated list of shift names from the SETTINGS sheet (column D).
 * Used to populate dropdowns in CONFIG without hardcoding shift names.
 */
function readShiftNamesFromSettings() {
    const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    const raw = settingsSheet.getRange("D2:D20").getValues().flat().filter(v => v !== "");
    return [...new Set(raw)];
}

// K.I.S.S. helper code
function formatShiftObject(start, end, hours) {
    if (!start || !end) return { text: "OFF", hours: 0 };
    const timeStr = Utilities.formatDate(new Date(start), TIME_ZONE, TIME_FORMAT_STRING) + " - " +
        Utilities.formatDate(new Date(end), TIME_ZONE, TIME_FORMAT_STRING);
    return { text: timeStr, hours: Number(hours) };
}
