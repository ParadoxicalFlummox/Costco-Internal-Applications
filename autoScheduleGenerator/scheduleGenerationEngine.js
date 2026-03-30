/**
 *       _____             _                                _            _____        _                _         _             
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |            
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __ 
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |   
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|   
 *       _____        _                _         _          _____                ______                _                       
 *      / ____|      | |              | |       | |        / ____|              |  ____|              (_)                      
 *     | (___    ___ | |__    ___   __| | _   _ | |  ___  | |  __   ___  _ __   | |__    _ __    __ _  _  _ __    ___          
 *      \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \ | | |_ | / _ \| '_ \  |  __|  | '_ \  / _` || || '_ \  / _ \         
 *      ____) || (__ | | | ||  __/| (_| || |_| || ||  __/ | |__| ||  __/| | | | | |____ | | | || (_| || || | | ||  __/         
 *     |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___|  \_____| \___||_| |_| |______||_| |_| \__, ||_||_| |_| \___|         
 *                                                                                               __/ |                         
 *                                                                                              |___/                          
 * Built by: Adam Roy
 * Version 0.0.16
 * /

/**
 * Creates a new Weekly Schedule based on the Template.
 */
function generateWeeklySchedule() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configurationSheet = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);
    const templateSheet = activeSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

    const weeklySheetLabel = generateWeeklyDateLabel();
    let newScheduleSheet = activeSpreadsheet.getSheetByName(weeklySheetLabel) || 
                           templateSheet.copyTo(activeSpreadsheet).setName(weeklySheetLabel);
    
    newScheduleSheet.showSheet();

    const rosterData = configurationSheet.getRange(2, 1, configurationSheet.getLastRow() - 1, 8).getValues();
    rosterData.sort((a, b) => b[COLUMN_INDEX_SENIORITY_RANK] - a[COLUMN_INDEX_SENIORITY_RANK]);

    let currentDestinationRow = 6;
    const weekDayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    rosterData.forEach(employeeRecord => {
        // 1. Labels (A) and Name (B) - Minimal formatting
        newScheduleSheet.getRange(currentDestinationRow, 1, 3, 1).setValues([["VAC"], ["RDO"], ["SHIFT"]]);
        newScheduleSheet.getRange(currentDestinationRow, 2).setValue(employeeRecord[COLUMN_INDEX_NAME]);
        newScheduleSheet.getRange(currentDestinationRow, 2, 3, 1).merge().setVerticalAlignment("center");

        // 2. Checkboxes
        newScheduleSheet.getRange(currentDestinationRow, 3, 2, 7).insertCheckboxes();

        // 3. Set RDO Preferences
        const preferences = [employeeRecord[COLUMN_INDEX_PREFERENCE_ONE], employeeRecord[COLUMN_INDEX_PREFERENCE_TWO]];
        preferences.forEach(day => {
            const dayIndex = weekDayNames.indexOf(day);
            if (dayIndex !== -1) newScheduleSheet.getRange(currentDestinationRow + 1, 3 + dayIndex).setValue(true);
        });

        currentDestinationRow += 3;
    });

    attachStaffingSummary(newScheduleSheet, currentDestinationRow - 1);
    
    // RUN THE BULK ENGINE
    resolveEntireWeek(newScheduleSheet);
}

/**
 * Adds the Staffing Summary block (REQUIRED/ACTUAL/STATUS) to the bottom of the sheet.
 */
function attachStaffingSummary(currentScheduleSheet, lastEmployeeRowNumber) {
    // 1. Calculate how many total employees are on this specific sheet
    // Formula: (Last Row used for employees - Header offset) / 2 rows per employee
    const totalEmployeeCount = (lastEmployeeRowNumber - 5) / 3;

    const summaryHeaderRowIndex = lastEmployeeRowNumber + 2;
    const weekDayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    // 2. Clear the summary area to ensure a clean slate
    currentScheduleSheet.getRange(summaryHeaderRowIndex, 2, 5, 8).clearContent();

    // 3. Set the Row Headers
    currentScheduleSheet.getRange(summaryHeaderRowIndex, 2, 3, 1)
        .setValues([["REQUIRED"], ["ACTUAL"], ["STATUS"]])
        .setFontWeight("bold");

    // 4. Loop through each day to set formulas
    for (let dayIndex = 0; dayIndex < 7; dayIndex++) {
        const columnNumber = 3 + dayIndex;
        const columnLetter = String.fromCharCode(67 + dayIndex);
        const minimumRequired = getMinimumStaffRequiredForDay(weekDayNames[dayIndex]);

        // Set the 'Required' value from SETTINGS
        currentScheduleSheet.getRange(summaryHeaderRowIndex, columnNumber).setValue(minimumRequired);

        // Define the range where checkboxes live (Row 6 to the last employee row)
        const startRow = 6;
        const endRow = lastEmployeeRowNumber;

        /** * ACTUAL STAFF FORMULA:
         * We take the total count and subtract anyone who has a checkmark in 
         * EITHER their vacation row OR their preference row.
         */
        const actualStaffFormula = `=${totalEmployeeCount} - COUNTIF(${columnLetter}${startRow}:${columnLetter}${endRow}, TRUE)`;

        currentScheduleSheet.getRange(summaryHeaderRowIndex + 1, columnNumber)
            .setFormula(actualStaffFormula)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

        // Set the Visual Status OK/UNDER
        const statusFormula = `=IF(${columnLetter}${summaryHeaderRowIndex + 1} >= ${columnLetter}${summaryHeaderRowIndex}, "OK", "UNDER")`;
        currentScheduleSheet.getRange(summaryHeaderRowIndex + 2, columnNumber)
            .setFormula(statusFormula)
            .setHorizontalAlignment("center");
    }
}

/**
 * Final resolution logic: Assigns Shift Times to working staff and bumps juniors if over-staffed.
 */
function resolveEntireWeek(currentScheduleSheet) {
    // Calculate the number of employees on the sheet
    const allNames = currentScheduleSheet.getRange("B6:B").getValues();
    const employeeNames = allNames.filter(row => row[0] !== "" && !["REQUIRED", "ACTUAL", "STATUS"].includes(row[0])).map(r => r[0]);
    const numEmployees = employeeNames.length;

    if (numEmployees === 0) return;

    // 1. Read the entire spreadsheet: Grab C6:J[End] (Checkboxes, Shifts, and Hours)
    const gridRange = currentScheduleSheet.getRange(6, 3, numEmployees * 3, 8);
    let gridValues = gridRange.getValues();

    const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURATION_SHEET_NAME);
    const roster = config.getRange(2, 1, config.getLastRow() - 1, 8).getValues();
    const rosterMap = {};
    roster.forEach(r => rosterMap[r[0]] = { status: r[1], pref: r[2] });

    for (let e = 0; e < numEmployees; e++) {gridValues[(e * 3) + 2][7] = 0; }

    // 3. Bulk process the entire schedule matrix
    for (let dayIdx = 0; dayIdx < 7; dayIdx++) {
        const dayName = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][dayIdx];
        const maxOff = numEmployees - getMinimumStaffRequiredForDay(dayName);
        let offCount = 0;

        for (let empIdx = 0; empIdx < numEmployees; empIdx++) {
            const base = empIdx * 3;
            const name = employeeNames[empIdx][0];
            const isVac = gridValues[base][dayIdx];
            const isRdo = gridValues[base + 1][dayIdx];

            if (isVac === true) {
                offCount++;
                gridValues[base + 1][dayIdx] = false;
                gridValues[base + 2][dayIdx] = "VAC";
            } else if (isRdo === true && offCount < maxOff) {
                offCount++;
                gridValues[base + 2][dayIdx] = "OFF";
            } else {
                gridValues[base + 1][dayIdx] = false; // Bumped
                const person = rosterMap[name] || { status: "PT", pref: "Morning" };
                const shift = getShiftTimingFromSettings(person.pref, person.status);
                
                gridValues[base + 2][dayIdx] = shift.text;
                // Add hours to Column J (Index 7)
                gridValues[base + 2][7] = (Number(gridValues[base + 2][7]) || 0) + shift.hours;
            }
        }
    }

    // 4. Bulk write, slams the entire week's worth of data down at once instead of multiple trips
    gridRange.setValues(gridValues);
}
/**
 * Generates a standardized label for the weekly schedule tab (e.g., "Week_03_23_26").
 * It identifies the coming Monday to ensure consistent naming.
 */
function generateWeeklyDateLabel() {
    const todayDate = new Date();
    const dayOfWeekIndex = todayDate.getDay(); // 0 = Sunday, 1 = Monday, etc.

    // Calculate how many days to subtract to get back to the most recent Monday
    // If today is Sunday (0), we go back 6 days. Otherwise, we go back (day - 1) days.
    const daysToSubtractToReachMonday = (dayOfWeekIndex === 0) ? 6 : dayOfWeekIndex - 1;

    const mondayDate = new Date(todayDate);
    mondayDate.setDate(todayDate.getDate() - daysToSubtractToReachMonday);

    // Format components with leading zeros
    const calendarMonth = (mondayDate.getMonth() + 1).toString().padStart(2, '0');
    const calendarDay = mondayDate.getDate().toString().padStart(2, '0');
    const shortYear = mondayDate.getFullYear().toString().slice(-2);

    const finalWeeklyLabel = "Week_" + calendarMonth + "_" + calendarDay + "_" + shortYear;

    Logger.log("Generated Sheet Label: " + finalWeeklyLabel);
    return finalWeeklyLabel;
}