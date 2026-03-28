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
 * Version 0.0.11
 * /

/**
 * Creates a new Weekly Schedule based on the Template.
 */
function generateWeeklySchedule() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configurationSheet = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);
    const templateSheet      = activeSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

    const weeklySheetLabel = generateWeeklyDateLabel();
    let newScheduleSheet   = activeSpreadsheet.getSheetByName(weeklySheetLabel);
    
    if (newScheduleSheet) {
        activeSpreadsheet.deleteSheet(newScheduleSheet);
    }
    
    newScheduleSheet = templateSheet.copyTo(activeSpreadsheet).setName(weeklySheetLabel);
    newScheduleSheet.showSheet();

    const rosterData = configurationSheet.getRange(2, 1, configurationSheet.getLastRow() - 1, 8).getValues();
    rosterData.sort((firstEmployee, secondEmployee) => secondEmployee[COLUMN_INDEX_SENIORITY_RANK] - firstEmployee[COLUMN_INDEX_SENIORITY_RANK]);

    let currentDestinationRow = 6;
    const weekDayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

    rosterData.forEach(employeeRecord => {
        const employeeName = employeeRecord[COLUMN_INDEX_NAME];

        // Row labels
        newScheduleSheet.getRange(currentDestinationRow, 1).setValue("VAC");
        newScheduleSheet.getRange(currentDestinationRow + 1, 1).setValue("RDO");
        newScheduleSheet.getRange(currentDestinationRow + 2, 1).setValue("SHIFT");
        
        // Optional styling so you know its just a label not actual data
        newScheduleSheet.getRange(currentDestinationRow, 1, 3, 1)
            .setFontSize(8)
            .setFontColor("#666666")
            .setHorizontalAlignment("right")
            .setVerticalAlignment("middle"); 
        
        // Merge the name blocks in column B across 3 rows
        newScheduleSheet.getRange(currentDestinationRow, 2).setValue(employeeName);
        newScheduleSheet.getRange(currentDestinationRow, 2, 3, 1).merge().setVerticalAlignment("center");

        // Insert checkboxes in the first two (of three) rows only
        newScheduleSheet.getRange(currentDestinationRow, 3, 2, 7).insertCheckboxes();

        const preferences = [employeeRecord[COLUMN_INDEX_PREFERENCE_ONE], employeeRecord[COLUMN_INDEX_PREFERENCE_TWO]];
        preferences.forEach(preferenceDay => {
            if (preferenceDay) {
                const dayIndex = weekDayNames.indexOf(preferenceDay);
                if (dayIndex !== -1) {
                    newScheduleSheet.getRange(currentDestinationRow, 3 + dayIndex).setValue(true);
                }
            }
        });

        const totalHoursCell = newScheduleSheet.getRange(currentDestinationRow + 2, 10);
        totalHoursCell.setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center");

        currentDestinationRow += 3;
    });

    attachStaffingSummary(newScheduleSheet, currentDestinationRow - 1);
    
    for (let columnCounter = 3; columnCounter <= 9; columnCounter++) {
        resolveSeniorityConflicts(newScheduleSheet, weekDayNames[columnCounter - 3], columnCounter);
    }
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
function resolveSeniorityConflicts(currentScheduleSheet, dayName, columnNumber) {
    const minimumRequiredStaff = getMinimumStaffRequiredForDay(dayName);
    const lastRowInSheet       = currentScheduleSheet.getLastRow();
    
    // We calculate count by looking at merged cells in Column B
    const employeeNamesList     = currentScheduleSheet.getRange("B6:B" + lastRowInSheet).getValues().filter(String);
    const totalEmployeesOnSheet = employeeNamesList.length;
    const maximumAllowedOff     = totalEmployeesOnSheet - minimumRequiredStaff;

    let currentOffCount = 0;
    let employeesMarkedAsWorking = [];

    // Step A: Determine Who is Working
    for (let rowIndex = 6; rowIndex < 6 + (totalEmployeesOnSheet * 3); rowIndex += 3) {
        const isOnVacation      = currentScheduleSheet.getRange(rowIndex, columnNumber).getValue();
        const requestedDayOff   = currentScheduleSheet.getRange(rowIndex + 1, columnNumber).getValue();
        const currentEmployeeName = currentScheduleSheet.getRange(rowIndex, 2).getValue();

        if (isOnVacation === true) {
            currentOffCount++;
            currentScheduleSheet.getRange(rowIndex + 1, columnNumber).setValue(false);
            currentScheduleSheet.getRange(rowIndex + 2, columnNumber).setValue("VAC").setFontColor("#cc0000");
        } else if (requestedDayOff === true && currentOffCount < maximumAllowedOff) {
            currentOffCount++;
            currentScheduleSheet.getRange(rowIndex + 2, columnNumber).setValue("OFF").setFontColor("#666666"); // Clear time if they are off
        } else {
            // This employee is Working
            currentScheduleSheet.getRange(rowIndex + 1, columnNumber).setValue(false);
            employeesMarkedAsWorking.push({ row: rowIndex, name: currentEmployeeName });
        }
    }

    // Step B: Assign Timing based on SETTINGS
    const configurationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURATION_SHEET_NAME);
    const fullRosterData     = configurationSheet.getRange(2, 1, configurationSheet.getLastRow(), 8).getValues();

    employeesMarkedAsWorking.forEach(worker => {
        const employeeConfigRecord = fullRosterData.find(record => record[COLUMN_INDEX_NAME] === worker.name);
        const employmentStatus     = employeeConfigRecord ? employeeConfigRecord[COLUMN_INDEX_EMPLOYMENT_STATUS] : "PT";
        const shiftPreference      = employeeConfigRecord ? employeeConfigRecord[COLUMN_INDEX_SHIFT_PREFERENCE] : "Morning";

        const formattedShiftTime = getShiftTimingFromSettings(shiftPreference, employmentStatus);

        currentScheduleSheet.getRange(worker.row + 2, columnNumber)
            .setValue(formattedShiftTime.text)
            .setFontSize(8)
            .setHorizontalAlignment("center")
            .setFontColor("#000000");
        
        const totalHoursCell = currentScheduleSheet.getRange(worker.row + 2, 10);
        const existingHours  = totalHoursCell.getValue() || 0;
        totalHoursCell.setValue(existingHours + shiftInfo.hours);
    });
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
    const calendarDay   = mondayDate.getDate().toString().padStart(2, '0');
    const shortYear     = mondayDate.getFullYear().toString().slice(-2);

    const finalWeeklyLabel = "Week_" + calendarMonth + "_" + calendarDay + "_" + shortYear;
    
    Logger.log("Generated Sheet Label: " + finalWeeklyLabel);
    return finalWeeklyLabel;
}