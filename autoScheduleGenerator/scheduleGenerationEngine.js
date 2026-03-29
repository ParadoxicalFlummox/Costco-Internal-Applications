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
 * Version 0.0.12
 * /

/**
 * Creates a new Weekly Schedule based on the Template.
 */
function generateWeeklySchedule() {
    const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const configurationSheet = activeSpreadsheet.getSheetByName(CONFIGURATION_SHEET_NAME);
    const templateSheet = activeSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

    const weeklySheetLabel = generateWeeklyDateLabel();
    let newScheduleSheet = activeSpreadsheet.getSheetByName(weeklySheetLabel);

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

        // Write row labels
        const labelRange = newScheduleSheet.getRange(currentDestinationRow, 1, 3, 1);
        labelRange.setValues([["VAC"], ["RDO"], ["SHIFT"]]);

        // Optional styling so you know its just a label not actual data
        labelRange.setFontSize(8)
            .setFontColor("#666666")
            .setHorizontalAlignment("right")
            .setVerticalAlignment("middle")
            .setFontStyle("italic");

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
    const minStaff = getMinimumStaffRequiredForDay(dayName);
    const lastRow = currentScheduleSheet.getLastRow();

    // Get all names and calculate total based on actual data present
    const names = currentScheduleSheet.getRange("B6:B" + lastRow).getValues().filter(String);
    const totalEmployees = names.length;
    const maxAllowedOff = totalEmployees - minStaff;

    let offCount = 0;

    // Reset Column J for this day's calculation to prevent stacking errors
    // (Only do this on Monday/Column 3 to start fresh)
    if (columnNumber === 3) {
        currentScheduleSheet.getRange(6, 10, totalEmployees * 3, 1).setValue(0);
    }

    // Step A: Determine Who is Working (Increment by 3)
    for (let i = 0; i < totalEmployees; i++) {
        const row = 6 + (i * 3);
        const isVacation = currentScheduleSheet.getRange(row, columnNumber).getValue();
        const isRDO = currentScheduleSheet.getRange(row + 1, columnNumber).getValue();
        const empName = currentScheduleSheet.getRange(row, 2).getValue();
        const shiftCell = currentScheduleSheet.getRange(row + 2, columnNumber);

        // 1. Check Vacation First (Non-negotiable)
        if (isVacation === true) {
            offCount++;
            currentScheduleSheet.getRange(row + 1, columnNumber).setValue(false); // Clear RDO
            shiftCell.setValue("VAC").setFontColor("#cc0000").setHorizontalAlignment("center");
        }
        // 2. Check RDO (Only if we haven't hit the limit)
        else if (isRDO === true && offCount < maxAllowedOff) {
            offCount++;
            shiftCell.setValue("OFF").setFontColor("#666666").setHorizontalAlignment("center");
        }
        // 3. Otherwise: Assign Shift
        else {
            currentScheduleSheet.getRange(row + 1, columnNumber).setValue(false); // Clear RDO if bumped

            // Get shift info from Roster/Settings
            const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIGURATION_SHEET_NAME);
            const roster = config.getRange(2, 1, config.getLastRow(), 8).getValues();
            const person = roster.find(r => r[COLUMN_INDEX_NAME] === empName);

            const status = person ? person[COLUMN_INDEX_EMPLOYMENT_STATUS] : "PT";
            const pref = person ? person[COLUMN_INDEX_SHIFT_PREFERENCE] : "Morning";

            const shiftInfo = getShiftTimingFromSettings(pref, status);

            // Stamp Shift Time
            shiftCell.setValue(shiftInfo.text).setFontColor("#000000").setFontSize(8).setHorizontalAlignment("center");

            // Update Total Hours in Column J
            const totalCell = currentScheduleSheet.getRange(row + 2, 10);
            totalCell.setValue((totalCell.getValue() || 0) + shiftInfo.hours);
        }
    }
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