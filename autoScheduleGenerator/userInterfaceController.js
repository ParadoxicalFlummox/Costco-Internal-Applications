/**
 *       _____             _                                _            _____        _                _         _             
 *      / ____|           | |                  /\          | |          / ____|      | |              | |       | |            
 *     | |      ___   ___ | |_  ___  ___      /  \   _   _ | |_  ___   | (___    ___ | |__    ___   __| | _   _ | |  ___  _ __ 
 *     | |     / _ \ / __|| __|/ __|/ _ \    / /\ \ | | | || __|/ _ \   \___ \  / __|| '_ \  / _ \ / _` || | | || | / _ \| '__|
 *     | |____| (_) |\__ \| |_| (__| (_) |  / ____ \| |_| || |_| (_) |  ____) || (__ | | | ||  __/| (_| || |_| || ||  __/| |   
 *      \_____|\___/ |___/ \__|\___|\___/  /_/    \_\\__,_| \__|\___/  |_____/  \___||_| |_| \___| \__,_| \__,_||_| \___||_|   
 *      _    _  _____    _____               _                _  _                                                             
 *     | |  | ||_   _|  / ____|             | |              | || |                                                            
 *     | |  | |  | |   | |      ___   _ __  | |_  _ __  ___  | || |  ___  _ __                                                 
 *     | |  | |  | |   | |     / _ \ | '_ \ | __|| '__|/ _ \ | || | / _ \| '__|                                                
 *     | |__| | _| |_  | |____| (_) || | | || |_ | |  | (_) || || ||  __/| |                                                   
 *      \____/ |_____|  \_____|\___/ |_| |_| \__||_|   \___/ |_||_| \___||_|                                                   
 *                                                                                                                             
 *                                                                                                                             
 * Built by: Adam Roy
 * Version 0.0.11
 */


/**
 * Creates the custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const userInterface = SpreadsheetApp.getUi();
  userInterface.createMenu('Schedule Admin')
      .addItem('1. Synchronize Roster', 'synchronizeEmployeeRoster')
      .addItem('2. Refresh Seniority Ranks', 'refreshAllSeniorityRanks')
      .addSeparator()
      .addItem('GENERATE WEEKLY DRAFT', 'generateWeeklySchedule')
      .addSeparator()
      .addSubMenu(userInterface.createMenu('Danger Zone')
          .addItem('Reset Department Configuration', 'resetDepartmentConfiguration'))
      .addToUi();
}

/**
 * Handles live spreadsheet edits.
 */
function onEdit(event) {
    const editedRange = event.range;
    const editedSheet = editedRange.getSheet();
    const sheetName   = editedSheet.getName();
    const columnNumber = editedRange.getColumn();
    const rowNumber    = editedRange.getRow();

    // Handle Seniority and Shift updates on the Configuration tab
    if (sheetName === CONFIGURATION_SHEET_NAME && rowNumber > 1) {
        const isStatusColumn = (columnNumber === (COLUMN_INDEX_EMPLOYMENT_STATUS + 1));
        const isShiftColumn  = (columnNumber === (COLUMN_INDEX_SHIFT_PREFERENCE + 1));
        
        if (isStatusColumn || isShiftColumn) {
            calculateEmployeeSeniority(rowNumber);
        }
        return;
    }

    // Handle Live Conflict Resolution on Weekly Schedule tabs
    if (sheetName.startsWith("Week_") && columnNumber >= 3 && columnNumber <= 9 && rowNumber >= 6) {
        const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
        const selectedDayName = dayNames[columnNumber - 3];
        resolveSeniorityConflicts(editedSheet, selectedDayName, columnNumber);
    }
}