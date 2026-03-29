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
 * Version 0.0.15
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
 * Handles live spreadsheet edits (now with bulk editing logic).
 */
function onEdit(event) {
    const editedRange = event.range;
    const editedSheet = editedRange.getSheet();
    const sheetName   = editedSheet.getName();
    const column      = editedRange.getColumn();
    const row         = editedRange.getRow();

    // 1. IGNORE edits on the CONFIG or SETTINGS sheets
    if (sheetName === CONFIGURATION_SHEET_NAME || sheetName === SETTINGS_SHEET_NAME) {
        return; 
    }

    // 2. TARGET: Weekly Schedule Tabs ("Week_")
    // Only trigger if the edit is in the Checkbox Range (Cols C-I, Row 6+)
    if (sheetName.startsWith("Week_") && column >= 3 && column <= 9 && row >= 6) {
        
        // K.I.S.S. Check: Is this a VAC or RDO row? 
        // (Row - 6) % 3 == 0 is VAC, % 3 == 1 is RDO.
        const relativeRow = (row - 6) % 3;
        
        if (relativeRow === 0 || relativeRow === 1) {
            // OPTIONAL: Provide a UI hint that the script is "thinking"
            SpreadsheetApp.getActiveSpreadsheet().toast("Re-calculating seniority and hours...", "System Sync", 2);
            
            // Trigger the Bulk Engine to fix the whole week instantly
            resolveEntireWeek(editedSheet);
        }
    }
}