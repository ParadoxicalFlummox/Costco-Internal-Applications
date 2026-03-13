/*
   _____          _                       _                               _   _       _   _  __ _           
  / ____|        | |                /\   | |                             | \ | |     | | (_)/ _(_)          
 | |     ___  ___| |_ ___ ___      /  \  | |__  ___  ___ _ __   ___ ___  |  \| | ___ | |_ _| |_ _  ___ _ __ 
 | |    / _ \/ __| __/ __/ _ \    / /\ \ | '_ \/ __|/ _ \ '_ \ / __/ _ \ | . ` |/ _ \| __| |  _| |/ _ \ '__|
 | |___| (_) \__ \ || (_| (_) |  / ____ \| |_) \__ \  __/ | | | (_|  __/ | |\  | (_) | |_| | | | |  __/ |   
  \_____\___/|___/\__\___\___/  /_/    \_\_.__/|___/\___|_| |_|\___\___| |_| \_|\___/ \__|_|_| |_|\___|_|   
  
VERSION: 0.2.1
LICNSE: MIT (Generic Logic Engine)
PURPOSE: A time windowed notification system to
         map internal spreadsheet rows to descriptive
         objects and send department managers emails
*/

/* SCRIPT CONFIGURATION (ONLY TO BE MAPPED ON THE CLOCK) */

const CONFIG = {
   // Logic for the sheet name calculations

   get activeSheetName() {return calculateWeekOf_();},
   
   // The time window configuration, see README for more information
   windowMinutes: 15,

   // MAP: Column letters to indices (0=A, 1=B, ...)
   // Update these to match the spreadsheet structure
   COLUMNS: {
      NAME: 0,
      EMPLOYEE_ID: 1,
      IS_CALLOUT: 3,
      IS_FMLA: 5,
      IS_NOSHOW: 6,
      DEPT: 7,
      TIME: 8,
      COMMENT: 13,
      DATE: 14,
   },

   MAILING_LIST: {
      "Dept": ["Dept.email.com"],
      "Dept2": ["Dept2.email.com"]
      // Add additional mappings to match your warehouse's email structure
   },

   FALLBACK_EMAIL: ["email1.email.com", "email2.email.com"]
};

/* MAIN EXECUTION ENGINE */

function scriptEngine(){
   try{
      const callLog = SpreadsheetApp.getActiveSpreadsheet();
      const currentLog = callLog.getSheetByName(CONFIG.activeSheetName);

      if(!currentLog){
         console.warn(`Sheet "${CONFIG.activeSheetName}" not found. Skipping.`);
         return;
      }

      // Define the time window to search for logs
      const window = getPreviousWindow_(CONFIG.windowMinutes);
      const timeZone = Session.getScriptTimeZone();
      const windowStartMinute = window.start.getTime();
      const windowEndMinute = window.end.getTime();

      // Data extraction
      const firstRow = 3;
      const lastRow = sheet.getLastRow();
      if (lastRow < firstRow) return;

      const values = sheet.getRange(firstRow, 1, (lastRow - firstRow + 1), 15).getValues();

      // Use the mapped columns from the config to make a map of all data in the spreadsheet
      const potentialAbsences = values.map((row, index) => {
         const entry = CONFIG.COLUMNS;
         return{
            rowNumber: firstRow + index,
            employeeName: String(row[entry.NAME] || 'Unknown'),
            employeeID: String(row[entry.EMPLOYEE_ID] || 'Unknown'),
            isAbsence: toBool(row[entry.IS_CALLOUT]) || toBool(row[entry.IS_FMLA]) || toBool(row[entry.IS_NOSHOW]),
            reason: toBool(row[entry.IS_CALLOUT]) ? 'Call-Out' : toBool(row[entry.IS_FMLA]) ? 'FMLA' : 'No-Show',
            department: String(row[entry.DEPT] || 'Unassigned').trim(),
            timeRaw: row[entry.TIME],
            comment: String(row[entry.COMMENT] || '-'),
            dateRaw: row[entry.DATE]
         };
      });

      // Filtering logic, only selects the log entries that actually require a notification

      const absencesInTimeWindow = potentialAbsences.filter(record => {
         if (!record.isAbsence) return false;

         // Time check
         const callMinute = parseTimeToMillisInWindow_(record.timeRaw, window);
         if (!callMinute || callMinute <= windowStartMinute || callMinute > windowEndMinute) return false;
         record.calledAt = new Date(callMinute);

         // Date check
         const callDate = coerceToLocalDateIfPresent_(record.dateRaw, timeZone);
         if(!callDate) return false;
      })
   }
}