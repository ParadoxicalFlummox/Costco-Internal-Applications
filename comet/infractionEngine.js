/**
 * infractionEngine.js — Main orchestrator for the infraction detection pipeline.
 * VERSION: 0.2.3
 *
 * This file contains a single public function: scanAndIssueCNs(). Its only
 * job is to call the other files in the correct order and pass data between
 * them. It contains no parsing logic, no email formatting, no log writes, and
 * no threshold math. Each of those concerns lives in its own file:
 *
 *   config.js            — All configurable values (thresholds, codes, layout)
 *   calendarParser.js    — Reads employee tabs and emits CalendarEvent objects
 *   infractionDetector.js — Rolling-window CN detection; emits CNProposal objects
 *   cnLog.js             — CN_Log sheet access, deduplication index, expiry job
 *   notifier.js          — Email construction and sending
 *   ui.js                — onOpen menu and menu handler wrappers
 *
 * HOW TO SET UP THE DAILY TRIGGER:
 *   1. In the Apps Script editor, go to Triggers → Add Trigger.
 *   2. Function: sendCNsDaily
 *   3. Event source: Time-driven → Day timer
 *   4. Choose a time (e.g. 6:00 AM) so the scan runs before the morning shift.
 *
 * HOW TO DRY-RUN (no emails, no log writes):
 *   Run dryRunCNs() from the menu or directly from the Apps Script editor.
 *   Output appears in View → Logs.
 */


// ---------------------------------------------------------------------------
// Public Entry Points
// ---------------------------------------------------------------------------

/**
 * Scans all employee tabs in dry-run mode.
 * No emails are sent and nothing is written to the CN_Log.
 * All proposals are logged to the Apps Script console for review.
 *
 * Callable from: "Infraction Notifier" → "Dry Run (Log Only)"
 */
function dryRunCNs() {
  scanAndIssueCNs({ dryRun: true });
}

/**
 * Scans all employee tabs and sends CN notifications for new infractions.
 * Intended to be called by the daily time-driven trigger.
 *
 * Uses DRY_RUN from config.js as the default — flip that constant to false
 * when you are ready to go live.
 *
 * Callable from: "Infraction Notifier" → "Send CNs (Live)" or the daily trigger.
 */
function sendCNsDaily() {
  scanAndIssueCNs({ dryRun: !!DRY_RUN }); // DRY_RUN defined in config.js
}


// ---------------------------------------------------------------------------
// Main Orchestrator
// ---------------------------------------------------------------------------

/**
 * Scans every employee tab in the attendance controller workbook for infraction
 * patterns and issues Counseling Notice notifications where thresholds are met.
 *
 * Pipeline:
 *   1. Open the CN_Log sheet and build a deduplication index.
 *   2. Iterate every sheet tab in the workbook.
 *   3. For each tab that matches the employee tab naming pattern, parse the
 *      calendar grid into CalendarEvent objects.
 *   4. Filter to infraction-only events and run them through the rolling-window
 *      detector to produce CNProposal objects.
 *   5. Deduplicate proposals against the CN_Log index (skip any whose CN_Key
 *      already exists with the same EventsHash).
 *   6. Send notifications for new proposals and append them to the CN_Log.
 *
 * Early exits (no action):
 *   - A sheet tab does not match the employee tab naming pattern.
 *   - An employee has no infraction events in their calendar.
 *   - All proposals for an employee were already logged with the same evidence.
 *
 * Errors on individual sheets are caught and logged so a single malformed tab
 * cannot abort the entire scan.
 *
 * @param {{ dryRun?: boolean, sendEmail?: boolean }} options
 * @returns {{ proposals: number, issued: number }}
 */
function scanAndIssueCNs(options) {
  const opts      = options || {};
  const dryRun    = opts.dryRun    != null ? !!opts.dryRun    : !!DRY_RUN; // config.js
  const sendEmail = opts.sendEmail != null ? !!opts.sendEmail : !dryRun;   // default: send when not dry run
  const timeZone = Session.getScriptTimeZone();

  console.log(`infractionEngine: Starting scan — dryRun=${dryRun}, daysBack=${DAYS_BACK}`);

  const workbook = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSpreadsheetId = workbook.getId(); // captured once; attached to every proposal
  const sheets = workbook.getSheets();

  // Step 1: Open the log and build the deduplication index
  const logSheet = getOrCreateLogSheet_();   // cnLog.js
  const logIndex = buildLogIndex_(logSheet); // cnLog.js
  console.log(`infractionEngine: CN_Log loaded — ${logIndex.size} existing entries.`);

  // Build active employee ID set from the Employees sheet so archived employees
  // are excluded from the scan without needing to hide their tabs.
  const activeEmployeeIds = new Set(
    getActiveEmployees_().map(e => e.id) // ukgImport.js
  );
  console.log(`infractionEngine: ${activeEmployeeIds.size} active employee(s) in roster.`);

  const newProposals = [];

  // Step 2–4: Parse each employee tab and detect infractions
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (!isEmployeeTab_(sheetName)) return; // calendarParser.js — skip non-employee tabs

    try {
      const ctx = readEmployeeContext_(sheet);   // calendarParser.js

      // Skip tabs whose employee is not in the active roster.
      // This respects the Archived status set via the COMET Admin panel without
      // requiring the tab itself to be hidden.
      if (activeEmployeeIds.size > 0 && ctx.employeeId && !activeEmployeeIds.has(ctx.employeeId)) {
        console.log(`infractionEngine: Skipping archived employee — "${sheetName}"`);
        return;
      }

      const year = parseYearFromTitle_(ctx.yearTitle) || new Date().getFullYear();

      const allEvents = parseCalendarEvents_(sheet, year, timeZone, ctx); // calendarParser.js
      const infractionEvents = allEvents
        .filter(e => e.isInfraction && !e.isIgnored)
        .sort((a, b) => a.date.getTime() - b.date.getTime());

      if (infractionEvents.length === 0) return;

      const proposals = buildAllCNProposals_(infractionEvents, timeZone, ctx); // infractionDetector.js
      if (proposals.length === 0) return;

      // Attach the source sheet name, spreadsheet ID, and sheet GID to each
      // proposal so cnLog.js can build a direct HYPERLINK to the employee tab.
      const sourceSheetGid = sheet.getSheetId();
      proposals.forEach(p => {
        p.sheetName = sheetName;
        p.sourceSpreadsheetId = sourceSpreadsheetId;
        p.sourceSheetGid = sourceSheetGid;
      });

      console.log(
        `infractionEngine: "${sheetName}" — ` +
        `${infractionEvents.length} infraction event(s), ` +
        `${proposals.length} proposal(s) generated.`
      );

      newProposals.push(...proposals);
    } catch (error) {
      console.error(`infractionEngine: Error parsing sheet "${sheetName}" — ${error.message}`);
    }
  });

  if (newProposals.length === 0) {
    console.log('infractionEngine: No CN proposals found across all employee tabs.');
    return { proposals: 0, issued: 0 };
  }

  // Step 5: Deduplicate against the log index
  const toSend = newProposals.filter(proposal => {
    const existing = logIndex.get(proposal.cnKey);
    if (existing && existing.eventsHash === proposal.eventsHash) {
      console.log(`infractionEngine: Skipping duplicate — ${proposal.cnKey}`);
      return false;
    }
    return true;
  });

  console.log(
    `infractionEngine: ${newProposals.length} proposal(s) generated, ` +
    `${newProposals.length - toSend.length} duplicate(s) skipped, ` +
    `${toSend.length} new to send.`
  );

  if (toSend.length === 0) return { proposals: newProposals.length, issued: 0 };

  // Step 6: Send notifications and/or write to the log
  // dryRun  → log to console only, no sheet writes, no email
  // !dryRun → always write to sheet; sendEmail controls whether emails go out
  if (dryRun) {
    sendCNNotifications_(toSend, true, timeZone); // notifier.js — dry-run: console only
    console.log('infractionEngine: Dry run complete. No sheets written.');
    return { proposals: newProposals.length, issued: 0 };
  }

  sendCNNotifications_(toSend, !sendEmail, timeZone); // notifier.js — suppress email when sendEmail=false

  const issuedAt = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
  const issuedBy = Session.getActiveUser() ? Session.getActiveUser().getEmail() : '';

  toSend.forEach(proposal => {
    // Write to CN_Log (dedup source of truth)
    appendLogRow_(logSheet, { // cnLog.js
      CN_Key: proposal.cnKey,
      EmployeeID: proposal.employeeId,
      EmployeeName: proposal.employeeName,
      Department: proposal.department,
      WindowStart: formatDateYmd_(proposal.windowStart, timeZone), // infractionDetector.js
      WindowEnd: formatDateYmd_(proposal.windowEnd, timeZone),
      Count: String(proposal.count),
      EventsHash: proposal.eventsHash,
      IssuedAt: issuedAt,
      IssuedBy: issuedBy,
      DryRun: 'FALSE',
      SheetName: proposal.sheetName || '',
      Status: 'Active',
      ExpiredAt: '',
      Rule: proposal.rule || 'GLOBAL',
      SourceSpreadsheetId: proposal.sourceSpreadsheetId || '',
      SourceSheetGid: proposal.sourceSheetGid != null ? String(proposal.sourceSheetGid) : '',
    });

    // Write to Active CNs (manager-facing view with hyperlink)
    appendActiveCNRow_(proposal, issuedAt, timeZone); // cnLog.js

    // Update the in-memory index so subsequent proposals in the same run
    // that share a key are not re-logged
    logIndex.set(proposal.cnKey, { eventsHash: proposal.eventsHash });
  });

  console.log(`infractionEngine: ${toSend.length} CN(s) logged to CN_Log. Email sent: ${sendEmail}`);
  console.log('infractionEngine: Scan complete.');
  return { proposals: newProposals.length, issued: toSend.length };
}
