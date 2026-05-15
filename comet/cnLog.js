/**
 * cnLog.js — CN store management, deduplication indexes, and CN expiry processing.
 * VERSION: 0.4.5
 *
 * STORAGE MODEL
 * -------------
 * Active CNs and Expired CNs are stored in two Google Sheets tabs, each using
 * a per-employee row layout:
 *
 *   Row 1: Title (styled header)
 *   Row 2: Column headers  (Employee ID | Employee Name | Department | CNs (JSON))
 *   Row 3+: One row per employee, where the last column is a JSON array of that
 *           employee's CN records.
 *
 * When a CN is issued:
 *   → Find the employee's row in Active CNs (by Employee ID).
 *   → If absent, create a new row.
 *   → Parse the JSON array, push the new CN record, write back.
 *
 * When a CN expires:
 *   → Find the employee's row in Active CNs.
 *   → Remove the expired record from their array and write back.
 *   → Find (or create) the employee's row in Expired CNs.
 *   → Push the record (with expiredAt set) into that array and write back.
 *
 * CN record shape (stored in the JSON arrays):
 * {
 *   cnKey:               string   — Deduplication key
 *   windowStart:         string   — YYYY-MM-DD
 *   windowEnd:           string   — YYYY-MM-DD
 *   count:               number
 *   eventsHash:          string
 *   issuedAt:            string   — "yyyy-MM-dd HH:mm:ss"
 *   issuedBy:            string
 *   sheetName:           string
 *   rule:                string
 *   sourceSpreadsheetId: string
 *   sourceSheetGid:      number|string
 *   consumedEvents:      string[] — ["YYYY-MM-DD|CODE", ...]
 *   expiredAt?:          string   — set only in Expired CNs
 * }
 */


// ---------------------------------------------------------------------------
// Workbook Resolution
// ---------------------------------------------------------------------------

/**
 * Returns the COMET workbook that hosts both CN store sheets.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function resolveLogWorkbook_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}


// ---------------------------------------------------------------------------
// Sheet Access
// ---------------------------------------------------------------------------

/**
 * Returns the Active CNs sheet, creating it if it does not yet exist.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateActiveCNsSheet_() {
  const workbook = resolveLogWorkbook_();
  let sheet = workbook.getSheetByName(ACTIVE_CNS_SHEET_NAME); // config.js
  if (!sheet) {
    sheet = workbook.insertSheet(ACTIVE_CNS_SHEET_NAME);
    initializeCNStoreSheet_(sheet, 'Active Counseling Notices', '#E31837', '#FFFFFF');
    console.log(`cnLog: Created Active CNs sheet in "${workbook.getName()}".`);
  }
  return sheet;
}

/**
 * Returns the (Expired CNs) sheet, creating it (hidden) if it does not exist.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateExpiredCNsSheet_() {
  const workbook = resolveLogWorkbook_();
  let sheet = workbook.getSheetByName(EXPIRED_CNS_SHEET_NAME); // config.js
  if (!sheet) {
    sheet = workbook.insertSheet(EXPIRED_CNS_SHEET_NAME);
    initializeCNStoreSheet_(sheet, 'Expired Counseling Notices (Archive)', '#B7B7B7', '#FFFFFF');
    sheet.hideSheet();
    console.log(`cnLog: Created (Expired CNs) sheet (hidden) in "${workbook.getName()}".`);
  }
  return sheet;
}

/**
 * Writes the title row, header row, and column widths to a freshly created CN
 * store sheet. Called once during sheet creation only.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} title     — Human-readable sheet title (row 1)
 * @param {string} bgColor   — Header row background color
 * @param {string} textColor — Header row text color
 */
function initializeCNStoreSheet_(sheet, title, bgColor, textColor) {
  // Row 1: title
  const titleRange = sheet.getRange(1, 1, 1, CN_STORE_HEADERS.length); // config.js
  titleRange.merge().setValue(title);
  titleRange
    .setBackground(bgColor)
    .setFontColor(textColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontSize(12);

  // Row 2: column headers
  const headerRange = sheet.getRange(2, 1, 1, CN_STORE_HEADERS.length);
  headerRange.setValues([CN_STORE_HEADERS]);
  applyHeaderStyle_(headerRange, bgColor, textColor);
  sheet.setFrozenRows(2);

  // Column widths
  sheet.setColumnWidth(CN_STORE_COL.employeeId,   100);  // config.js
  sheet.setColumnWidth(CN_STORE_COL.employeeName, 200);
  sheet.setColumnWidth(CN_STORE_COL.department,   130);
  sheet.setColumnWidth(CN_STORE_COL.cnsJson,      600);
}


// ---------------------------------------------------------------------------
// Per-Employee Row Read / Write
// ---------------------------------------------------------------------------

/**
 * Finds the 1-based sheet row for the given employeeId (searches column A
 * starting at row 3). Returns -1 if the employee does not have a row yet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} employeeId
 * @returns {number}
 */
function findEmployeeRow_(sheet, employeeId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return -1;
  const ids = sheet.getRange(3, CN_STORE_COL.employeeId, lastRow - 2, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === String(employeeId)) return i + 3;
  }
  return -1;
}

/**
 * Reads the CN JSON array for a single employee from the given sheet.
 * Returns an empty array if the employee has no row or the cell is empty.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} employeeId
 * @returns {Object[]}
 */
function readEmployeeCNs_(sheet, employeeId) {
  const row = findEmployeeRow_(sheet, employeeId);
  if (row === -1) return [];
  const raw = sheet.getRange(row, CN_STORE_COL.cnsJson).getValue();
  if (!raw) return [];
  try {
    const parsed = JSON.parse(String(raw));
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error(`cnLog: Could not parse CNs JSON for employee ${employeeId} in "${sheet.getName()}": ${e.message}`);
    return [];
  }
}

/**
 * Writes the CN JSON array for a single employee.
 * If the employee has no row yet, a new row is appended.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} employeeId
 * @param {string} employeeName
 * @param {string} department
 * @param {Object[]} cnArray
 */
function writeEmployeeCNs_(sheet, employeeId, employeeName, department, cnArray) {
  const row = findEmployeeRow_(sheet, employeeId);
  const jsonString = JSON.stringify(cnArray);
  if (row === -1) {
    // appendRow is guaranteed atomic: GAS always places it after the last
    // populated row regardless of any pending (unflushed) writes in the same
    // execution. Using getLastRow() + setValues() is unreliable here because
    // GAS may return a stale lastRow value before earlier setValues calls are
    // committed, causing two back-to-back appends to land on the same row.
    sheet.appendRow([employeeId, employeeName, department, jsonString]);
    console.log(`cnLog: writeEmployeeCNs_ — new row appended for employee ${employeeId} (${employeeName}) in "${sheet.getName()}".`);
  } else {
    sheet.getRange(row, CN_STORE_COL.cnsJson).setValue(jsonString);
    console.log(`cnLog: writeEmployeeCNs_ — updated row ${row} for employee ${employeeId} (${employeeName}) in "${sheet.getName()}".`);
  }
}


// ---------------------------------------------------------------------------
// Full-Store Readers (used for index building)
// ---------------------------------------------------------------------------

/**
 * Reads every employee row in a CN store sheet and returns a flat array of all
 * CN records across all employees, each augmented with the employee identity
 * fields from the row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object[]}
 */
function readAllCNsFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const rows = sheet.getRange(3, 1, lastRow - 2, CN_STORE_HEADERS.length).getValues();
  const results = [];

  for (const row of rows) {
    const employeeId   = String(row[CN_STORE_COL.employeeId   - 1] || '').trim();
    const employeeName = String(row[CN_STORE_COL.employeeName - 1] || '').trim();
    const department   = String(row[CN_STORE_COL.department   - 1] || '').trim();
    const rawJson      = String(row[CN_STORE_COL.cnsJson      - 1] || '').trim();

    if (!employeeId || !rawJson) continue;

    let cnArray;
    try {
      cnArray = JSON.parse(rawJson);
    } catch (e) {
      continue;
    }
    if (!Array.isArray(cnArray)) continue;

    for (const cn of cnArray) {
      results.push({ ...cn, employeeId, employeeName, department });
    }
  }

  return results;
}


// ---------------------------------------------------------------------------
// Index Builders
// ---------------------------------------------------------------------------

/**
 * Builds an in-memory Map for deduplication lookups.
 * Reads from both Active CNs and Expired CNs so the full history of issued CNs
 * is covered (prevents re-issuing a CN for the same window even after expiry).
 *
 * Keyed by cnKey; value is the most-recently-seen eventsHash for that key.
 *
 * @returns {Map<string, { eventsHash: string }>}
 */
function buildLogIndex_() {
  const index = new Map();

  const activeCNs  = readAllCNsFromSheet_(getOrCreateActiveCNsSheet_());
  const expiredCNs = readAllCNsFromSheet_(getOrCreateExpiredCNsSheet_());

  for (const cn of [...activeCNs, ...expiredCNs]) {
    const key  = String(cn.cnKey      || '').trim();
    const hash = String(cn.eventsHash || '').trim();
    if (key) index.set(key, { eventsHash: hash });
  }

  return index;
}

/**
 * Builds a map of consumed event keys per employee from all ACTIVE CNs.
 * Events from expired CNs are released back into the pool — only active CNs
 * lock their events.
 *
 * Keyed by employeeId; value is a Set of "YYYY-MM-DD|CODE" strings.
 *
 * @returns {Map<string, Set<string>>}
 */
function buildConsumedEventsIndex_() {
  const index = new Map();
  const activeCNs = readAllCNsFromSheet_(getOrCreateActiveCNsSheet_());

  for (const cn of activeCNs) {
    const employeeId = String(cn.employeeId || '').trim();
    if (!employeeId) continue;

    const consumedEvents = Array.isArray(cn.consumedEvents) ? cn.consumedEvents : [];
    if (consumedEvents.length === 0) continue;

    if (!index.has(employeeId)) index.set(employeeId, new Set());
    const employeeSet = index.get(employeeId);
    consumedEvents.forEach(eventKey => employeeSet.add(eventKey));
  }

  return index;
}


// ---------------------------------------------------------------------------
// CN Write
// ---------------------------------------------------------------------------

/**
 * Appends one CN record to the employee's row in the Active CNs sheet.
 * Finds the employee's row (or creates it) and pushes the new record into
 * the JSON array stored in the CNs (JSON) column.
 *
 * @param {CNProposal} proposal  — The CN proposal that was just issued.
 * @param {string}     issuedAt  — Formatted timestamp string ("yyyy-MM-dd HH:mm:ss").
 * @param {string}     issuedBy  — Email of the issuing user.
 * @param {string}     timeZone  — For date formatting.
 */
function appendActiveCN_(proposal, issuedAt, issuedBy, timeZone) {
  const formatDate = d => Utilities.formatDate(d, timeZone, 'yyyy-MM-dd');

  const cnRecord = {
    cnKey:               proposal.cnKey,
    status:              'proposed',
    windowStart:         formatDate(proposal.windowStart),
    windowEnd:           formatDate(proposal.windowEnd),
    count:               proposal.count,
    eventsHash:          proposal.eventsHash,
    issuedAt:            issuedAt,
    issuedBy:            issuedBy || '',
    sheetName:           proposal.sheetName || '',
    rule:                proposal.rule || 'GLOBAL',
    sourceSpreadsheetId: proposal.sourceSpreadsheetId || '',
    sourceSheetGid:      proposal.sourceSheetGid != null ? proposal.sourceSheetGid : '',
    consumedEvents:      (proposal.events || []).map(
      e => `${formatDate(e.date)}|${e.code}`
    ),
  };

  const sheet  = getOrCreateActiveCNsSheet_();
  const existing = readEmployeeCNs_(sheet, proposal.employeeId);
  console.log(`cnLog: appendActiveCN_ — writing CN for ${proposal.employeeName} (${proposal.employeeId}), existing count: ${existing.length}`);
  existing.push(cnRecord);
  writeEmployeeCNs_(
    sheet,
    proposal.employeeId,
    proposal.employeeName || '',
    proposal.department   || '',
    existing
  );
}


/**
 * Approves a proposed CN — changes its status from "proposed" to "active".
 *
 * @param {string} cnKey      — The CN_Key of the record to approve.
 * @param {string} employeeId — The employee whose row to update.
 */
function approveCN_(cnKey, employeeId) {
  const sheet = getOrCreateActiveCNsSheet_();
  const row   = findEmployeeRow_(sheet, employeeId);
  if (row === -1) throw new Error(`Employee ${employeeId} not found in Active CNs.`);

  const rowData      = sheet.getRange(row, 1, 1, CN_STORE_HEADERS.length).getValues()[0]; // config.js
  const employeeName = String(rowData[CN_STORE_COL.employeeName - 1] || '');
  const department   = String(rowData[CN_STORE_COL.department   - 1] || '');
  const rawJson      = String(rowData[CN_STORE_COL.cnsJson      - 1] || '');

  let cnArray;
  try { cnArray = JSON.parse(rawJson); } catch (e) { cnArray = []; }

  const targetIndex = cnArray.findIndex(cn => cn.cnKey === cnKey);
  if (targetIndex === -1) throw new Error(`CN "${cnKey}" not found for employee ${employeeId}.`);

  cnArray[targetIndex] = Object.assign({}, cnArray[targetIndex], { status: 'active' });
  writeEmployeeCNs_(sheet, employeeId, employeeName, department, cnArray);
}

/**
 * Rejects a proposed CN — removes it from Active CNs and pushes it to
 * (Expired CNs) with status "rejected".
 *
 * @param {string} cnKey      — The CN_Key of the record to reject.
 * @param {string} employeeId — The employee whose row to update.
 */
function rejectCN_(cnKey, employeeId) {
  const timeZone     = Session.getScriptTimeZone();
  const expiredStamp = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
  const activeSheet  = getOrCreateActiveCNsSheet_();
  const expiredSheet = getOrCreateExpiredCNsSheet_();

  const activeRow = findEmployeeRow_(activeSheet, employeeId);
  if (activeRow === -1) throw new Error(`Employee ${employeeId} not found in Active CNs.`);

  const rowData      = activeSheet.getRange(activeRow, 1, 1, CN_STORE_HEADERS.length).getValues()[0];
  const employeeName = String(rowData[CN_STORE_COL.employeeName - 1] || '');
  const department   = String(rowData[CN_STORE_COL.department   - 1] || '');
  const rawJson      = String(rowData[CN_STORE_COL.cnsJson      - 1] || '');

  let cnArray;
  try { cnArray = JSON.parse(rawJson); } catch (e) { cnArray = []; }

  const targetIndex = cnArray.findIndex(cn => cn.cnKey === cnKey);
  if (targetIndex === -1) throw new Error(`CN "${cnKey}" not found for employee ${employeeId}.`);

  const rejectedRecord = Object.assign({}, cnArray[targetIndex], {
    status:    'rejected',
    expiredAt: expiredStamp,
  });
  const remaining = cnArray.filter((_, index) => index !== targetIndex);

  writeEmployeeCNs_(activeSheet, employeeId, employeeName, department, remaining);

  const existingExpired = readEmployeeCNs_(expiredSheet, employeeId);
  existingExpired.push(rejectedRecord);
  writeEmployeeCNs_(expiredSheet, employeeId, employeeName, department, existingExpired);

  if (!expiredSheet.isSheetHidden()) expiredSheet.hideSheet();
}

/**
 * Returns all CN records for a single employee from the (Expired CNs) sheet.
 *
 * @param {string} employeeId
 * @returns {Object[]}
 */
function getExpiredCNsForEmployee_(employeeId) {
  return readEmployeeCNs_(getOrCreateExpiredCNsSheet_(), employeeId);
}


// ---------------------------------------------------------------------------
// CN Expiry
// ---------------------------------------------------------------------------

/**
 * Scans the Active CNs sheet for records that have passed EXPIRY_DAYS.
 *
 * For each expired CN record:
 *   1. Removes it from the employee's active array.
 *   2. Pushes it (with expiredAt) into the employee's row in (Expired CNs).
 *   3. Sends an expiry notification email if sendEmails is enabled.
 *
 * @param {boolean} dryRun — If true, logs what would happen but makes no changes.
 */
function expireCNsDaily(dryRun) {
  const config         = readCometConfig_(); // setup.js
  const shouldSendEmails = !!config.sendEmails;
  const timeZone       = Session.getScriptTimeZone();
  const activeSheet    = getOrCreateActiveCNsSheet_();
  const expiredSheet   = getOrCreateExpiredCNsSheet_();
  const lastRow        = activeSheet.getLastRow();

  if (lastRow < 3) {
    console.log('cnLog: Active CNs sheet is empty — nothing to expire.');
    return;
  }

  const now      = new Date();
  const expiryMs = EXPIRY_DAYS * 24 * 60 * 60 * 1000; // config.js
  let   expiredCount = 0;

  const rows = activeSheet.getRange(3, 1, lastRow - 2, CN_STORE_HEADERS.length).getValues();

  for (let rowOffset = 0; rowOffset < rows.length; rowOffset++) {
    const row          = rows[rowOffset];
    const employeeId   = String(row[CN_STORE_COL.employeeId   - 1] || '').trim();
    const employeeName = String(row[CN_STORE_COL.employeeName - 1] || '').trim();
    const department   = String(row[CN_STORE_COL.department   - 1] || '').trim();
    const rawJson      = String(row[CN_STORE_COL.cnsJson      - 1] || '').trim();

    if (!employeeId || !rawJson) continue;

    let cnArray;
    try { cnArray = JSON.parse(rawJson); } catch (e) { continue; }
    if (!Array.isArray(cnArray)) continue;

    const remaining = [];
    const toExpire  = [];

    for (const cn of cnArray) {
      // Only auto-expire approved (active) CNs; proposed ones need explicit
      // manager approval before they count toward expiry.
      if ((cn.status || 'proposed') !== 'active') {
        remaining.push(cn);
        continue;
      }
      const issuedAt = parseTimestamp_(String(cn.issuedAt || ''));
      if (!issuedAt || isNaN(issuedAt.getTime())) {
        remaining.push(cn);
        continue;
      }
      if (now.getTime() - issuedAt.getTime() >= expiryMs) {
        toExpire.push(cn);
      } else {
        remaining.push(cn);
      }
    }

    if (toExpire.length === 0) continue;

    const expiredStamp = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss');

    for (const cn of toExpire) {
      console.log(
        `cnLog: CN expired — ${employeeName} (${employeeId}) | ` +
        `Rule: ${cn.rule} | Window: ${cn.windowStart}–${cn.windowEnd}`
      );
    }

    if (!dryRun) {
      // Update the employee's active array (remove expired records)
      writeEmployeeCNs_(activeSheet, employeeId, employeeName, department, remaining);

      // Push expired records into the employee's row in (Expired CNs)
      const expiredExisting = readEmployeeCNs_(expiredSheet, employeeId);
      for (const cn of toExpire) {
        expiredExisting.push({ ...cn, expiredAt: expiredStamp });
      }
      writeEmployeeCNs_(expiredSheet, employeeId, employeeName, department, expiredExisting);

      // Keep (Expired CNs) hidden
      if (!expiredSheet.isSheetHidden()) expiredSheet.hideSheet();
    }

    // Send expiry notification emails (construction and sending live in notifier.js)
    for (const cn of toExpire) {
      sendCNExpiryNotification_(employeeName, employeeId, department, cn, expiredStamp, dryRun, shouldSendEmails); // notifier.js
    }

    expiredCount += toExpire.length;
  }

  console.log(`cnLog: Expiry scan complete — ${expiredCount} CN(s) expired.`);
}


// ---------------------------------------------------------------------------
// Dashboard Read
// ---------------------------------------------------------------------------

/**
 * Returns all active CN records as plain objects for the web UI dashboard.
 * Computes daysUntilExpiry from issuedAt so the frontend needs no date math.
 *
 * @returns {Array<{
 *   cnKey:           string,
 *   employeeName:    string,
 *   employeeId:      string,
 *   department:      string,
 *   rule:            string,
 *   count:           number,
 *   windowStart:     string,
 *   windowEnd:       string,
 *   issuedAt:        string,
 *   sheetName:       string,
 *   daysUntilExpiry: number|null,
 * }>}
 */
function getActiveCNsForDashboard_() {
  const now    = new Date();
  const allCNs = readAllCNsFromSheet_(getOrCreateActiveCNsSheet_());

  return allCNs.map(cn => {
    let daysUntilExpiry = null;
    const issuedAt = parseTimestamp_(String(cn.issuedAt || ''));
    if (issuedAt && !isNaN(issuedAt.getTime())) {
      const elapsedDays = (now.getTime() - issuedAt.getTime()) / (24 * 60 * 60 * 1000);
      daysUntilExpiry   = Math.max(0, Math.round(EXPIRY_DAYS - elapsedDays)); // config.js
    }

    return {
      cnKey:           String(cn.cnKey           || '').trim(),
      status:          String(cn.status          || 'proposed').trim(),
      employeeName:    String(cn.employeeName     || '').trim(),
      employeeId:      String(cn.employeeId       || '').trim(),
      department:      String(cn.department       || '').trim(),
      rule:            String(cn.rule             || '').trim(),
      count:           Number(cn.count)           || 0,
      windowStart:     String(cn.windowStart      || '').trim(),
      windowEnd:       String(cn.windowEnd        || '').trim(),
      issuedAt:        String(cn.issuedAt         || '').trim(),
      sheetName:       String(cn.sheetName        || '').trim(),
      daysUntilExpiry,
    };
  });
}


// ---------------------------------------------------------------------------
// Formatting Utility
// ---------------------------------------------------------------------------

/**
 * Applies a standard header row style to the given range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 * @param {string} bgColor   — Hex background color
 * @param {string} textColor — Hex text color
 */
function applyHeaderStyle_(range, bgColor, textColor) {
  range
    .setBackground(bgColor)
    .setFontColor(textColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}


// ---------------------------------------------------------------------------
// Timestamp Parsing
// ---------------------------------------------------------------------------

/**
 * Parses a "yyyy-MM-dd HH:mm:ss" timestamp string into a Date object.
 * Returns null if the string does not match the expected format.
 *
 * @param {string} timestampString
 * @returns {Date|null}
 */
function parseTimestamp_(timestampString) {
  if (!timestampString) return null;
  const match = String(timestampString).match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/
  );
  if (!match) return null;
  return new Date(
    parseInt(match[1], 10),
    parseInt(match[2], 10) - 1,
    parseInt(match[3], 10),
    parseInt(match[4], 10),
    parseInt(match[5], 10),
    parseInt(match[6], 10)
  );
}
