/**
 * cnLog.js — CN_Log sheet management, Active CNs view, and CN expiry processing.
 * VERSION: 1.0.0
 *
 * This file owns all interactions with the three CN tracking sheets:
 *
 *   CN_Log         — The internal source of truth used for idempotent deduplication.
 *                    Every CN ever issued lives here permanently (Active or Expired).
 *                    Not intended for day-to-day manager review.
 *
 *   Active CNs     — The manager-facing view. Shows all currently Active CNs with
 *                    a clickable hyperlink in the Employee Name column that opens
 *                    the employee's tab in the attendance controller. Rows are
 *                    removed from here when a CN expires.
 *
 *   (Expired CNs)  — The hidden archive. When a CN expires, its row is moved here
 *                    from Active CNs and the sheet is kept hidden (parentheses
 *                    prefix = hidden in attendance controller convention). The full
 *                    record is preserved for audit purposes.
 *
 * All three sheets live in the same target workbook — either the external log
 * spreadsheet configured in "Infraction Config" B2, or the active workbook as
 * a fallback.
 *
 * SHEET LIFECYCLE:
 *   CN issued  → row appended to CN_Log (Status=Active) + row appended to Active CNs
 *   CN expires → CN_Log row Status updated to Expired + Active CNs row moved to
 *                (Expired CNs) + expiry email sent to payroll
 */


// ---------------------------------------------------------------------------
// Workbook Resolution
// ---------------------------------------------------------------------------

/**
 * Determines which spreadsheet should hold the CN tracking sheets and returns it.
 *
 * Reads the log spreadsheet ID from cell B2 of the "Infraction Config" sheet
 * in the active workbook. If configured and accessible, that external workbook
 * is returned. Otherwise, the active workbook is used as a fallback.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function resolveLogWorkbook_() {
  try {
    const activeWorkbook = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = activeWorkbook.getSheetByName(INFRACTION_CONFIG_SHEET_NAME); // config.js
    if (!configSheet) return activeWorkbook;

    const idValue = configSheet.getRange(LOG_SPREADSHEET_ID_CELL).getValue(); // config.js
    if (!idValue || typeof idValue !== 'string' || !idValue.trim()) {
      return activeWorkbook;
    }

    const externalWorkbook = SpreadsheetApp.openById(idValue.trim());
    console.log(`cnLog: Using external CN log workbook "${externalWorkbook.getName()}".`);
    return externalWorkbook;
  } catch (error) {
    console.warn(`cnLog: Could not open external log workbook — ${error.message}. Using active workbook.`);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}


// ---------------------------------------------------------------------------
// Sheet Access — CN_Log
// ---------------------------------------------------------------------------

/**
 * Returns the CN_Log sheet, creating it if it does not yet exist.
 *
 * If the sheet already exists but is missing columns (e.g. after a config
 * update added new headers), the missing columns are appended to the right.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateLogSheet_() {
  const workbook = resolveLogWorkbook_();
  let sheet = workbook.getSheetByName(CN_LOG_SHEET_NAME); // config.js

  if (!sheet) {
    sheet = workbook.insertSheet(CN_LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, CN_LOG_HEADERS.length).setValues([CN_LOG_HEADERS]); // config.js
    applyHeaderStyle_(sheet.getRange(`1:1`), '#263238', '#FFFFFF');
    console.log(`cnLog: Created CN_Log sheet in "${workbook.getName()}".`);
  } else {
    upgradeHeaders_(sheet, CN_LOG_HEADERS);
  }

  return sheet;
}


// ---------------------------------------------------------------------------
// Sheet Access — Active CNs
// ---------------------------------------------------------------------------

/**
 * Returns the "Active CNs" sheet, creating it if it does not yet exist.
 *
 * The Active CNs sheet is the manager-facing view of outstanding CNs.
 * It is formatted with a visible header row and auto-resized columns.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateActiveCNsSheet_() {
  const workbook = resolveLogWorkbook_();
  let sheet = workbook.getSheetByName(ACTIVE_CNS_SHEET_NAME); // config.js

  if (!sheet) {
    sheet = workbook.insertSheet(ACTIVE_CNS_SHEET_NAME);
    sheet.getRange(1, 1, 1, ACTIVE_CNS_HEADERS.length).setValues([ACTIVE_CNS_HEADERS]); // config.js
    applyHeaderStyle_(sheet.getRange('1:1'), '#B71C1C', '#FFFFFF'); // red header — draws attention
    sheet.setFrozenRows(1);

    // Set column widths for readability
    const widths = { 1: 220, 2: 180, 3: 100, 4: 120, 5: 60, 6: 55, 7: 110, 8: 110, 9: 160, 10: 180 };
    Object.entries(widths).forEach(([col, px]) => sheet.setColumnWidth(Number(col), px));

    console.log(`cnLog: Created Active CNs sheet in "${workbook.getName()}".`);
  }

  return sheet;
}


// ---------------------------------------------------------------------------
// Sheet Access — (Expired CNs)
// ---------------------------------------------------------------------------

/**
 * Returns the "(Expired CNs)" sheet, creating it (hidden) if it does not exist.
 *
 * The sheet is hidden on creation and kept hidden. It can be revealed for
 * audit purposes via right-click → Show Sheet. The parentheses prefix follows
 * the attendance controller convention for reference/archive sheets.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateExpiredCNsSheet_() {
  const workbook = resolveLogWorkbook_();
  let sheet = workbook.getSheetByName(EXPIRED_CNS_SHEET_NAME); // config.js

  if (!sheet) {
    sheet = workbook.insertSheet(EXPIRED_CNS_SHEET_NAME);
    sheet.getRange(1, 1, 1, EXPIRED_CNS_HEADERS.length).setValues([EXPIRED_CNS_HEADERS]); // config.js
    applyHeaderStyle_(sheet.getRange('1:1'), '#546E7A', '#FFFFFF'); // muted gray header
    sheet.setFrozenRows(1);

    const widths = { 1: 220, 2: 180, 3: 100, 4: 120, 5: 60, 6: 55, 7: 110, 8: 110, 9: 160, 10: 180, 11: 160 };
    Object.entries(widths).forEach(([col, px]) => sheet.setColumnWidth(Number(col), px));

    sheet.hideSheet(); // hidden by default — accessible via right-click → Show Sheet
    console.log(`cnLog: Created (Expired CNs) sheet (hidden) in "${workbook.getName()}".`);
  }

  return sheet;
}


// ---------------------------------------------------------------------------
// Log Index
// ---------------------------------------------------------------------------

/**
 * Builds an in-memory Map from the CN_Log for deduplication lookups.
 *
 * Keyed by CN_Key; stores the most recent EventsHash for that key.
 * infractionEngine.js uses this to skip proposals already logged with
 * the same evidence.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet
 * @returns {Map<string, { eventsHash: string }>}
 */
function buildLogIndex_(logSheet) {
  const index = new Map();
  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) return index;

  const keys = logSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const hashes = logSheet.getRange(2, 8, lastRow - 1, 1).getValues(); // EventsHash col 8

  for (let i = 0; i < keys.length; i++) {
    const key = String(keys[i][0] || '').trim();
    if (!key) continue;
    index.set(key, { eventsHash: String(hashes[i][0] || '').trim() });
  }

  return index;
}


// ---------------------------------------------------------------------------
// Row Writing
// ---------------------------------------------------------------------------

/**
 * Appends one CN record to the CN_Log sheet.
 *
 * Values are mapped from rowData using CN_LOG_HEADERS as the key list so
 * column order is always consistent with the header row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} logSheet
 * @param {Object} rowData — Keys match CN_LOG_HEADERS entries.
 */
function appendLogRow_(logSheet, rowData) {
  const values = CN_LOG_HEADERS.map(header => rowData[header] != null ? rowData[header] : '');
  logSheet.appendRow(values);
}

/**
 * Appends one row to the Active CNs manager-facing sheet.
 *
 * The Employee Name cell is written as a HYPERLINK formula that links
 * directly to the employee's tab in the attendance controller. If the
 * source spreadsheet ID or sheet GID is not available (e.g. the proposal
 * came from a local fallback), the name is written as plain text.
 *
 * @param {CNProposal} proposal     — The CN proposal that was just issued.
 * @param {string}     issuedAt     — Formatted timestamp string.
 * @param {string}     timeZone     — For date formatting.
 */
function appendActiveCNRow_(proposal, issuedAt, timeZone) {
  const sheet = getOrCreateActiveCNsSheet_();

  // Build the employee name cell — hyperlink if we have source location info,
  // plain text otherwise.
  let nameCell;
  if (proposal.sourceSpreadsheetId && proposal.sourceSheetGid != null) {
    const url = `https://docs.google.com/spreadsheets/d/${proposal.sourceSpreadsheetId}/edit#gid=${proposal.sourceSheetGid}`;
    const escapedName = String(proposal.employeeName || '').replace(/"/g, '""');
    nameCell = `=HYPERLINK("${url}","${escapedName}")`;
  } else {
    nameCell = proposal.employeeName || '';
  }

  const formatDate = d => Utilities.formatDate(d, timeZone, 'yyyy-MM-dd');

  // Build the row in ACTIVE_CNS_HEADERS order:
  // CN_Key | Employee Name | Employee ID | Department | Rule | Count |
  // Window Start | Window End | Issued At | Sheet
  const rowValues = [
    proposal.cnKey,
    nameCell,
    proposal.employeeId || '',
    proposal.department || '',
    proposal.rule || 'GLOBAL',
    proposal.count,
    formatDate(proposal.windowStart),
    formatDate(proposal.windowEnd),
    issuedAt,
    proposal.sheetName || '',
  ];

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, rowValues.length).setValues([rowValues]);

  // The name cell (column B = index 2) contains a formula — write it separately
  // so it is treated as a formula rather than a literal string.
  if (proposal.sourceSpreadsheetId && proposal.sourceSheetGid != null) {
    sheet.getRange(newRow, 2).setFormula(nameCell);
  }
}


// ---------------------------------------------------------------------------
// CN Expiry
// ---------------------------------------------------------------------------

/**
 * Scans the CN_Log for Active CNs that have passed EXPIRY_DAYS and processes each.
 *
 * For each expired CN:
 *   1. Updates CN_Log: Status → "Expired", ExpiredAt → current timestamp.
 *      (Written BEFORE the email send so the record is never lost on failure.)
 *   2. Moves the row from Active CNs to (Expired CNs) and hides that sheet.
 *   3. Sends an expiry notification email to payroll.
 *
 * @param {boolean} dryRun — If true, logs what would happen but makes no changes.
 */
function expireCNsDaily(dryRun) {
  const timeZone = Session.getScriptTimeZone();
  const logSheet = getOrCreateLogSheet_();
  const lastRow = logSheet.getLastRow();

  if (lastRow < 2) {
    console.log('cnLog: CN_Log is empty — nothing to expire.');
    return;
  }

  // Build a dynamic column index from the header row so we are resilient to
  // column reordering. col.X returns the 1-based column number for header "X".
  const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getDisplayValues()[0];
  const col = {};
  headerRow.forEach((name, i) => { col[String(name).trim()] = i + 1; });

  const allRows = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn()).getValues();
  const now = new Date();
  const expiryMs = EXPIRY_DAYS * 24 * 60 * 60 * 1000; // config.js
  let expiredCount = 0;

  allRows.forEach((row, rowOffset) => {
    const sheetRow = rowOffset + 2;
    const status = String(row[(col.Status || 13) - 1] || '').trim() || 'Active';
    if (status !== 'Active') return;

    const issuedStr = String(row[(col.IssuedAt || 9) - 1] || '').trim();
    const issuedAt = parseTimestamp_(issuedStr) || new Date(row[(col.IssuedAt || 9) - 1]);
    if (!issuedAt || isNaN(issuedAt.getTime())) return;
    if (now.getTime() - issuedAt.getTime() < expiryMs) return;

    // --- This CN has expired ---
    const cnKey = String(row[(col.CN_Key || 1) - 1] || '').trim();
    const employeeName = String(row[(col.EmployeeName || 3) - 1] || '').trim();
    const employeeId = String(row[(col.EmployeeID || 2) - 1] || '').trim();
    const department = String(row[(col.Department || 4) - 1] || '').trim();
    const windowStart = String(row[(col.WindowStart || 5) - 1] || '').trim();
    const windowEnd = String(row[(col.WindowEnd || 6) - 1] || '').trim();
    const rule = String(row[(col.Rule || 15) - 1] || '').trim();
    const sheetName = String(row[(col.SheetName || 12) - 1] || '').trim();
    const expiredStamp = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd HH:mm:ss');

    console.log(`cnLog: CN expired — ${employeeName} (${employeeId}) | Rule: ${rule} | Window: ${windowStart}–${windowEnd}`);

    if (!dryRun) {
      // Step 1: Update CN_Log status FIRST (before email, so record is never lost)
      if (col.Status) logSheet.getRange(sheetRow, col.Status).setValue('Expired');
      if (col.ExpiredAt) logSheet.getRange(sheetRow, col.ExpiredAt).setValue(expiredStamp);

      // Step 2: Move the row from Active CNs to (Expired CNs)
      moveToExpiredSheet_(cnKey, expiredStamp, employeeName, employeeId);
    }

    // Step 3: Send expiry notification
    const subject = buildExpiryEmailSubject_(employeeName, employeeId, windowStart, windowEnd, rule);
    const body = buildExpiryEmailBody_(employeeName, employeeId, department, windowStart, windowEnd, rule, issuedStr, expiredStamp);

    if (!dryRun) {
      try {
        GmailApp.sendEmail(PAYROLL_RECIPIENTS.join(','), subject, body); // config.js
        console.log(`cnLog: Expiry email sent for ${employeeName}.`);
      } catch (emailError) {
        console.error(`cnLog: Expiry email failed for row ${sheetRow} — ${emailError.message}`);
      }
    } else {
      console.log(`cnLog: [DRY RUN] Would expire CN and send:\n  Subject: ${subject}\n${body}`);
    }

    expiredCount++;
  });

  console.log(`cnLog: Expiry scan complete — ${expiredCount} CN(s) expired.`);
}

/**
 * Finds the row in Active CNs matching the given CN_Key, copies it to
 * (Expired CNs) with the expiry timestamp appended, then deletes it from
 * Active CNs. Ensures (Expired CNs) remains hidden after the move.
 *
 * If no matching row is found in Active CNs (e.g. it was manually deleted),
 * the function logs a warning and skips the move without throwing.
 *
 * @param {string} cnKey        — The CN_Key to find in Active CNs (column A).
 * @param {string} expiredStamp — The expiry timestamp string to append.
 * @param {string} employeeName — For logging only.
 * @param {string} employeeId   — For logging only.
 */
function moveToExpiredSheet_(cnKey, expiredStamp, employeeName, employeeId) {
  try {
    const activeSheet = getOrCreateActiveCNsSheet_();
    const expiredSheet = getOrCreateExpiredCNsSheet_();

    const lastActiveRow = activeSheet.getLastRow();
    if (lastActiveRow < 2) {
      console.warn(`cnLog: Active CNs sheet is empty — cannot move ${cnKey}.`);
      return;
    }

    // Find the row by CN_Key (column A)
    const keys = activeSheet.getRange(2, 1, lastActiveRow - 1, 1).getDisplayValues();
    let foundRow = -1;
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === cnKey) {
        foundRow = i + 2; // 1-based sheet row
        break;
      }
    }

    if (foundRow === -1) {
      console.warn(`cnLog: Row for CN_Key "${cnKey}" not found in Active CNs — skipping move.`);
      return;
    }

    // Read the full row from Active CNs
    const activeRowData = activeSheet
      .getRange(foundRow, 1, 1, ACTIVE_CNS_HEADERS.length)
      .getValues()[0];

    // Append to (Expired CNs) with the expiry timestamp as the last column
    const expiredRowData = activeRowData.concat([expiredStamp]);
    const newExpiredRow = expiredSheet.getLastRow() + 1;
    expiredSheet.getRange(newExpiredRow, 1, 1, expiredRowData.length).setValues([expiredRowData]);

    // The Employee Name cell in Active CNs may be a HYPERLINK formula — copy it
    // to (Expired CNs) as a formula so the link is preserved in the archive.
    const nameFormula = activeSheet.getRange(foundRow, 2).getFormula();
    if (nameFormula) {
      expiredSheet.getRange(newExpiredRow, 2).setFormula(nameFormula);
    }

    // Delete the row from Active CNs
    activeSheet.deleteRow(foundRow);

    // Keep (Expired CNs) hidden — deleteRow can un-hide a sheet in some GAS versions
    if (!expiredSheet.isSheetHidden()) {
      expiredSheet.hideSheet();
    }

    console.log(`cnLog: Moved CN for ${employeeName} (${employeeId}) from Active CNs to (Expired CNs).`);
  } catch (error) {
    console.error(`cnLog: Failed to move CN_Key "${cnKey}" to (Expired CNs) — ${error.message}`);
  }
}


// ---------------------------------------------------------------------------
// Expiry Email Builders
// ---------------------------------------------------------------------------

/**
 * Builds the subject line for a CN expiry notification email.
 */
function buildExpiryEmailSubject_(employeeName, employeeId, windowStart, windowEnd, rule) {
  const ruleLabel = rule ? ` [${rule}]` : '';
  return `CN Expired: ${employeeName} (${employeeId || 'N/A'})${ruleLabel} — ${windowStart} to ${windowEnd}`;
}

/**
 * Builds the plain-text body for a CN expiry notification email.
 */
function buildExpiryEmailBody_(employeeName, employeeId, department, windowStart, windowEnd, rule, issuedStr, expiredStamp) {
  return [
    `Employee:          ${employeeName} (${employeeId || 'N/A'})`,
    `Department:        ${department || 'Unknown'}`,
    `Original window:   ${windowStart} — ${windowEnd}${rule ? '  [' + rule + ']' : ''}`,
    `Originally issued: ${issuedStr}`,
    `Expired:           ${expiredStamp}`,
    '',
    `This Counseling Notice has automatically expired after ${EXPIRY_DAYS} days.`, // config.js
    'The record has been moved to the (Expired CNs) sheet in the CN Log workbook.',
    'No further action is required unless the employee\'s situation has changed.',
    '',
    'Auto-generated by the Costco Infraction Notifier.',
  ].join('\n');
}


// ---------------------------------------------------------------------------
// Header Upgrade Utility
// ---------------------------------------------------------------------------

/**
 * Appends any expected headers that are missing from the sheet's current
 * header row. Allows new columns to be added to a headers array in config.js
 * without requiring a manual sheet reset.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} expectedHeaders
 */
function upgradeHeaders_(sheet, expectedHeaders) {
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const existingSet = new Set(existing.map(h => String(h || '').trim()));
  const missing = expectedHeaders.filter(h => !existingSet.has(h));
  if (missing.length === 0) return;
  const merged = existing.concat(missing);
  sheet.getRange(1, 1, 1, merged.length).setValues([merged]);
  console.log(`cnLog: Added missing headers to "${sheet.getName()}": ${missing.join(', ')}`);
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
