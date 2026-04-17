/**
 * notifier.js — Recipient resolution, email body construction, and digest sending.
 * VERSION: 1.0.0
 *
 * This file is the final stage of the notification pipeline. It receives a filtered
 * list of AbsenceRecord objects from dataIngestion.js and is responsible for:
 *
 *   1. GROUPING:    Organizing records by department so that each department's
 *                   manager receives a single digest for the window rather than
 *                   one email per absence.
 *
 *   2. RECIPIENTS:  Looking up who should receive each department's digest.
 *                   Matching is case-insensitive so that minor capitalization
 *                   differences in column H do not silently drop notifications.
 *                   Any department with no match falls back to FALLBACK_RECIPIENTS.
 *
 *   3. EMAIL BODY:  Building a readable plain-text summary of the absences in
 *                   the window, including the time range, employee names, reasons,
 *                   and any comments left by the employees.
 *
 *   4. SENDING:     Calling GmailApp.sendEmail once per department. To run the
 *                   script in dry-run mode without sending emails, comment out
 *                   the GmailApp.sendEmail line in sendDepartmentDigests_() and
 *                   replace it with the console.log dry-run line shown below it.
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Groups the given absence records by department and sends one digest email
 * per department that has records in the current window.
 *
 * This is the only function called from outside this file. It orchestrates the
 * three sub-steps (group → resolve recipients → send) without containing any
 * grouping, lookup, or formatting logic itself.
 *
 * @param {AbsenceRecord[]} absenceRecords — Filtered records from dataIngestion.js.
 * @param {{ start: Date, end: Date }} window — The time window these records belong to.
 * @param {string} timeZone — The script's time zone string, used for date formatting.
 */
function sendDepartmentDigests_(absenceRecords, window, timeZone) {
  const recordsByDepartment = groupRecordsByDepartment_(absenceRecords);

  Object.keys(recordsByDepartment).forEach(department => {
    const departmentRecords = recordsByDepartment[department];
    const recipients = resolveRecipientsForDepartment_(department);

    if (recipients.length === 0) {
      console.log(`notifier: No recipients found for department "${department}". Skipping.`);
      return;
    }

    // Sort records within each department by call time (ascending) so the
    // email reads chronologically regardless of row order in the sheet.
    const sortedRecords = departmentRecords.slice().sort((first, second) => {
      return first.calledAt.getTime() - second.calledAt.getTime();
    });

    const subject = buildEmailSubject_(department, sortedRecords);
    const body = buildEmailBody_(department, sortedRecords, window, timeZone);

    GmailApp.sendEmail(recipients.join(','), subject, body);

    console.log(`notifier: Sent digest for "${department}" to ${recipients.join(', ')} — ${sortedRecords.length} record(s).`);
  });
}


// ---------------------------------------------------------------------------
// Step 1: Grouping
// ---------------------------------------------------------------------------

/**
 * Groups an array of AbsenceRecord objects by their department field.
 *
 * Returns an object where each key is a department name and each value is the
 * array of records belonging to that department. Records with an empty department
 * field are grouped under an empty string key ("") and will be routed to
 * FALLBACK_RECIPIENTS by resolveRecipientsForDepartment_().
 *
 * @param {AbsenceRecord[]} records — The filtered absence records to group.
 * @returns {Object.<string, AbsenceRecord[]>} Records indexed by department name.
 */
function groupRecordsByDepartment_(records) {
  return records.reduce((groupedSoFar, record) => {
    const departmentKey = record.department || '';
    if (!groupedSoFar[departmentKey]) {
      groupedSoFar[departmentKey] = [];
    }
    groupedSoFar[departmentKey].push(record);
    return groupedSoFar;
  }, {});
}


// ---------------------------------------------------------------------------
// Step 2: Recipient Resolution
// ---------------------------------------------------------------------------

/**
 * Returns the list of email addresses that should receive a digest for the
 * given department name.
 *
 * Lookup strategy:
 *   1. Try an exact match against MAILING_LIST keys.
 *   2. If no exact match, try a case-insensitive match.
 *   3. If still no match, or if the department string is empty, use FALLBACK_RECIPIENTS.
 *
 * All values from MAILING_LIST are normalized before being returned. An entry
 * may be a plain string, an array of strings, or a string containing multiple
 * addresses separated by commas or semicolons — all are flattened into a clean
 * array of individual email addresses with duplicates removed.
 *
 * @param {string} department — The department name from column H.
 * @returns {string[]} A deduplicated array of email address strings.
 */
function resolveRecipientsForDepartment_(department) {
  const departmentKey = (department == null ? '' : String(department)).trim();

  let rawRecipientList = [];

  if (departmentKey && MAILING_LIST) {
    // Try exact key match first (fastest path for correctly-cased entries)
    const exactMatch = MAILING_LIST[departmentKey];
    if (exactMatch) {
      rawRecipientList = Array.isArray(exactMatch) ? exactMatch.slice() : [exactMatch];
    }

    // Fall back to a case-insensitive scan if the exact match produced nothing
    if (rawRecipientList.length === 0) {
      const lowercasedDepartment = departmentKey.toLowerCase();
      const matchingKey = Object.keys(MAILING_LIST).find(
        key => key.toLowerCase() === lowercasedDepartment
      );
      if (matchingKey) {
        const caseInsensitiveMatch = MAILING_LIST[matchingKey];
        rawRecipientList = Array.isArray(caseInsensitiveMatch)
          ? caseInsensitiveMatch.slice()
          : [caseInsensitiveMatch];
      }
    }
  }

  // If neither lookup found anything, use the configured fallback list
  if (rawRecipientList.length === 0) {
    rawRecipientList = (FALLBACK_RECIPIENTS || []).slice();
  }

  // Normalize: split any comma/semicolon-separated strings into individual addresses,
  // trim whitespace, remove empty strings, and deduplicate.
  const splitOnDelimiters = (emailString) =>
    String(emailString).split(/[;,]+/).map(address => address.trim()).filter(Boolean);

  const flattenedAddresses = rawRecipientList
    .reduce((flat, item) => flat.concat(Array.isArray(item) ? item : [item]), [])
    .flatMap(splitOnDelimiters);

  return Array.from(new Set(flattenedAddresses));
}


// ---------------------------------------------------------------------------
// Step 3 & 4: Email Construction and Sending
// ---------------------------------------------------------------------------

/**
 * Builds the email subject line for a department's digest.
 *
 * The subject follows the format:
 *   "Call Log Update - Week Ending MM/DD/YY - Department (N)"
 * where N is the number of absences in this window.
 *
 * The count in parentheses allows managers to see at a glance whether an email
 * is a single unusual event or a batch of multiple call-outs.
 *
 * @param {string}          department — The department name.
 * @param {AbsenceRecord[]} records    — The records being included in this email.
 * @returns {string} The complete subject line.
 */
function buildEmailSubject_(department, records) {
  const sheetTitle = getActiveCallLogSheetName_(); // defined in sheetUtils.js
  const displayName = department || 'Unknown Department';
  const recordCount = records.length;
  return `Call Log Update - ${sheetTitle} - ${displayName} (${recordCount})`;
}

/**
 * Builds a plain-text email body summarizing the absences for one department.
 *
 * The body starts with a header block showing the department name, the time
 * window covered, and a total count. Each absence is then listed with:
 *   - A sequential number (#1, #2, ...)
 *   - The time the employee called in
 *   - The employee's name and ID
 *   - The reason for the absence
 *   - Any comment the employee left
 *   - The spreadsheet row number (useful when a manager needs to look up the original entry)
 *
 * @param {string}          department     — The department name.
 * @param {AbsenceRecord[]} records        — Records sorted chronologically.
 * @param {{ start: Date, end: Date }} window — The time window for the header.
 * @param {string}          timeZone       — Used to format dates and times correctly.
 * @returns {string} The complete plain-text email body.
 */
function buildEmailBody_(department, records, window, timeZone) {
  const formatDateTime = (date) => Utilities.formatDate(date, timeZone, 'MMM d, yyyy h:mm a');

  // For manually sent single-row notifications, start === end; show a single
  // "Sent at" time rather than an identical "X – X" range that looks like an error.
  const windowLine = (window.start.getTime() === window.end.getTime())
    ? `Sent at:     ${formatDateTime(window.end)}`
    : `Window:      ${formatDateTime(window.start)} – ${formatDateTime(window.end)}`;

  const header = [
    `Department:  ${department || 'Unknown'}`,
    windowLine,
    `Absences:    ${records.length}`,
    '--------',
  ];

  const recordLines = records.map((record, index) => {
    const commentDisplay = record.employeeComment ? record.employeeComment : '(none)';
    return [
      `#${index + 1}`,
      `  Time:     ${formatDateTime(record.calledAt)}`,
      `  Employee: ${record.employeeName} (ID: ${record.employeeId})`,
      `  Reason:   ${record.absenceReason}`,
      `  Comment:  ${commentDisplay}`,
      `  Sheet Row: ${record.rowNumber}`,
    ].join('\n');
  });

  return header.concat(recordLines).join('\n');
}
