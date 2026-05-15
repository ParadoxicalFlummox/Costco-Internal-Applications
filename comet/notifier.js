/**
 * notifier.js — COMET email construction and sending.
 * VERSION: 0.3.0
 *
 * This file owns ALL outbound email notifications in COMET:
 *
 *   1. CN PROPOSALS: Sent when the infraction engine proposes a new CN.
 *      Entry point: sendCNNotifications_(proposals, dryRun, timeZone)
 *
 *   2. CN EXPIRY: Sent when a CN automatically expires after EXPIRY_DAYS.
 *      Entry point: sendCNExpiryNotification_(employeeName, employeeId,
 *                   department, cn, expiredStamp, dryRun, shouldSendEmails)
 *
 *   3. ABSENCE NOTIFICATIONS: Sent when a call-out or no-show is logged.
 *      Entry point: sendAbsenceEmail_(entry, recipients)
 *                   resolveAbsenceRecipients_(department)
 *
 * All emails follow a shared layout convention:
 *   [COMET] subject prefix | Unicode divider | aligned key-value pairs |
 *   optional detail section | closing divider | standard footer.
 *
 * Recipient routing:
 *   - CN emails  → PAYROLL_RECIPIENTS (config.js)
 *   - Absence emails → department mailing list from the Settings sheet
 */


// ---------------------------------------------------------------------------
// CN Proposal — Entry Point
// ---------------------------------------------------------------------------

/**
 * Sends (or dry-run logs) one email per CN proposal.
 *
 * Called by infractionEngine.js after deduplication. Each proposal in the
 * array has already been confirmed as new (not previously logged with the
 * same evidence).
 *
 * @param {CNProposal[]} proposals — Proposals to notify on.
 * @param {boolean}      dryRun    — If true, logs but does not send.
 * @param {string}       timeZone  — Script time zone for date formatting.
 */
function sendCNNotifications_(proposals, dryRun, timeZone) {
  if (!proposals || proposals.length === 0) return;

  proposals.forEach(proposal => {
    const recipients = resolveCNRecipients_();
    const subject    = buildCNProposalSubject_(proposal);
    const body       = buildCNProposalBody_(proposal, timeZone);

    if (!dryRun) {
      try {
        GmailApp.sendEmail(recipients.join(','), subject, body);
        console.log(`notifier: CN proposal email sent — ${proposal.employeeName} | Subject: ${subject}`);
      } catch (error) {
        console.error(`notifier: Failed to send CN proposal email — ${error.message}`);
      }
    } else {
      console.log(
        `notifier: [DRY RUN] Would send CN proposal email:\n` +
        `  To: ${recipients.join(', ')}\n` +
        `  Subject: ${subject}\n${body}`
      );
    }
  });
}


// ---------------------------------------------------------------------------
// CN Expiry — Entry Point
// ---------------------------------------------------------------------------

/**
 * Sends (or logs) an expiry notification for a single CN record.
 *
 * Called by cnLog.js inside expireCNsDaily for each CN that has passed
 * EXPIRY_DAYS. Centralised here so cnLog owns lifecycle logic and notifier
 * owns all email construction and sending.
 *
 * @param {string}  employeeName
 * @param {string}  employeeId
 * @param {string}  department
 * @param {Object}  cn               — The CN record being expired.
 * @param {string}  expiredStamp     — "yyyy-MM-dd HH:mm:ss" timestamp.
 * @param {boolean} dryRun
 * @param {boolean} shouldSendEmails
 */
function sendCNExpiryNotification_(employeeName, employeeId, department, cn, expiredStamp, dryRun, shouldSendEmails) {
  const recipients = resolveCNRecipients_();
  const subject    = buildCNExpirySubject_(employeeName, employeeId, department, cn.windowStart, cn.windowEnd, cn.rule);
  const body       = buildCNExpiryBody_(employeeName, employeeId, department, cn.windowStart, cn.windowEnd, cn.rule, cn.issuedAt, expiredStamp);

  if (!dryRun && shouldSendEmails) {
    try {
      GmailApp.sendEmail(recipients.join(','), subject, body);
      console.log(`notifier: CN expiry email sent for ${employeeName}.`);
    } catch (error) {
      console.error(`notifier: Failed to send CN expiry email — ${error.message}`);
    }
  } else if (dryRun) {
    console.log(`notifier: [DRY RUN] Would send CN expiry email:\n  Subject: ${subject}\n${body}`);
  } else {
    console.log(`notifier: CN expired (emails disabled) — ${employeeName}. Would send:\n  Subject: ${subject}\n${body}`);
  }
}


// ---------------------------------------------------------------------------
// Absence — Entry Points
// ---------------------------------------------------------------------------

/**
 * Builds and sends an absence notification email.
 *
 * Called by callLog.js after it has already validated the entry, confirmed
 * the email hasn't been sent, and resolved the recipient list. All sheet
 * writes (marking the row as sent) are handled by the caller.
 *
 * @param {CallLogEntry} entry      — The parsed absence record.
 * @param {string[]}     recipients — Resolved list of email addresses.
 */
function sendAbsenceEmail_(entry, recipients) {
  const { subject, body } = buildAbsenceEmailContent_(entry);
  GmailApp.sendEmail(recipients.join(','), subject, body);
  console.log(`notifier: Absence email sent — ${entry.name} (${entry.department}) to ${recipients.join(', ')}.`);
}

/**
 * Resolves the mailing list for absence notifications for the given department.
 * Reads from the department's "mailing" field in the Settings sheet.
 * Returns an empty array if no recipients are configured.
 *
 * @param {string} department
 * @returns {string[]}
 */
function resolveAbsenceRecipients_(department) {
  try {
    const settings = ensureDeptSettingsBaseStructure_(department); // scheduleSettings.js
    if (settings.mailing && Array.isArray(settings.mailing) && settings.mailing.length > 0) {
      return settings.mailing;
    }
    return [];
  } catch (error) {
    console.error(`notifier: Error resolving absence recipients for "${department}" — ${error.message}`);
    return [];
  }
}


// ---------------------------------------------------------------------------
// Recipient Resolution — CN (Payroll)
// ---------------------------------------------------------------------------

/**
 * Returns the list of email recipients for CN notifications.
 * All CN emails (proposals and expiry) route to PAYROLL_RECIPIENTS.
 *
 * @returns {string[]}
 */
function resolveCNRecipients_() {
  const raw   = (PAYROLL_RECIPIENTS || []).slice(); // config.js
  const split = s => String(s).split(/[;,]+/).map(addr => addr.trim()).filter(Boolean);
  const flat  = raw.flatMap(item => Array.isArray(item) ? item.flatMap(split) : split(item));
  return Array.from(new Set(flat));
}


// ---------------------------------------------------------------------------
// CN Proposal Email Construction
// ---------------------------------------------------------------------------

/**
 * Builds the subject line for a CN proposal email.
 *
 * Format: "[COMET] CN Proposed — First Last (MAINTENANCE) — 3× TD (Apr 1 – Apr 13)"
 *
 * @param {CNProposal} proposal
 * @returns {string}
 */
function buildCNProposalSubject_(proposal) {
  const name       = proposal.employeeName || 'Unknown Employee';
  const department = proposal.department   || 'Unknown';
  const rulePart   = proposal.rule && proposal.rule !== 'GLOBAL'
    ? `${proposal.count}× ${proposal.rule}`
    : `${proposal.count} infraction(s)`;
  const startFmt = formatShortDate_(proposal.windowStart);
  const endFmt   = formatShortDate_(proposal.windowEnd);

  return `[COMET] CN Proposed — ${name} (${department}) — ${rulePart} (${startFmt} – ${endFmt})`;
}

/**
 * Builds the plain-text body for a CN proposal email.
 *
 * @param {CNProposal} proposal
 * @param {string}     timeZone
 * @returns {string}
 */
function buildCNProposalBody_(proposal, timeZone) {
  const formatDate = date => Utilities.formatDate(date, timeZone, 'MMMM d, yyyy');
  const divider    = '──────────────────────────';

  const ruleDetail = proposal.rule && proposal.rule !== 'GLOBAL'
    ? `${proposal.rule} — ${proposal.count} occurrence(s) in ${proposal.windowDays} days`
    : `${proposal.count} infraction(s) in ${proposal.windowDays} days`;

  const lines = [
    'COMET Counseling Notice',
    divider,
    `Employee:    ${proposal.employeeName || 'Unknown'}`,
    `ID:          ${proposal.employeeId   || '—'}`,
    `Department:  ${proposal.department   || '—'}`,
    `Rule:        ${ruleDetail}`,
    `Window:      ${formatDate(proposal.windowStart)} — ${formatDate(proposal.windowEnd)}`,
    '',
  ];

  proposal.events.forEach((event, index) => {
    lines.push(`  ${index + 1}.  ${formatDate(event.date)}   ${event.code}`);
  });

  lines.push(
    '',
    divider,
    'This CN has been recorded as Proposed and is awaiting manager approval.',
    'This notification was generated automatically by COMET.',
  );

  return lines.join('\n');
}


// ---------------------------------------------------------------------------
// CN Expiry Email Construction
// ---------------------------------------------------------------------------

/**
 * Builds the subject line for a CN expiry email.
 *
 * Format: "[COMET] CN Expired — First Last (MAINTENANCE) — TD (Apr 1, 2026 – Apr 13, 2026)"
 *
 * @param {string} employeeName
 * @param {string} employeeId
 * @param {string} department
 * @param {string} windowStart  — YYYY-MM-DD
 * @param {string} windowEnd    — YYYY-MM-DD
 * @param {string} rule
 * @returns {string}
 */
function buildCNExpirySubject_(employeeName, employeeId, department, windowStart, windowEnd, rule) {
  const dept      = department || 'Unknown';
  const ruleLabel = rule ? `${rule} ` : '';
  const startFmt  = formatCNDateString_(windowStart);
  const endFmt    = formatCNDateString_(windowEnd);
  return `[COMET] CN Expired — ${employeeName} (${dept}) — ${ruleLabel}(${startFmt} – ${endFmt})`;
}

/**
 * Builds the plain-text body for a CN expiry email.
 *
 * @param {string} employeeName
 * @param {string} employeeId
 * @param {string} department
 * @param {string} windowStart  — YYYY-MM-DD
 * @param {string} windowEnd    — YYYY-MM-DD
 * @param {string} rule
 * @param {string} issuedAt     — "yyyy-MM-dd HH:mm:ss"
 * @param {string} expiredStamp — "yyyy-MM-dd HH:mm:ss"
 * @returns {string}
 */
function buildCNExpiryBody_(employeeName, employeeId, department, windowStart, windowEnd, rule, issuedAt, expiredStamp) {
  const divider = '──────────────────────────';

  return [
    'COMET Counseling Notice Expired',
    divider,
    `Employee:    ${employeeName || 'Unknown'}`,
    `ID:          ${employeeId   || '—'}`,
    `Department:  ${department   || '—'}`,
    `Rule:        ${rule         || '—'}`,
    `Window:      ${formatCNDateString_(windowStart)} — ${formatCNDateString_(windowEnd)}`,
    `Issued:      ${issuedAt     || '—'}`,
    `Expired:     ${expiredStamp || '—'}`,
    '',
    divider,
    `This Counseling Notice has automatically expired after ${EXPIRY_DAYS} days.`, // config.js
    'This notification was generated automatically by COMET.',
  ].join('\n');
}


// ---------------------------------------------------------------------------
// Absence Email Construction
// ---------------------------------------------------------------------------

/**
 * Builds the subject and plain-text body for an absence notification email.
 *
 * @param {CallLogEntry} entry
 * @returns {{ subject: string, body: string }}
 */
function buildAbsenceEmailContent_(entry) {
  const timeZone = Session.getScriptTimeZone();
  const dateStr  = entry.dateRaw
    ? Utilities.formatDate(coerceCallLogDate_(entry.dateRaw) || new Date(), timeZone, 'MMMM d, yyyy') // callLog.js
    : 'Unknown date';

  const types = [];
  if (entry.isCallout) types.push('Call-Out');
  if (entry.isFmla)    types.push('FMLA');
  if (entry.isNoShow)  types.push('No Show');
  const typeLabel = types.length > 0 ? types.join(', ') : 'Absence';

  const subject = `[COMET] ${typeLabel} — ${entry.name} (${entry.department}) — ${dateStr}`;

  const lines = [
    'COMET Absence Notification',
    '──────────────────────────',
    `Employee:        ${entry.name}`,
    `ID:              ${entry.employeeId || '—'}`,
    `Department:      ${entry.department || '—'}`,
    `Date:            ${dateStr}`,
    `Time Called:     ${entry.time || '—'}`,
    `Type:            ${typeLabel}`,
  ];

  if (entry.manager)       lines.push(`Manager:         ${entry.manager}`);
  if (entry.scheduledShift) lines.push(`Scheduled Shift: ${entry.scheduledShift}`);
  if (entry.comment)       lines.push(`Comment:         ${entry.comment}`);

  lines.push('', '──────────────────────────');
  lines.push('This notification was generated automatically by COMET.');

  return { subject, body: lines.join('\n') };
}


// ---------------------------------------------------------------------------
// Date Formatting Utilities
// ---------------------------------------------------------------------------

/**
 * Formats a Date object as a short subject-line string.
 * e.g. new Date(2026, 3, 1) → "Apr 1"
 *
 * @param {Date} date
 * @returns {string}
 */
function formatShortDate_(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[date.getMonth()]} ${date.getDate()}`;
}

/**
 * Formats a YYYY-MM-DD date string into a short human-readable form.
 * e.g. "2026-04-01" → "Apr 1, 2026"
 *
 * Used where dates are stored as strings rather than Date objects.
 *
 * @param {string} dateString — YYYY-MM-DD
 * @returns {string}
 */
function formatCNDateString_(dateString) {
  if (!dateString) return '—';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const match = String(dateString).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return dateString;
  const month = months[parseInt(match[2], 10) - 1] || match[2];
  return `${month} ${parseInt(match[3], 10)}, ${match[1]}`;
}
