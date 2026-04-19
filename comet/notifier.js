/**
 * notifier.js — CN email construction and sending.
 * VERSION: 0.1.0
 *
 * This file is the final stage of the notification pipeline. It receives a
 * list of new CN proposals from infractionEngine.js and is responsible for:
 *
 *   1. RECIPIENT RESOLUTION: All CNs currently route to PAYROLL_RECIPIENTS.
 *      Per-department routing can be added in a future version by introducing
 *      a MAILING_LIST map in config.js (same pattern as the absence notifier).
 *
 *   2. EMAIL CONSTRUCTION: Building a readable plain-text subject and body
 *      for each CN. The body lists the triggering events chronologically with
 *      their date, code, and cell reference so payroll can verify against the
 *      original attendance controller.
 *
 *   3. SENDING: Calling GmailApp.sendEmail once per CN proposal. In dry-run
 *      mode, the email content is logged but no email is sent.
 */


// ---------------------------------------------------------------------------
// Public Entry Point
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
    const subject = buildEmailSubject_(proposal);
    const body = buildEmailBody_(proposal, timeZone);

    if (!dryRun) {
      try {
        GmailApp.sendEmail(recipients.join(','), subject, body);
        console.log(
          `notifier: CN email sent — To: ${recipients.join(', ')} | Subject: ${subject}`
        );
      } catch (error) {
        console.error(`notifier: Failed to send CN email — ${error.message}`);
      }
    } else {
      console.log(
        `notifier: [DRY RUN] Would send CN email:\n` +
        `  To: ${recipients.join(', ')}\n` +
        `  Subject: ${subject}\n` +
        body
      );
    }
  });
}


// ---------------------------------------------------------------------------
// Recipient Resolution
// ---------------------------------------------------------------------------

/**
 * Returns the list of email recipients for a CN notification.
 *
 * Currently all CNs route to PAYROLL_RECIPIENTS regardless of department.
 * The per-department mailing list pattern from the absence notifier can be
 * added here in a future version — just introduce a MAILING_LIST object in
 * config.js and add a department lookup before falling back to payroll.
 *
 * @returns {string[]} Deduplicated array of email address strings.
 */
function resolveCNRecipients_() {
  const raw = (PAYROLL_RECIPIENTS || []).slice(); // config.js
  const split = s => String(s).split(/[;,]+/).map(addr => addr.trim()).filter(Boolean);
  const flat = raw.flatMap(item => Array.isArray(item) ? item.flatMap(split) : split(item));
  return Array.from(new Set(flat));
}


// ---------------------------------------------------------------------------
// Email Construction
// ---------------------------------------------------------------------------

/**
 * Builds the subject line for a CN notification email.
 *
 * Format:
 *   "CN Alert: First Last (ID) — 3x TD in 30 days (Apr 1 – Apr 13, 2026)"
 *
 * @param {CNProposal} proposal
 * @returns {string}
 */
function buildEmailSubject_(proposal) {
  const name = proposal.employeeName || 'Unknown Employee';
  const idPart = proposal.employeeId ? ` (${proposal.employeeId})` : '';
  const rulePart = proposal.rule && proposal.rule !== 'GLOBAL'
    ? `${proposal.count}x ${proposal.rule}`
    : `${proposal.count} infraction(s)`;
  const daysPart = `in ${proposal.windowDays} days`;
  const startFmt = formatShortDate_(proposal.windowStart);
  const endFmt = formatShortDate_(proposal.windowEnd);

  return `CN Alert: ${name}${idPart} — ${rulePart} ${daysPart} (${startFmt} – ${endFmt})`;
}

/**
 * Builds the plain-text email body for a CN notification.
 *
 * The body includes all identifying details for the employee and the
 * triggering event window, followed by a numbered list of each event
 * with date, code, and spreadsheet cell reference.
 *
 * @param {CNProposal} proposal
 * @param {string}     timeZone
 * @returns {string}
 */
function buildEmailBody_(proposal, timeZone) {
  const formatDate = date => Utilities.formatDate(date, timeZone, 'MMM d, yyyy');

  const header = [
    `Employee:   ${proposal.employeeName || 'Unknown'} (ID: ${proposal.employeeId || 'N/A'})`,
    `Department: ${proposal.department || 'Unknown'}`,
    `Rule:       ${proposal.rule || 'GLOBAL'} — ${proposal.count} occurrence(s) in ${proposal.windowDays} days`,
    `Window:     ${formatDate(proposal.windowStart)} — ${formatDate(proposal.windowEnd)}`,
    `Sheet:      ${proposal.sheetName || 'Unknown'}`,
    '--------',
    'Events:',
  ];

  const eventLines = proposal.events.map((event, index) => {
    const cellRef = event.a1 ? ` (cell ${event.a1})` : '';
    return `  #${index + 1}  ${formatDate(event.date)}  |  ${event.code}${cellRef}`;
  });

  const footer = [
    '',
    'Please review and issue a Counseling Notice if appropriate.',
    'This CN has been recorded in the CN_Log with status: Active.',
    '',
    'Auto-generated by the Costco Infraction Notifier.',
  ];

  return header.concat(eventLines).concat(footer).join('\n');
}


// ---------------------------------------------------------------------------
// Utility
// ---------------------------------------------------------------------------

/**
 * Formats a Date as a short human-readable string for the email subject.
 * e.g. new Date(2026, 3, 1) → "Apr 1"
 *
 * Uses a fixed format so the subject line stays compact and scannable.
 *
 * @param {Date} date
 * @returns {string}
 */
function formatShortDate_(date) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${months[date.getMonth()]} ${date.getDate()}`;
}
