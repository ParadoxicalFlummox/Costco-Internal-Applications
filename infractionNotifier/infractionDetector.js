/**
 * infractionDetector.js — Rolling-window CN detection logic.
 * VERSION: 1.0.0
 *
 * This file owns the logic for determining whether a sequence of infraction
 * events is severe enough to trigger a Counseling Notice (CN).
 *
 * DETECTION STRATEGY:
 *   Each set of events is evaluated using a sliding window algorithm.
 *   The window is defined by a threshold count and a number of days. As the
 *   pointer advances through the chronologically sorted event list, events
 *   older than (windowDays - 1) days before the current event are dropped
 *   from the left of the queue. When the queue length reaches the threshold,
 *   a CN proposal is emitted.
 *
 *   To avoid flooding payroll with one email per event (e.g. a 5-tardy window
 *   would otherwise emit 3 proposals: at events 3, 4, and 5), proposals are
 *   coalesced: only one proposal is emitted per unique window-end date.
 *
 * TWO MODES:
 *   1. RULE-BASED (default): Each code in CODE_RULES is evaluated independently
 *      using its own threshold and window length. This correctly enforces that
 *      3 tardies in 30 days and 1 no-show in 30 days are separate triggers.
 *
 *   2. GLOBAL (fallback): All infraction codes are pooled together and evaluated
 *      against the global THRESHOLD_COUNT / WINDOW_DAYS. This is used when no
 *      CODE_RULES are configured.
 *
 * IDEMPOTENCY:
 *   Each CN proposal carries a CN_Key (employee|rule|windowStart|windowEnd)
 *   and an EventsHash (SHA-1 of the sorted event list). The log sheet check in
 *   infractionEngine.js uses these to skip proposals that were already issued
 *   with the same evidence, preventing duplicate emails on re-runs.
 *
 * OUTPUT SHAPE — CNProposal:
 *   {
 *     cnKey:        string   — Deduplication key
 *     eventsHash:   string   — SHA-1 of sorted event dates+codes
 *     employeeId:   string
 *     employeeName: string
 *     department:   string
 *     windowStart:  Date     — Date of the first event in the triggering window
 *     windowEnd:    Date     — Date of the last event (the one that crossed the threshold)
 *     count:        number   — Number of events in the window
 *     events:       Array    — Snapshot of events: [{ date, code, a1 }]
 *     rule:         string   — Code that triggered this CN, or "GLOBAL"
 *     windowDays:   number   — Window length used for this proposal
 *   }
 */

/**
 * @typedef {{
 *   cnKey:        string,
 *   eventsHash:   string,
 *   employeeId:   string,
 *   employeeName: string,
 *   department:   string,
 *   windowStart:  Date,
 *   windowEnd:    Date,
 *   count:        number,
 *   events:       Array<{ date: Date, code: string, a1: string }>,
 *   rule:         string,
 *   windowDays:   number
 * }} CNProposal
 */


// ---------------------------------------------------------------------------
// Public Entry Point
// ---------------------------------------------------------------------------

/**
 * Builds CN proposals for a single employee's infraction event list.
 *
 * If CODE_RULES is configured and has at least one entry, each code is
 * evaluated independently with its own threshold and window. Otherwise,
 * all events are pooled and evaluated against the global defaults.
 *
 * Returns an empty array if no threshold is crossed.
 *
 * @param {CalendarEvent[]} infractionEvents — Only infraction events (isInfraction=true),
 *   pre-sorted chronologically. Caller is responsible for filtering out ignored events.
 * @param {string} timeZone — Script time zone for date formatting.
 * @param {{ employeeId, employeeName, department }} ctx — Employee identity.
 * @returns {CNProposal[]}
 */
function buildAllCNProposals_(infractionEvents, timeZone, ctx) {
  if (!infractionEvents || infractionEvents.length === 0) return [];

  const hasCodeRules = CODE_RULES && Object.keys(CODE_RULES).length > 0; // config.js

  if (hasCodeRules) {
    return buildRuleBasedCNs_(infractionEvents, timeZone, ctx);
  }

  return buildCNProposals_(
    infractionEvents,
    THRESHOLD_COUNT, // config.js
    WINDOW_DAYS,     // config.js
    timeZone,
    ctx,
    'GLOBAL'
  );
}


// ---------------------------------------------------------------------------
// Rule-Based Detection
// ---------------------------------------------------------------------------

/**
 * Evaluates each code in CODE_RULES independently and returns all proposals.
 *
 * For each code that has a rule, the infraction event list is filtered to just
 * that code and run through the sliding window algorithm with the code's own
 * threshold and windowDays. This means a TD rule and an NS rule can both fire
 * independently for the same employee in the same time period.
 *
 * @param {CalendarEvent[]} events — All infraction events, sorted chronologically.
 * @param {string} timeZone
 * @param {{ employeeId, employeeName, department }} ctx
 * @returns {CNProposal[]}
 */
function buildRuleBasedCNs_(events, timeZone, ctx) {
  const proposals = [];
  const rules = CODE_RULES || {}; // config.js

  Object.keys(rules).forEach(code => {
    const rule = rules[code];
    const codeEvents = events
      .filter(e => e.code === code)
      .sort((a, b) => a.date.getTime() - b.date.getTime());

    if (codeEvents.length === 0) return;

    const codeProposals = buildCNProposals_(
      codeEvents,
      rule.threshold || THRESHOLD_COUNT,
      rule.windowDays || WINDOW_DAYS,
      timeZone,
      ctx,
      code
    );

    proposals.push(...codeProposals);
  });

  return proposals;
}


// ---------------------------------------------------------------------------
// Sliding Window Algorithm
// ---------------------------------------------------------------------------

/**
 * Sliding-window CN detection for a pre-filtered, sorted event list.
 *
 * The algorithm maintains a queue of events. As the cursor advances through
 * the sorted list, events older than (windowDays - 1) days before the current
 * event are dropped from the front of the queue. When the queue length reaches
 * the threshold, a proposal is emitted for the window [queue[0].date, cursor.date].
 *
 * Coalescing: multiple events on the same end date (e.g. two tardies logged
 * on the same day) would emit duplicate proposals with the same windowEnd.
 * The lastEndKey guard suppresses these duplicates.
 *
 * @param {CalendarEvent[]} events     — Infraction events for a single code (or all),
 *   sorted chronologically oldest-first.
 * @param {number} threshold           — Minimum events in window to trigger a CN.
 * @param {number} windowDays          — Rolling window length in days.
 * @param {string} timeZone            — For date key formatting.
 * @param {{ employeeId, employeeName, department }} ctx
 * @param {string} ruleLabel           — e.g. "TD", "NS", "GLOBAL"
 * @returns {CNProposal[]}
 */
function buildCNProposals_(events, threshold, windowDays, timeZone, ctx, ruleLabel) {
  const proposals = [];
  if (!events || events.length === 0) return proposals;

  const msPerDay = 24 * 60 * 60 * 1000;
  const windowSpan = (windowDays - 1) * msPerDay; // inclusive: [end - N-1 days, end]

  const queue = [];
  let lastEndKey = null;

  for (let i = 0; i < events.length; i++) {
    const current = events[i];
    const endMs = startOfDayMs_(current.date);

    // Evict events from the front of the queue that fall outside the window
    while (queue.length > 0 && (endMs - startOfDayMs_(queue[0].date)) > windowSpan) {
      queue.shift();
    }
    queue.push(current);

    if (queue.length >= threshold) {
      const windowStart = queue[0].date;
      const windowEnd = current.date;
      const endKey = formatDateYmd_(windowEnd, timeZone);

      // Coalesce: skip if we already emitted a proposal for this same end date
      if (endKey === lastEndKey) continue;

      const eventsSnapshot = queue.map(e => ({
        date: new Date(e.date),
        code: e.code,
        a1: e.a1,
      }));

      const eventsHash = computeEventsHash_(eventsSnapshot, timeZone);

      const cnKey = [
        ctx.employeeId || ctx.employeeName || 'UNKNOWN',
        `RULE:${ruleLabel}`,
        formatDateYmd_(windowStart, timeZone),
        formatDateYmd_(windowEnd, timeZone),
      ].join('|');

      proposals.push({
        cnKey: cnKey,
        eventsHash: eventsHash,
        employeeId: ctx.employeeId || '',
        employeeName: ctx.employeeName || '',
        department: ctx.department || '',
        windowStart: windowStart,
        windowEnd: windowEnd,
        count: queue.length,
        events: eventsSnapshot,
        rule: ruleLabel,
        windowDays: windowDays,
      });

      lastEndKey = endKey;
    }
  }

  return proposals;
}


// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

/**
 * Returns the start-of-day millisecond timestamp for a Date, stripping the
 * time component. Used to compute day-level differences without DST drift.
 *
 * @param {Date} date
 * @returns {number}
 */
function startOfDayMs_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}

/**
 * Formats a Date as "yyyy-MM-dd" in the given time zone.
 * Used as a consistent key for deduplication.
 *
 * @param {Date}   date
 * @param {string} timeZone
 * @returns {string} e.g. "2026-04-13"
 */
function formatDateYmd_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}

/**
 * Computes a SHA-1 hex digest of the event list, used as the EventsHash
 * for idempotent deduplication.
 *
 * The hash is computed over the sorted "date|code" strings so that the same
 * set of events always produces the same hash regardless of array order.
 * If a new infraction is added to a window that was already logged, the hash
 * changes and a new CN is issued.
 *
 * @param {Array<{ date: Date, code: string }>} events
 * @param {string} timeZone
 * @returns {string} 40-character hex SHA-1 string.
 */
function computeEventsHash_(events, timeZone) {
  const sortedKeys = events
    .map(e => `${formatDateYmd_(e.date, timeZone)}|${e.code}`)
    .sort();
  const rawBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    JSON.stringify(sortedKeys)
  );
  return rawBytes
    .map(b => {
      const hex = (b < 0 ? b + 256 : b).toString(16);
      return hex.length === 1 ? '0' + hex : hex;
    })
    .join('');
}

