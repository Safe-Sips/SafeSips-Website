/**
 * SafeSips — Waitlist subscription (Google Apps Script)
 * =====================================================
 *
 * Receives a POST from the SafeSips website, saves the subscriber to a Google
 * Sheet, and sends a welcome email FROM the Google account that owns this
 * script.
 *
 * ── ONE-TIME SETUP ───────────────────────────────────────────────────────────
 * 1. Sign in to the Google account that will own the waitlist and create a new
 *    Google Sheet (e.g. "SafeSips Waitlist").
 * 2. In that Sheet: Extensions ▸ Apps Script. Delete the sample code and paste
 *    THIS entire file. Save.
 * 3. Deploy ▸ New deployment ▸ gear icon ▸ "Web app".
 *      - Description:        SafeSips waitlist
 *      - Execute as:         Me
 *      - Who has access:     **Anyone**  (NOT "Anyone with Google Account")
 *
 *    ⚠️ THIS MATTERS: the website POSTs anonymously (no Google sign-in). If access
 *    is set to "Anyone with Google Account" the endpoint returns HTTP 401 and every
 *    signup silently fails. It MUST be "Anyone" for the public form to work.
 *
 *    Click Deploy, then "Authorize access" and grant the Gmail + Sheets
 *    permissions (you may see an "unverified app" screen — choose
 *    Advanced ▸ Go to … ▸ Allow, since this is your own script).
 * 4. Copy the Web app URL (it ends with "/exec").
 * 5. Paste that URL into the website: open index.html and set
 *      const WAITLIST_ENDPOINT = 'https://script.google.com/macros/s/XXXX/exec';
 *    (or define window.SAFESIPS_WAITLIST_URL before the script runs).
 *
 * If you ever change this code, redeploy with Deploy ▸ Manage deployments ▸
 * edit ▸ Version: New version, so the live URL picks up the changes.
 *
 * NOTE ON LIMITS: a consumer Gmail account can send to ~100 recipients/day via
 * Apps Script. That is fine for a waitlist welcome email, but not for bulk
 * newsletters — use a dedicated email provider when you start broadcasting.
 *
 * ── ABUSE CONTROLS (why the extra code below exists) ─────────────────────────
 * The endpoint is public and anonymous, so it must defend itself:
 *  - RATE LIMIT: a global token bucket (CacheService) caps signups per window
 *    so a script kiddie can't flood the sheet or drain the Gmail send quota.
 *  - DAILY EMAIL CAP: hard ceiling on welcome emails per calendar day, kept
 *    below the Gmail quota so real signups are still recorded even when a
 *    flood eats the email budget (rows are written first, emails are best-effort).
 *  - FORMULA-INJECTION GUARD: user text starting with = + - @ is prefixed with
 *    an apostrophe before it is written to the sheet, otherwise Sheets would
 *    execute it as a live formula (e.g. =IMPORTXML exfiltration) when opened.
 *  - LOCK: a document lock serializes the dedupe-check + append so concurrent
 *    posts can't double-register the same address.
 *  - LENGTH CAPS + GENERIC ERRORS: inputs are truncated, and internal error
 *    details are logged server-side but never echoed back to the client.
 */

const SHEET_NAME = 'Waitlist';
const FROM_NAME = 'SafeSips';
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Abuse-control knobs
const RATE_LIMIT_WINDOW_SEC = 600; // 10-minute window
const RATE_LIMIT_MAX = 20;         // max accepted signups per window (global)
const DAILY_EMAIL_CAP = 80;        // welcome emails per day (Gmail quota ≈ 100)
const MAX_EMAIL_LEN = 254;
const MAX_NAME_LEN = 80;
const MAX_SOURCE_LEN = 40;
const MAX_UA_LEN = 200;

/** Handles the website's POST submission. */
function doPost(e) {
  try {
    const data = parseBody(e);
    const email = String(data.email || '').trim().toLowerCase().slice(0, MAX_EMAIL_LEN);
    const name = String(data.name || '').trim().slice(0, MAX_NAME_LEN);

    // Honeypot: bots fill hidden fields; humans never do.
    if (String(data.botcheck || data.company || '').trim() !== '') {
      return json({ ok: true }); // silently accept, store nothing
    }
    if (!EMAIL_RE.test(email)) {
      return json({ ok: false, error: 'invalid_email' });
    }

    // Global rate limit: reject floods before touching the sheet or Gmail.
    if (!rateLimitOk()) {
      return json({ ok: false, error: 'rate_limited' });
    }

    // Serialize dedupe-check + append so concurrent posts can't double-write.
    const lock = LockService.getDocumentLock();
    lock.waitLock(10000);
    let isNew;
    try {
      const sheet = getSheet();
      isNew = !emailExists(sheet, email);

      // Always persist the subscriber FIRST so a flaky/unauthorized email never
      // costs us the signup.
      if (isNew) {
        sheet.appendRow([
          new Date(),
          email, // regex-validated: cannot start with = + - @
          sheetSafe(name),
          sheetSafe(String(data.source || 'website').slice(0, MAX_SOURCE_LEN)),
          sheetSafe(String(data.ua || '').slice(0, MAX_UA_LEN))
        ]);
      }
    } finally {
      lock.releaseLock();
    }

    // The welcome email is isolated in its own try/catch: a send failure must
    // not undo the row write above, nor bubble up as a generic failure.
    let emailSent = false;
    if (isNew && underDailyEmailCap()) {
      try {
        sendWelcome(email, name);
        emailSent = true;
        bumpDailyEmailCount();
      } catch (mailErr) {
        // Details stay in the server log; the client only sees emailSent=false.
        Logger.log('sendWelcome failed for %s: %s',
          email, String(mailErr && mailErr.message ? mailErr.message : mailErr));
      }
    }

    return json({
      ok: true,
      alreadySubscribed: !isNew,
      emailSent: emailSent
    });
  } catch (err) {
    // Log the real error server-side; never echo internals to the client.
    Logger.log('doPost failed: %s', String(err && err.message ? err.message : err));
    return json({ ok: false, error: 'server_error' });
  }
}

/** Lightweight health check / friendly response for GET requests. */
function doGet() {
  return json({ ok: true, service: 'safesips-waitlist' });
}

/* ----------------------------- helpers ----------------------------------- */

function parseBody(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (_) {
      /* fall through to form params */
    }
  }
  return (e && e.parameter) || {};
}

/**
 * Global sliding-window rate limit via CacheService. Apps Script does not
 * expose the caller's IP, so this is a global bucket: it can't single out one
 * abuser, but it caps total damage (sheet flooding, Gmail quota burn).
 */
function rateLimitOk() {
  const cache = CacheService.getScriptCache();
  const key = 'signup_bucket';
  const count = Number(cache.get(key) || 0);
  if (count >= RATE_LIMIT_MAX) return false;
  cache.put(key, String(count + 1), RATE_LIMIT_WINDOW_SEC);
  return true;
}

function underDailyEmailCap() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const raw = props.getProperty('email_count_' + today);
  return Number(raw || 0) < DAILY_EMAIL_CAP;
}

function bumpDailyEmailCount() {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const key = 'email_count_' + today;
  props.setProperty(key, String(Number(props.getProperty(key) || 0) + 1));
}

/**
 * Google Sheets treats cell values starting with = + - @ as live formulas,
 * even when written via appendRow — a classic spreadsheet-injection vector
 * (e.g. a "name" of =IMPORTXML(...) runs when the owner opens the sheet).
 * Prefixing with an apostrophe forces plain text.
 */
function sheetSafe(value) {
  const s = String(value);
  return /^[=+\-@]/.test(s) ? ("'" + s) : s;
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Email', 'Name', 'Source', 'User Agent']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function emailExists(sheet, email) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const emails = sheet.getRange('B2:B' + lastRow).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (String(emails[i][0]).trim().toLowerCase() === email) return true;
  }
  return false;
}

function sendWelcome(email, name) {
  const hello = name ? ('Hi ' + name + ',') : 'Hi there,';
  const subject = "You're on the SafeSips waitlist 🩷";

  const plain = [
    hello,
    '',
    "Thanks for joining the SafeSips waitlist — your best party friend is on the way.",
    '',
    "SafeSips is a discreet, pen-like drink-screening concept with a privacy-first app.",
    "We'll email you with progress updates and early access before launch.",
    '',
    "Safety shouldn't be a privilege — it should be something you can have in your pocket.",
    '',
    '— The SafeSips team',
    '',
    "You received this because you signed up at the SafeSips website. " +
      "Reply to this email if you'd like to be removed."
  ].join('\n');

  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;color:#001F3F;">' +
      '<div style="background:#001F3F;padding:24px 28px;border-radius:16px 16px 0 0;">' +
        '<h1 style="margin:0;font-size:26px;letter-spacing:1px;color:#fff;">Safe<span style="color:#FF1493;">Sips</span></h1>' +
        '<p style="margin:6px 0 0;color:rgba(255,255,255,.7);font-size:13px;">Your best party friend</p>' +
      '</div>' +
      '<div style="background:#f7f3fb;padding:28px;border-radius:0 0 16px 16px;">' +
        '<p style="font-size:16px;margin:0 0 14px;">' + escapeHtml(hello) + '</p>' +
        '<p style="font-size:15px;line-height:1.6;margin:0 0 14px;">' +
          "Thanks for joining the <strong>SafeSips waitlist</strong> — your best party friend is on the way." +
        '</p>' +
        '<p style="font-size:15px;line-height:1.6;margin:0 0 14px;">' +
          "SafeSips is a discreet, pen-like drink-screening concept with a privacy-first app. " +
          "We'll email you with progress updates and early access before launch." +
        '</p>' +
        '<p style="font-size:14px;line-height:1.6;font-style:italic;color:#5a5a72;margin:18px 0 0;">' +
          "Safety shouldn't be a privilege — it should be something you can have in your pocket." +
        '</p>' +
        '<p style="font-size:14px;margin:18px 0 0;">— The SafeSips team</p>' +
        '<p style="font-size:11px;color:#9a9ab0;margin:22px 0 0;line-height:1.5;">' +
          "You received this because you signed up at the SafeSips website. " +
          "Reply to this email if you'd like to be removed." +
        '</p>' +
      '</div>' +
    '</div>';

  GmailApp.sendEmail(email, subject, plain, { name: FROM_NAME, htmlBody: html });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
