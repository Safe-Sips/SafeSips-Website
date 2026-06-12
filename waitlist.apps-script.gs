/**
 * SafeSips — Waitlist subscription (Google Apps Script)
 * =====================================================
 *
 * Receives a POST from the SafeSips website, saves the subscriber to a Google
 * Sheet, and sends a welcome email FROM the Google account that owns this
 * script (jamld2135@gmail.com).
 *
 * ── ONE-TIME SETUP ───────────────────────────────────────────────────────────
 * 1. Sign in to Google as jamld2135@gmail.com and create a new Google Sheet
 *    (e.g. "SafeSips Waitlist").
 * 2. In that Sheet: Extensions ▸ Apps Script. Delete the sample code and paste
 *    THIS entire file. Save.
 * 3. Deploy ▸ New deployment ▸ gear icon ▸ "Web app".
 *      - Description:        SafeSips waitlist
 *      - Execute as:         Me (jamld2135@gmail.com)
 *      - Who has access:     Anyone
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
 */

const SHEET_NAME = 'Waitlist';
const FROM_NAME = 'SafeSips';
const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/** Handles the website's POST submission. */
function doPost(e) {
  try {
    const data = parseBody(e);
    const email = String(data.email || '').trim().toLowerCase();
    const name = String(data.name || '').trim();

    // Honeypot: bots fill hidden fields; humans never do.
    if (String(data.company || '').trim() !== '') {
      return json({ ok: true }); // silently accept, store nothing
    }
    if (!EMAIL_RE.test(email)) {
      return json({ ok: false, error: 'invalid_email' });
    }

    const sheet = getSheet();
    const isNew = !emailExists(sheet, email);
    if (isNew) {
      sheet.appendRow([
        new Date(),
        email,
        name,
        String(data.source || 'website'),
        String(data.ua || '')
      ]);
      sendWelcome(email, name);
    }

    return json({ ok: true, alreadySubscribed: !isNew });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
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
  const emails = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
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
    .replace(/"/g, '&quot;');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
