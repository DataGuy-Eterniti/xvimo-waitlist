/**
 * Xvimo early-access waitlist — Google Apps Script backend.
 * Saves signups to the Google Sheet this script is attached to.
 */

const SHEET_NAME = 'Waitlist';
const HEADERS = ['Timestamp', 'Name', 'Email', 'Country', 'I am a', 'WhatsApp (optional)',
                 'How did you hear about Xvimo?', 'Consent', 'Source'];
const SEND_CONFIRMATION_EMAIL = true;   // set to false to skip the thank-you email

/** Serves the signup page, or the waitlist count as JSON when called with ?action=count. */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'count') {
    return json_({ count: getCount() });
  }
  const page = HtmlService.createTemplateFromFile('Index');
  page.source = (e && e.parameter && (e.parameter.utm_source || e.parameter.ref)) || 'direct';
  return page.evaluate()
    .setTitle('Xvimo — early access')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Receives signups from the externally hosted page (GitHub Pages etc.).
 * The page sends JSON as text/plain, which avoids browser CORS pre-flight checks.
 */
function doPost(e) {
  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    data = (e && e.parameter) || {};
  }
  try {
    return json_(submitSignup(data));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, message: 'Something went wrong. Please try again in a moment.' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Returns the sheet, creating it with headers on first use. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Number of people on the waitlist (shown on the page). */
function getCount() {
  return Math.max(getSheet_().getLastRow() - 1, 0);
}

/** Called from the page. Validates, de-duplicates and saves one signup. */
function submitSignup(form) {
  // Honeypot: real people never fill the hidden "website" field.
  if (form.website) return { ok: true, message: "You're on the list!" };

  const name = clean_(form.name, 80);
  const email = clean_(form.email, 120).toLowerCase();
  const country = clean_(form.country, 60);
  const role = clean_(form.role, 60);
  const whatsapp = clean_(form.whatsapp, 25).replace(/[^\d+]/g, '');
  const heard = clean_(form.heard, 120);
  const source = clean_(form.source, 60) || 'direct';

  if (!name) return { ok: false, message: 'Please enter your name.' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return { ok: false, message: 'Please enter a valid email address.' };
  if (form.consent !== 'yes') return { ok: false, message: 'Please agree to be contacted about early access.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const last = sheet.getLastRow();
    if (last > 1) {
      const emails = sheet.getRange(2, 3, last - 1, 1).getValues().flat().map(String);
      if (emails.indexOf(email) !== -1) {
        return { ok: true, duplicate: true, message: "This email is already registered — we'll be in touch." };
      }
    }
    sheet.appendRow([new Date(), name, email, country, role, whatsapp, heard, 'Yes', source]);
  } finally {
    lock.releaseLock();
  }

  if (SEND_CONFIRMATION_EMAIL) {
    try {
      MailApp.sendEmail({
        to: email,
        subject: "You're on the Xvimo early-access list",
        htmlBody:
          '<p>Hi ' + escape_(name.split(' ')[0]) + ',</p>' +
          '<p>Thanks for joining the Xvimo waitlist. Xvimo checks suspicious messages on WhatsApp — ' +
          'investment offers, bank alerts, job invites — and tells you if they look like a scam, with the evidence.</p>' +
          "<p>We'll email you as soon as early access opens.</p>" +
          '<p>— The Xvimo team<br><i>Scam check before you pay.</i></p>' +
          '<p style="color:#888;font-size:12px">You received this because you signed up at the Xvimo waitlist. ' +
          'Reply to this email to be removed.</p>'
      });
    } catch (err) {
      console.warn('Confirmation email failed: ' + err);
    }
  }
  return { ok: true, message: SEND_CONFIRMATION_EMAIL ? 'Check your inbox for a confirmation email.' : "We'll email you when early access opens." };
}

function clean_(value, max) {
  return String(value || '').replace(/[<>]/g, '').trim().slice(0, max);
}

function escape_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
