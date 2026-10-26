/**
 * Tenant Information Form — Google Apps Script
 * Saves each form submission as a new row in the "Responses" sheet.
 * Property details (P-Ref, B-No, Agent fee) come from the "Properties" sheet.
 */

// ============================================================
//  Your web page address (the GitHub Pages link), ending with /
//  رابط صفحتك على GitHub Pages، وينتهي بـ /
// ============================================================
const PAGE_URL = "https://ali-kalash.github.io/tenant-form/";
// ============================================================

const PROPERTIES_SHEET = "Properties";
const RESPONSES_SHEET = "Responses";
const PROPERTIES_HEADERS = ["Key", "P-Ref", "B-No", "Agent fee (%)", "Link"];
const RESPONSES_HEADERS = [
  "Timestamp", "Submission ID", "Key", "P-Ref", "B-No", "Agent fee (%)",
  "Full Name", "Email address", "Nationality", "Work at", "In position of",
  "A family of (Number)", "Rental amount agreed upon", "Payment method",
  "Security payment", "Contract start date", "Potential move-in date",
  "Tenant's Note & Feedback", "Confirmed", "Tenant Signature", "Date", "Language"
];
// Form fields in the same order as the columns after "Agent fee (%)"
const FIELDS = [
  "full_name", "email", "nationality", "work_at", "position", "family_size",
  "rental_amount", "payment_method", "security_payment",
  "contract_start_date", "potential_move_in_date", "tenant_note"
];
const REQUIRED = [
  "full_name", "nationality", "work_at", "position", "family_size",
  "rental_amount", "payment_method", "security_payment",
  "contract_start_date", "potential_move_in_date", "tenant_signature", "tenant_date"
];
const DATE_FIELDS = ["contract_start_date", "potential_move_in_date", "tenant_date"];
const KEY_PATTERN = /^[A-Za-z0-9]{12,32}$/;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const DATE_PATTERN = /^\d{2}\/\d{2}\/\d{4}$/;
const ID_PATTERN = /^[A-Za-z0-9-]{16,64}$/;
const KEY_LENGTH = 12;
const KEY_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
const MAX_LENGTH = 1000;
const FEE_PATTERN = /^\d+(\.\d+)?$/;
const HONEYPOT = "mep_xq_alt7";   // hidden field, must match index.html

// ---------- Menu inside the spreadsheet ----------

function onOpen() {
  setupSheets();
  SpreadsheetApp.getUi()
    .createMenu("Tenant Form")
    .addItem("Generate link", "generateLinks")
    .addToUi();
}

/** Creates the two sheets with their header rows if they don't exist yet. */
function setupSheets() {
  getSheet_(PROPERTIES_SHEET, PROPERTIES_HEADERS);
  getSheet_(RESPONSES_SHEET, RESPONSES_HEADERS);
}

/**
 * For every complete row in "Properties" (P-Ref, B-No and a numeric Agent fee)
 * that has no Key yet: creates a random 12-character key and writes the link.
 */
function generateLinks() {
  const sheet = getSheet_(PROPERTIES_SHEET, PROPERTIES_HEADERS);
  const lastRow = sheet.getLastRow();
  let created = 0;
  const incomplete = [];
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 1, lastRow - 1, 5);
    const rows = range.getDisplayValues();
    const used = {};
    rows.forEach(function (r) { if (r[0]) used[r[0]] = true; });
    rows.forEach(function (r, i) {
      const filled = r.slice(1, 4).some(function (x) { return String(x).trim() !== ""; });
      if (!isCompleteProperty_(r[1], r[2], r[3])) {
        if (filled || r[0]) incomplete.push(i + 2);
        return;
      }
      if (!r[0]) {
        let key;
        do { key = randomKey_(); } while (used[key]);
        used[key] = true;
        sheet.getRange(i + 2, 1).setNumberFormat("@").setValue(key);
        sheet.getRange(i + 2, 5).setValue(PAGE_URL + "?k=" + key);
        created++;
      } else if (r[0] && !r[4]) {
        sheet.getRange(i + 2, 5).setValue(PAGE_URL + "?k=" + r[0]);
      }
    });
  }
  let message = created ? created + " new link(s) created — copy them from the Link column." :
    "No new links. Fill P-Ref, B-No and Agent fee in a new row first.";
  if (incomplete.length) {
    message += " Skipped row(s) " + incomplete.join(", ") +
      ": P-Ref, B-No and Agent fee must all be filled, and Agent fee must be a number (e.g. 2.5).";
  }
  SpreadsheetApp.getActive().toast(message, "Tenant Form", 10);
}

/** Random key from secure random UUIDs, letters and digits only. */
function randomKey_() {
  let key = "";
  while (key.length < KEY_LENGTH) {
    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
    for (let i = 0; i < bytes.length && key.length < KEY_LENGTH; i++) {
      const b = bytes[i] & 0xff;
      if (b < 248) key += KEY_ALPHABET.charAt(b % 62); // 248 = 4 × 62, keeps it even
    }
  }
  return key;
}

// ---------- Web App ----------

/** Page load: returns P-Ref, B-No and Agent fee for one key, nothing else. */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action !== "property") return json_({ ok: false });
  const property = findProperty_(p.k);
  if (!property) return json_({ ok: false });
  return json_({ ok: true, pref: property.pref, bno: property.bno, fee: property.fee });
}

/**
 * Form submission: adds one row.
 * Replies only "success", "failure", or "invalid_link" (unknown or incomplete property).
 */
function doPost(e) {
  try {
    return json_({ result: saveSubmission_(e) });
  } catch (err) {
    console.error(err);
    return json_({ result: "failure" });
  }
}

function saveSubmission_(e) {
  const FAIL = "failure";
  if (!e || !e.postData || !e.postData.contents) return FAIL;
  if (e.postData.contents.length > 20000) return FAIL;
  const data = JSON.parse(e.postData.contents);
  if (!data || data.action !== "submit") return FAIL;
  if (data[HONEYPOT]) return FAIL;                      // hidden field filled = bot
  if (!ID_PATTERN.test(String(data.id || ""))) return FAIL;

  // Clean every value: text only, trimmed, length-capped
  const v = {};
  FIELDS.concat(["tenant_signature", "tenant_date"]).forEach(function (f) {
    v[f] = typeof data[f] === "string" ? data[f].trim().slice(0, MAX_LENGTH) : "";
  });

  v.family_size = toWesternDigits_(v.family_size);

  // Check again — never trust the browser alone
  for (let i = 0; i < REQUIRED.length; i++) if (!v[REQUIRED[i]]) return FAIL;
  if (v.email && !EMAIL_PATTERN.test(v.email)) return FAIL;
  if (!/^[1-9][0-9]?$/.test(v.family_size)) return FAIL;
  for (let i = 0; i < DATE_FIELDS.length; i++) if (!parseDate_(v[DATE_FIELDS[i]])) return FAIL;
  if (isFutureDate_(parseDate_(v.tenant_date))) return FAIL;
  if (data.confirmed !== true) return FAIL;

  // Property values always come from the sheet, never from the browser
  const property = findProperty_(data.k);
  if (!property) return "invalid_link";

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_(RESPONSES_SHEET, RESPONSES_HEADERS);
    // Same submission ID already saved (double tap or retry) → don't add it again
    if (sheet.getLastRow() >= 2) {
      const found = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1)
        .createTextFinder(data.id).matchEntireCell(true).matchCase(true).findNext();
      if (found) return "success";
    }
    const row = [
      data.id, data.k, property.pref, property.bno, property.fee
    ].concat(FIELDS.map(function (f) { return v[f]; }))
      .concat(["Yes", v.tenant_signature, v.tenant_date, data.lang === "ar" ? "Arabic" : "English"])
      .map(asText_);
    const r = sheet.getLastRow() + 1;
    sheet.getRange(r, 1).setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(r, 2, 1, row.length).setNumberFormat("@").setValues([row]);
    SpreadsheetApp.flush();
    return "success";
  } finally {
    lock.releaseLock();
  }
}

// ---------- Helpers ----------

function findProperty_(key) {
  key = String(key || "");
  if (!KEY_PATTERN.test(key)) return null;
  const sheet = getSheet_(PROPERTIES_SHEET, PROPERTIES_HEADERS);
  if (sheet.getLastRow() < 2) return null;
  const cell = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(key).matchEntireCell(true).matchCase(true).findNext();
  if (!cell) return null;
  const r = sheet.getRange(cell.getRow(), 1, 1, 4).getDisplayValues()[0];
  if (!isCompleteProperty_(r[1], r[2], r[3])) return null;
  return { pref: r[1].trim(), bno: r[2].trim(), fee: r[3].trim() };
}

/** P-Ref, B-No and Agent fee all filled, and Agent fee is a plain number like 2.5 */
function isCompleteProperty_(pref, bno, fee) {
  return String(pref).trim() !== "" && String(bno).trim() !== "" &&
    FEE_PATTERN.test(String(fee).trim());
}

/** Arabic-Indic digits (٠-٩ and ۰-۹) → 0-9 */
function toWesternDigits_(s) {
  return String(s).replace(/[\u0660-\u0669\u06F0-\u06F9]/g, function (d) {
    return String(d.charCodeAt(0) & 0xf);
  });
}

/** "DD/MM/YYYY" → {y, m, d} if it is a real calendar date, otherwise null */
function parseDate_(s) {
  if (!DATE_PATTERN.test(s)) return null;
  const d = Number(s.slice(0, 2)), m = Number(s.slice(3, 5)), y = Number(s.slice(6));
  if (y < 1900 || y > 2100 || m < 1 || m > 12 || d < 1) return null;
  const daysInMonth = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return d <= daysInMonth ? { y: y, m: m, d: d } : null;
}

/**
 * True if the date is after today. "Today" is taken in the spreadsheet's time zone,
 * plus one day of leeway for customers whose phone is in a time zone ahead of it.
 */
function isFutureDate_(date) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const tomorrow = new Date(Date.now() + 24 * 60 * 60 * 1000);
  const limit = Number(Utilities.formatDate(tomorrow, tz, "yyyyMMdd"));
  return date.y * 10000 + date.m * 100 + date.d > limit;
}

/** Values starting with = + - @ are stored as plain text, never as a formula. */
function asText_(value) {
  const s = String(value == null ? "" : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#11502C").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
