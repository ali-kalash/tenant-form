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
//  Google Drive folder for the owner PDF copies (the code after /folders/ in its link)
//  رقم مجلد Google Drive لنسخ المالك (ما بعد ‎/folders/‎ في رابط المجلد)
// ============================================================
const OWNER_PDF_FOLDER_ID = "1RO0T25iDJMrtDsFWXNoeLgCsWh51e4yH";
// ============================================================

const PROPERTIES_SHEET = "Properties";
const RESPONSES_SHEET = "Responses";
const PROPERTIES_HEADERS = ["Key", "P-Ref", "B-No", "Agent fee (%)", "Link"];
const RESPONSES_HEADERS = [
  "Timestamp", "Submission ID", "Key", "P-Ref", "B-No", "Agent fee (%)",
  "Full Name", "Email address", "Nationality", "Work at", "In position of",
  "A family of (Number)", "Rental amount agreed upon", "Payment method",
  "Security payment", "Contract start date", "Potential move-in date",
  "Tenant's Note & Feedback", "Confirmed", "Tenant Signature", "Date", "Language",
  "Reference", "Owner PDF"
];
const OWNER_PDF_COLUMN = RESPONSES_HEADERS.length;   // last column
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
const TIME_ZONE = "Asia/Muscat";

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
 * Form submission: adds one row, then saves the owner PDF copy.
 * Replies only "failure", "invalid_link" (unknown or incomplete property),
 * or "success" with the submission time and reference number.
 */
function doPost(e) {
  let saved;
  try {
    saved = saveSubmission_(e);
  } catch (err) {
    console.error(err);
    return json_({ result: "failure" });
  }
  if (typeof saved === "string") return json_({ result: saved });
  // The row is saved. A problem with the owner PDF must never fail the submission.
  if (saved.owner) saveOwnerPdfSafely_(saved.owner);
  return json_({ result: "success", time: saved.time, ref: saved.ref });
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

  const ref = referenceOf_(data.id);
  const lang = data.lang === "ar" ? "ar" : "en";
  const now = new Date();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_(RESPONSES_SHEET, RESPONSES_HEADERS);
    // Same submission ID already saved (double tap or retry) → don't add it again
    if (sheet.getLastRow() >= 2) {
      const found = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1)
        .createTextFinder(data.id).matchEntireCell(true).matchCase(true).findNext();
      if (found) {
        const savedAt = sheet.getRange(found.getRow(), 1).getValue();
        return { time: formatTime_(savedAt instanceof Date ? savedAt : new Date()), ref: ref };
      }
    }
    const row = [
      data.id, data.k, property.pref, property.bno, property.fee
    ].concat(FIELDS.map(function (f) { return v[f]; }))
      .concat(["Yes", v.tenant_signature, v.tenant_date, lang === "ar" ? "Arabic" : "English", ref])
      .map(asText_);
    const r = sheet.getLastRow() + 1;
    sheet.getRange(r, 1).setValue(now).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(r, 2, 1, row.length).setNumberFormat("@").setValues([row]);
    SpreadsheetApp.flush();
    return {
      time: formatTime_(now),
      ref: ref,
      // Everything the owner copy needs — note: no agent fee
      owner: { id: data.id, lang: lang, ref: ref, time: formatTime_(now),
        day: Utilities.formatDate(now, TIME_ZONE, "yyyy-MM-dd"),
        pref: property.pref, bno: property.bno, values: v }
    };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Owner PDF (saved to Drive, never shared, no agent fee) ----------

const OWNER_TEXT = {
  en: {
    title: "Tenant Information Form", tagline: "Real Estate Brokerage · Muscat",
    pref: "P- Ref:", bno: "B- No.:", ref: "Reference No.:", submitted: "Submitted:",
    details: "Tenant Details:", item: "Item", detailsCol: "Details",
    full_name: "Full Name", email: "Email address", nationality: "Nationality",
    work_at: "Work at", position: "In position of", family_size: "A family of (Number)",
    rental_amount: "Rental amount agreed upon", payment_method: "Payment method",
    security_payment: "Security payment", contract_start_date: "Contract start date",
    potential_move_in_date: "Potential move-in date",
    note: "Tenant’s Note & Feedback:",
    confirm: "The tenant confirmed the accuracy of the above data and agreed to it.",
    signature: "Tenant Signature", signatureHint: "Typed full name — counts as the tenant’s signature.",
    date: "Date",
    footer: "Bousher – Muscat – Sultanate of Oman - P.O.Box: 434 - PC: 320 | alikalach@outlook.com"
  },
  ar: {
    title: "نموذج معلومات المستأجر", tagline: "وساطة عقارية · مسقط",
    pref: "مرجع العقار:", bno: "رقم المبنى:", ref: "الرقم المرجعي:", submitted: "تاريخ الإرسال:",
    details: "بيانات المستأجر:", item: "البند", detailsCol: "التفاصيل",
    full_name: "الاسم الكامل", email: "البريد الإلكتروني", nationality: "الجنسية",
    work_at: "جهة العمل", position: "المسمى الوظيفي", family_size: "عدد أفراد العائلة",
    rental_amount: "مبلغ الإيجار المتفق عليه", payment_method: "طريقة الدفع",
    security_payment: "مبلغ التأمين", contract_start_date: "تاريخ بدء العقد",
    potential_move_in_date: "تاريخ الانتقال المتوقع",
    note: "ملاحظات المستأجر:",
    confirm: "أقرّ المستأجر بصحة البيانات أعلاه ووافق عليها.",
    signature: "توقيع المستأجر", signatureHint: "الاسم الكامل مكتوباً، ويُعتبر توقيعاً من المستأجر.",
    date: "التاريخ",
    footer: "بوشر – مسقط – سلطنة عُمان - ص.ب: 434 - الرمز البريدي: 320 | alikalach@outlook.com"
  }
};
const OWNER_ROWS = [
  "full_name", "email", "nationality", "work_at", "position", "family_size",
  "rental_amount", "payment_method", "security_payment",
  "contract_start_date", "potential_move_in_date"
];

/** Creates the owner PDF and writes its link (or "failed") in the "Owner PDF" column. */
function saveOwnerPdfSafely_(owner) {
  let cellText = "failed";
  try {
    cellText = createOwnerPdf_(owner).getUrl();
  } catch (err) {
    console.error("Owner PDF failed for " + owner.ref + ": " + err);
  }
  try {
    const sheet = getSheet_(RESPONSES_SHEET, RESPONSES_HEADERS);
    const found = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 1), 1)
      .createTextFinder(owner.id).matchEntireCell(true).matchCase(true).findNext();
    if (found) sheet.getRange(found.getRow(), OWNER_PDF_COLUMN).setValue(cellText);
  } catch (err) {
    console.error("Could not write Owner PDF cell for " + owner.ref + ": " + err);
  }
}

/** Builds the PDF and saves it in the owner folder. Returns the Drive file. */
function createOwnerPdf_(owner) {
  const name = [owner.pref, owner.ref, owner.day].join(" - ").replace(/[\\/:*?"<>|]/g, "-") + ".pdf";
  const pdf = Utilities.newBlob(ownerPdfHtml_(owner), "text/html", "owner.html")
    .getAs("application/pdf").setName(name);
  return DriveApp.getFolderById(OWNER_PDF_FOLDER_ID).createFile(pdf);
}

/** The owner copy as HTML (Google turns it into a PDF). Contains no agent fee. */
function ownerPdfHtml_(o) {
  const T = OWNER_TEXT[o.lang] || OWNER_TEXT.en;
  const rtl = o.lang === "ar";
  const v = o.values;
  const start = rtl ? "right" : "left", end = rtl ? "left" : "right";
  const serif = rtl ? "'Amiri', 'Times New Roman', serif" : "'EB Garamond', Garamond, Georgia, serif";
  const sans = rtl ? "'Noto Sans Arabic', Arial, sans-serif" : "'Barlow', Helvetica, Arial, sans-serif";
  const caps = rtl ? "" : "text-transform:uppercase;letter-spacing:0.1em;";
  const green = "#11502C", gold = "#C5A16F";
  const val = function (x) { return esc_(x || "—"); };

  const rows = OWNER_ROWS.map(function (f, i) {
    const bg = i % 2 ? "background:#EAEAEA;" : "";
    return '<tr><td style="' + bg + 'width:44%;border:1px solid #D5D5D5;padding:6px 10px;color:#333;text-align:' + start + '">' +
      esc_(T[f]) + '</td><td dir="auto" style="' + bg + 'border:1px solid #D5D5D5;padding:6px 10px;color:' + green +
      ';text-align:' + start + '">' + val(v[f]) + "</td></tr>";
  }).join("");

  return '<!DOCTYPE html><html lang="' + o.lang + '" dir="' + (rtl ? "rtl" : "ltr") + '"><head><meta charset="utf-8">' +
    "<style>@page{size:A4;margin:14mm 15mm}body{margin:0;font-family:" + sans + ";font-size:10.5pt;color:#2A2A2A;direction:" +
    (rtl ? "rtl" : "ltr") + "}td{vertical-align:top}</style></head><body>" +
    // Header: company name and logo
    '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<td dir="ltr" style="text-align:' + start + ';font-family:\'EB Garamond\',Georgia,serif;color:' + green +
    ';font-size:19pt;line-height:1.12">Middle East Properties Global<br><span style="font-size:13pt">S.P.C</span>' +
    '<div style="font-family:' + sans + ';font-size:8pt;color:#7C7C7C;margin-top:5px;' + caps + '">' + esc_(T.tagline) + "</div></td>" +
    '<td style="text-align:' + end + ';width:120px"><img src="data:image/png;base64,' + LOGO_PNG_BASE64 +
    '" style="height:68px" alt=""></td></tr></table>' +
    '<div style="border-top:1px solid ' + gold + ';margin:10px 0 14px"></div>' +
    // Title and references
    '<div style="font-family:' + serif + ";font-size:19pt;color:" + green + ";" + caps + 'text-align:' + start + '">' + esc_(T.title) + "</div>" +
    '<table style="width:100%;border-collapse:collapse;margin:10px 0 16px;font-size:9.5pt;color:#555"><tr>' +
    '<td style="text-align:' + start + '">' + esc_(T.pref) + ' <b dir="ltr" style="color:' + green + '">' + val(o.pref) + "</b>" +
    "&nbsp;&nbsp;&nbsp;&nbsp;" + esc_(T.bno) + ' <b dir="ltr" style="color:' + green + '">' + val(o.bno) + "</b></td>" +
    '<td style="text-align:' + end + '">' + esc_(T.ref) + ' <b dir="ltr" style="color:' + green + '">' + esc_(o.ref) + "</b>" +
    "&nbsp;&nbsp;&nbsp;&nbsp;" + esc_(T.submitted) + ' <b dir="ltr" style="color:' + green + '">' + esc_(o.time) + "</b></td></tr></table>" +
    // Details table
    '<div style="font-family:' + serif + ";font-size:13pt;color:" + green + ';margin-bottom:6px;text-align:' + start + '">' + esc_(T.details) + "</div>" +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
    '<th style="background:' + green + ";color:#fff;font-weight:500;font-size:8.5pt;padding:7px 10px;text-align:" + start + ";" + caps + '">' + esc_(T.item) + "</th>" +
    '<th style="background:' + green + ";color:#fff;font-weight:500;font-size:8.5pt;padding:7px 10px;text-align:" + start + ";" + caps + '">' + esc_(T.detailsCol) + "</th></tr>" +
    rows + "</table>" +
    // Note
    '<div style="font-family:' + serif + ";font-size:13pt;color:" + green + ';margin:16px 0 6px;text-align:' + start + '">' + esc_(T.note) + "</div>" +
    '<div dir="auto" style="border-bottom:1px solid ' + gold + ';padding:4px 2px 6px;color:' + green + ';white-space:pre-wrap;text-align:' + start + '">' + val(v.tenant_note) + "</div>" +
    // Confirmation
    '<div style="margin:16px 0;font-size:9.5pt;color:#3A3A3A;text-align:' + start + '">&#9745; ' + esc_(T.confirm) + "</div>" +
    // Signature and date
    '<table style="width:100%;border-collapse:collapse;margin-top:8px"><tr>' +
    '<td style="width:50%;padding-' + end + ':24px"><div dir="auto" style="border-bottom:1px solid ' + green + ';font-size:13pt;color:' + green +
    ';padding:4px 2px;text-align:' + start + '">' + val(v.tenant_signature) + "</div>" +
    '<div style="font-size:8.5pt;color:' + green + ';margin-top:5px;' + caps + 'text-align:' + start + '">' + esc_(T.signature) + "</div>" +
    '<div style="font-size:8pt;color:#7C7C7C;margin-top:2px;text-align:' + start + '">' + esc_(T.signatureHint) + "</div></td>" +
    '<td style="width:50%;padding-' + start + ':24px"><div dir="ltr" style="border-bottom:1px solid ' + green + ';font-size:13pt;color:' + green +
    ';padding:4px 2px;text-align:' + start + '">' + val(v.tenant_date) + "</div>" +
    '<div style="font-size:8.5pt;color:' + green + ';margin-top:5px;' + caps + 'text-align:' + start + '">' + esc_(T.date) + "</div></td></tr></table>" +
    // Footer
    '<div style="border-top:1px solid ' + gold + ';margin-top:28px;padding-top:8px;text-align:center;font-size:7.5pt;color:#8A8A8A">' + esc_(T.footer) + "</div>" +
    "</body></html>";
}

/**
 * Run this once from the Apps Script editor (choose testOwnerPdf, then Run).
 * It saves two SAMPLE owner PDFs (English and Arabic) in the owner folder
 * and writes how long each one took in the Execution log.
 */
function testOwnerPdf() {
  const values = {
    full_name: "Ahmed Al-Balushi", email: "ahmed@example.com", nationality: "Omani",
    work_at: "Ministry of Health", position: "Engineer", family_size: "4",
    rental_amount: "450 OMR", payment_method: "Cheque", security_payment: "450 OMR",
    contract_start_date: "01/11/2026", potential_move_in_date: "05/11/2026",
    tenant_note: "Needs parking for 2 cars.", tenant_signature: "Ahmed Al-Balushi", tenant_date: "27/10/2026"
  };
  const arValues = Object.assign({}, values, {
    full_name: "أحمد البلوشي", nationality: "عُماني", work_at: "وزارة الصحة", position: "مهندس",
    rental_amount: "450 ر.ع", payment_method: "شيك", security_payment: "450 ر.ع",
    tenant_note: "يحتاج موقفين للسيارات.", tenant_signature: "أحمد البلوشي"
  });
  [["en", values], ["ar", arValues]].forEach(function (x) {
    const started = Date.now();
    const now = new Date();
    const file = createOwnerPdf_({
      lang: x[0], ref: "SAMPLE" + x[0].toUpperCase(), time: formatTime_(now),
      day: Utilities.formatDate(now, TIME_ZONE, "yyyy-MM-dd"), pref: "SAMPLE", bno: "45", values: x[1]
    });
    console.log(x[0] + ": " + file.getName() + " — " + ((Date.now() - started) / 1000).toFixed(1) + " seconds — " + file.getUrl());
  });
}

// ---------- Helpers ----------

/** Reference number shown everywhere: first 8 characters of the Submission ID, upper case */
function referenceOf_(id) {
  return String(id).replace(/-/g, "").slice(0, 8).toUpperCase();
}

function formatTime_(date) {
  return Utilities.formatDate(date, TIME_ZONE, "dd/MM/yyyy HH:mm");
}

function esc_(s) {
  return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

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
  } else {
    // Sheet made by an older version: add any new header at the end, keep the rest as is
    const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    current.forEach(function (h, i) {
      if (h === "") {
        sheet.getRange(1, i + 1).setValue(headers[i])
          .setFontWeight("bold").setBackground("#11502C").setFontColor("#FFFFFF");
      }
    });
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Logo (small PNG, used only in the owner PDF) ----------
const LOGO_PNG_BASE64 = "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";
