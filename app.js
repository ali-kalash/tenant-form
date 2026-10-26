// ============================================================
//  Paste your Google Apps Script Web App link between the quotes.
//  ضع رابط Google Apps Script بين علامتي التنصيص.
// ============================================================
const SCRIPT_URL = "PASTE_YOUR_SCRIPT_URL_HERE";
// ============================================================

(function () {
  "use strict";

  const TEXT = {
    en: {
      tagline: "Real Estate Brokerage · Muscat",
      title: "Tenant Information Form",
      pref: "P- Ref:",
      bno: "B- No.:",
      detailsTitle: "Tenant Details:",
      colItem: "Item",
      colDetails: "Details",
      full_name: "Full Name",
      email: "Email address",
      optional: "(optional)",
      nationality: "Nationality",
      work_at: "Work at",
      position: "In position of",
      family_size: "A family of (Number)",
      rental_amount: "Rental amount agreed upon",
      payment_method: "Payment method",
      security_payment: "Security payment",
      contract_start_date: "Contract start date",
      potential_move_in_date: "Potential move-in date",
      agent_fee: "Agent fee (%)",
      noteTitle: "Tenant’s Note & Feedback:",
      confirm: "The tenant confirmed the accuracy of the above data and agreed to it.",
      signature: "Tenant Signature",
      signatureHint: "Type your full name above — it counts as your signature.",
      date: "Date",
      send: "Send",
      sending: "Sending…",
      successTitle: "Thank you!",
      successText: "Your information has been received successfully.",
      footer: "Bousher – Muscat – Sultanate of Oman - P.O.Box: 434 - PC: 320",
      loading: "Loading property details…",
      invalidLink: "This link is not valid. Please contact your agent for a correct link.",
      loadError: "We couldn’t load the property details. Please check your connection and reload the page.",
      notConfigured: "This form is not set up yet. Please contact your agent.",
      errRequired: "This field is required.",
      errEmail: "Please enter a valid email address.",
      errNumber: "Please enter a number (e.g. 4).",
      errConfirm: "Please confirm that the data is accurate.",
      errFuture: "The date can’t be in the future.",
      errSummary: "Please correct the fields marked in red.",
      errSend: "The form was not sent. Please try again.",
      langButton: "العربية"
    },
    ar: {
      tagline: "وساطة عقارية · مسقط",
      title: "نموذج معلومات المستأجر",
      pref: "مرجع العقار:",
      bno: "رقم المبنى:",
      detailsTitle: "بيانات المستأجر:",
      colItem: "البند",
      colDetails: "التفاصيل",
      full_name: "الاسم الكامل",
      email: "البريد الإلكتروني",
      optional: "(اختياري)",
      nationality: "الجنسية",
      work_at: "جهة العمل",
      position: "المسمى الوظيفي",
      family_size: "عدد أفراد العائلة",
      rental_amount: "مبلغ الإيجار المتفق عليه",
      payment_method: "طريقة الدفع",
      security_payment: "مبلغ التأمين",
      contract_start_date: "تاريخ بدء العقد",
      potential_move_in_date: "تاريخ الانتقال المتوقع",
      agent_fee: "عمولة الوكيل (%)",
      noteTitle: "ملاحظات المستأجر:",
      confirm: "أقرّ المستأجر بصحة البيانات أعلاه ووافق عليها.",
      signature: "توقيع المستأجر",
      signatureHint: "اكتب اسمك الكامل أعلاه، ويُعتبر ذلك توقيعاً منك.",
      date: "التاريخ",
      send: "إرسال",
      sending: "جارٍ الإرسال…",
      successTitle: "شكراً لك!",
      successText: "تم استلام بياناتك بنجاح.",
      footer: "بوشر – مسقط – سلطنة عُمان - ص.ب: 434 - الرمز البريدي: 320",
      loading: "جارٍ تحميل بيانات العقار…",
      invalidLink: "هذا الرابط غير صالح. يرجى التواصل مع الوكيل للحصول على رابط صحيح.",
      loadError: "تعذّر تحميل بيانات العقار. تحقق من اتصالك بالإنترنت ثم أعد تحميل الصفحة.",
      notConfigured: "هذا النموذج غير جاهز بعد. يرجى التواصل مع الوكيل.",
      errRequired: "هذا الحقل مطلوب.",
      errEmail: "صيغة البريد الإلكتروني غير صحيحة.",
      errNumber: "اكتب رقماً (مثلاً 4).",
      errConfirm: "يرجى تأكيد صحة البيانات.",
      errFuture: "لا يمكن أن يكون التاريخ في المستقبل.",
      errSummary: "يرجى تصحيح الحقول المظللة بالأحمر.",
      errSend: "لم يتم الإرسال. يرجى المحاولة مرة أخرى.",
      langButton: "English"
    }
  };

  const REQUIRED_TEXT = [
    "full_name", "nationality", "work_at", "position", "family_size",
    "rental_amount", "payment_method", "security_payment", "tenant_signature"
  ];
  const REQUIRED_DATES = ["contract_start_date", "potential_move_in_date", "tenant_date"];
  const ALL_FIELDS = [
    "full_name", "email", "nationality", "work_at", "position", "family_size",
    "rental_amount", "payment_method", "security_payment",
    "contract_start_date", "potential_move_in_date", "tenant_note",
    "tenant_signature", "tenant_date"
  ];
  const KEY_PATTERN = /^[A-Za-z0-9]{12,32}$/;
  const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  const HONEYPOT = "mep_xq_alt7"; // hidden field, must match Code.gs

  const $ = (id) => document.getElementById(id);
  const form = $("tenant-form");
  const statusPanel = $("status-panel");
  const statusText = $("status-text");
  const successPanel = $("success-panel");
  const sendBtn = $("send");
  const formError = $("form-error");

  let lang = "en";
  let statusKey = "loading";
  let errors = {};          // field id -> text key
  let attempted = false;
  let sending = false;
  let sendFailed = false;
  const key = new URLSearchParams(location.search).get("k") || "";
  const submissionId = newId();

  // ---------- Language ----------
  function t(k) { return TEXT[lang][k]; }

  function applyLanguage() {
    document.documentElement.lang = lang;
    document.documentElement.dir = lang === "ar" ? "rtl" : "ltr";
    document.querySelectorAll("[data-i18n]").forEach((el) => {
      el.textContent = t(el.dataset.i18n);
    });
    const btn = $("lang-toggle");
    btn.textContent = t("langButton");
    btn.lang = lang === "ar" ? "en" : "ar";
    statusText.textContent = t(statusKey);
    if (sending) sendBtn.textContent = t("sending");
    renderErrors();
  }

  $("lang-toggle").addEventListener("click", () => {
    lang = lang === "en" ? "ar" : "en";
    applyLanguage();
  });

  // ---------- Status panel ----------
  function showStatus(k, isError) {
    statusKey = k;
    statusText.textContent = t(k);
    statusPanel.classList.toggle("panel-error", !!isError);
    statusPanel.hidden = false;
    form.hidden = true;
  }

  // ---------- Talking to Apps Script ----------
  function scriptConfigured() {
    return /^https:\/\/script\.google(usercontent)?\.com\//.test(SCRIPT_URL);
  }

  async function fetchJson(url, options) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), 25000);
    try {
      const res = await fetch(url, Object.assign({ signal: ctrl.signal, redirect: "follow" }, options));
      if (!res.ok) throw new Error("HTTP " + res.status);
      return await res.json();
    } finally {
      clearTimeout(timer);
    }
  }

  async function loadProperty() {
    if (!scriptConfigured()) return showStatus("notConfigured", true);
    if (!KEY_PATTERN.test(key)) return showStatus("invalidLink", true);
    let data;
    try {
      data = await fetchJson(SCRIPT_URL + "?action=property&k=" + encodeURIComponent(key));
    } catch (e) {
      return showStatus("loadError", true);
    }
    if (!data || data.ok !== true) return showStatus("invalidLink", true);
    $("v-pref").textContent = data.pref || "—";
    $("v-bno").textContent = data.bno || "—";
    const fee = String(data.fee || "").replace(/\s*%\s*$/, "");
    $("v-fee").textContent = fee ? fee + " %" : "—";
    statusPanel.hidden = true;
    form.hidden = false;
  }

  // ---------- Validation ----------
  function value(id) { return $(id).value.trim(); }

  // Arabic-Indic digits (٠-٩ and ۰-۹) → 0-9
  function toWesternDigits(s) {
    return s.replace(/[\u0660-\u0669\u06F0-\u06F9]/g, (d) => String(d.charCodeAt(0) & 0xf));
  }

  function validate() {
    errors = {};
    REQUIRED_TEXT.concat(REQUIRED_DATES).forEach((id) => {
      if (!value(id)) errors[id] = "errRequired";
    });
    if (value("email") && !EMAIL_PATTERN.test(value("email"))) errors.email = "errEmail";
    if (value("family_size") && !/^[1-9][0-9]?$/.test(toWesternDigits(value("family_size")))) errors.family_size = "errNumber";
    if (value("tenant_date") && value("tenant_date") > todayISO()) errors.tenant_date = "errFuture";
    if (!$("confirmed").checked) errors.confirmed = "errConfirm";
    renderErrors();
    return Object.keys(errors).length === 0;
  }

  function renderErrors() {
    ALL_FIELDS.concat("confirmed").forEach((id) => {
      const input = $(id);
      const holder = input.closest(".tr, .sign-col, .confirm, section");
      const errEl = id === "confirmed" ? $("err-confirmed") : holder.querySelector(".err");
      const k = errors[id];
      if (holder) holder.classList.toggle("invalid", !!k);
      input.setAttribute("aria-invalid", k ? "true" : "false");
      if (errEl) errEl.textContent = k ? t(k) : "";
    });
    const k = Object.keys(errors).length ? "errSummary" : sendFailed ? "errSend" : "";
    formError.textContent = k ? t(k) : "";
    formError.hidden = !k;
  }

  form.addEventListener("input", () => { if (attempted) validate(); });
  form.addEventListener("change", () => { if (attempted) validate(); });

  // ---------- Dates ----------
  // <input type="date"> gives YYYY-MM-DD; the sheet gets DD/MM/YYYY.
  function toDMY(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
    return m ? m[3] + "/" + m[2] + "/" + m[1] : "";
  }

  function todayISO() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }

  // ---------- Submit ----------
  function newId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    const b = new Uint8Array(16);
    crypto.getRandomValues(b);
    return Array.from(b, (x) => x.toString(16).padStart(2, "0")).join("");
  }

  form.addEventListener("submit", async (ev) => {
    ev.preventDefault();
    if (sending) return;
    attempted = true;
    sendFailed = false;
    if (!validate()) {
      const first = form.querySelector('[aria-invalid="true"]');
      if (first) first.focus();
      return;
    }

    const payload = {
      action: "submit",
      k: key,
      id: submissionId,
      lang: lang,
      [HONEYPOT]: $(HONEYPOT).value,
      confirmed: true
    };
    ALL_FIELDS.forEach((id) => { payload[id] = value(id); });
    REQUIRED_DATES.forEach((id) => { payload[id] = toDMY(value(id)); });
    payload.family_size = toWesternDigits(payload.family_size);

    sending = true;
    sendBtn.disabled = true;
    sendBtn.textContent = t("sending");
    renderErrors();

    let ok = false;
    let result = "";
    try {
      const data = await fetchJson(SCRIPT_URL, {
        method: "POST",
        // text/plain keeps the browser from blocking the request to Google
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload)
      });
      result = data && data.result;
      ok = result === "success";
    } catch (e) {
      ok = false;
    }

    sending = false;
    if (ok) {
      form.hidden = true;
      successPanel.hidden = false;
      successPanel.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    if (result === "invalid_link") return showStatus("invalidLink", true);
    sendBtn.disabled = false;
    sendBtn.textContent = t("send");
    sendFailed = true;
    renderErrors();
  });

  // ---------- Start ----------
  $("tenant_date").value = todayISO();
  $("tenant_date").max = todayISO();
  applyLanguage();
  loadProperty();
})();
