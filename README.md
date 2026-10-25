# Tenant Information Form

A web page for Middle East Properties Global S.P.C, built for phones and available in English and Arabic. Each submission is saved as a new row in a Google Sheet through a Google Apps Script web app.

- **Setup guide (Arabic):** [SETUP-ar.md](SETUP-ar.md)
- **Customer link:** `https://ali-kalash.github.io/tenant-form/?k=<KEY>`. The key comes from the `Properties` sheet; use the sheet menu **Tenant Form → Generate link**.
- **Where the script link goes:** set `SCRIPT_URL` at the top of [app.js](app.js).

## Files

| File | What it is |
|---|---|
| `index.html`, `styles.css`, `app.js` | The page. Static files only, with no build step. |
| `assets/logo.png` | A smaller copy of the logo. |
| `apps-script/Code.gs` | The Apps Script that runs inside the Google Sheet. |
| `design/` | The original Claude Design export, kept for reference only. |
