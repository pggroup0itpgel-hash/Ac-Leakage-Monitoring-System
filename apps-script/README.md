# Apps Script Add-on (OTP + Sheet Mirror)

This keeps Firebase storage as-is and adds Google Sheet sync + OTP.

## 1) Create Google Sheet tabs (or let script create)

- `activity_logs`
- `admins`
- `employees`
- `settings`
- `locations`
- `plants`
- `lines`
- `defects`
- `defects_by_location` (location, defectsCsv)
- `otp`
- `defectreports_locationplantwise`
- `reports_<location>` (auto-created per location for non-Pune submissions)
- `product_catalog` (location-wise QR product/form config)
- `ui_settings` (`settingsPin`, `excelPin`)
- `setup_sheet` (paste location-wise defect + QR form config rows)

## 2) Deploy Apps Script

1. Open [script.new](https://script.new)
2. Paste `Code.gs` content
3. Set `SHEET_ID`
4. Deploy -> New deployment -> Web app
  - Execute as: **Me**
  - Who has access: **Anyone**
5. Copy Web App URL

## 3) Netlify frontend setup (automatic, no manual localStorage)

Project already includes:

- `netlify/functions/otp.js` (proxy)
- `netlify.toml` redirect (`/api/otp` -> function)

Frontend calls `/api/otp` automatically.
No need to run any browser console command.

Important:

- In `netlify/functions/otp.js`, Apps Script URL is already set to:
`https://script.google.com/macros/s/AKfycbzkSCykxKGF_HDNvNiEjCMoW25C3HD7DkGHFTYwfEeBpYPgPhWq6RcQTsISiDgsN_5KgA/exec`

## Local testing

- `Live Server (127.0.0.1:5500)` cannot run Netlify Functions.
- Use:
  - `npm i -g netlify-cli`
  - `netlify dev`
- Or run one-click file: `start-local.bat`
- App auto-uses `http://localhost:8888/api/otp` in local mode.

## 4) Behavior

- OTP send/verify goes via Apps Script (`MailApp`).
- Login/session remains in web app (7 days).
- Report submission:
  - Firebase write continues (`leakageReportsByPlant/...`)
  - Sheet mirror also writes to `defectreports_locationplantwise`
- Settings save from IT admin mirrors to:
  - `settings`, `plants`, `locations`, `lines`, `defects`, `admins`, `employees`
- Activity logs mirror to `activity_logs`.
- You can fetch recent audit logs from API action `getActivityLogs`.
- Dashboard can pull sheet reports via API action `getDefectReports`.

