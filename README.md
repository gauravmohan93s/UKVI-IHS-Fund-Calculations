# UK Visa Calculator (v2) - IHS + Funds Required + Funds Available + PDF

## Run locally
Node.js 18+ required.

```bash
npm install
npm start
```
Open http://localhost:3000

## Optional: build minified assets
Creates `public/app.min.js` and `public/styles.min.css`. The server will serve these automatically if present.
```bash
npm run build
```

## Online config (recommended)
Host a JSON file (same shape as `data/ukvi_config.json`) and set:

### Windows PowerShell
```powershell
$env:CONFIG_URL="https://raw.githubusercontent.com/<org>/<repo>/main/ukvi_config.json"
npm start
```

### Mac/Linux
```bash
CONFIG_URL="https://raw.githubusercontent.com/<org>/<repo>/main/ukvi_config.json" npm start
```

## Data sync (recommended for Render)
The app can auto-sync Google Sheets (or any URL) into local cache files and serve
from disk for fast user requests.

Set URLs for exports and a sync interval in milliseconds:
```powershell
$env:STUDENTS_SOURCE_URL="https://docs.google.com/spreadsheets/d/<id>/export?format=csv"
$env:COUNSELORS_SOURCE_URL="https://docs.google.com/spreadsheets/d/<id>/export?format=csv"
$env:COUNTRY_CURRENCY_SOURCE_URL="https://example.com/country_currency.json"
$env:STUDENTS_SYNC_MS="86400000"        # daily
$env:COUNSELORS_SYNC_MS="1728000000"    # 20 days
$env:COUNTRY_CURRENCY_SYNC_MS="86400000"
npm start
```

Optional env vars:
- `DATA_SYNC_ENABLED=true|false` (default: true)
- `STUDENTS_SOURCE_URL`, `COUNSELORS_SOURCE_URL`, `COUNTRY_CURRENCY_SOURCE_URL`
- `STUDENTS_SYNC_MS`, `COUNSELORS_SYNC_MS`, `COUNTRY_CURRENCY_SYNC_MS`
- `STUDENTS_SYNC_MIN_AGE_MINUTES`, `COUNSELORS_SYNC_MIN_AGE_MINUTES`, `COUNTRY_CURRENCY_SYNC_MIN_AGE_MINUTES`
- `DATA_DIR` (base data directory; on Render use `/var/data`)
- `CONFIG_CACHE_MS` (cache `/api/config` remote fetch, default: 300000)
- `QUERY_CACHE_MS`, `QUERY_CACHE_MAX` (cache student/counselor search results)
- `AUDIT_LOG_PATH` (audit log file path, default: `data/audit.log`)
- `AUDIT_LOG_MAX_BYTES`, `AUDIT_LOG_RETENTION` (rotate audit logs)
- `AUDIT_WEBHOOK_URL`, `AUDIT_WEBHOOK_TIMEOUT_MS` (send audit events to a webhook)
- `STUDENTS_CSV_PATH` (override default `data/students.csv`)
- `BACKUP_ROOT` (backup target folder, default `backups/`)
- `TEAMS_WEBHOOK_URL` (send sync status to Teams/Power Automate)
- `FX_FALLBACK_URL` (override FX fallback provider, default: open.er-api.com)

The app always serves from local files (`data/students.csv`, `data/counselors.csv`,
`data/country_currency.json`) and refreshes them in the background.

To trigger a manual sync (useful for Render cron), call:
```bash
curl -X POST "https://<your-app>.onrender.com/api/sync?targets=students"
```
If access control is enabled, send `X-Access-Code` header.

Optional: avoid redundant syncs when multiple schedulers are used:
```bash
curl -X POST "https://<your-app>.onrender.com/api/sync?targets=students&min_age_minutes=30"
```
This skips the sync if the last successful update is newer than the threshold.

## Helpers
Manual students CSV sync:
```bash
npm run sync:students
```

Backup `data/`:
```bash
npm run backup:data
```

## Uptime (Windows)
Use a process manager:
- PM2: `npm i -g pm2`, then `pm2 start server.js --name kc-visa` and `pm2 save`
- Or Windows Task Scheduler to start `npm start` at boot

## Audit logs to Google Sheet (Apps Script)
If you’re on Render free tier (no persistent disk), you can push audit events to a Google Sheet.

1) Create a Google Sheet with headers matching your fields (e.g., `ts`, `event`, `correlationId`, `method`, `path`, `ip`, `accessCode`, `studentAck`, `counselorEmail`, `currency`, `gapEligibleGbp`, `ok`).
2) In the Sheet, open **Extensions → Apps Script** and paste:
```javascript
function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = JSON.parse(e.postData.contents || "{}");
  sheet.appendRow([
    data.ts || "",
    data.event || "",
    data.correlationId || "",
    data.method || "",
    data.path || "",
    data.ip || "",
    data.accessCode || "",
    data.studentAck || "",
    data.counselorEmail || "",
    data.currency || "",
    data.gapEligibleGbp ?? "",
    data.ok ?? ""
  ]);
  return ContentService.createTextOutput("ok");
}
```
3) Deploy: **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
4) Copy the Web App URL and set:
```
AUDIT_WEBHOOK_URL=<your apps script web app url>
```

## Render setup (step-by-step)
1) Create a new **Web Service** on Render and connect your GitHub repo.
2) Set **Build Command**: `npm install && npm run build`
3) Set **Start Command**: `npm start`
4) Set **Health Check Path**: `/healthz`
5) Add a **Persistent Disk** mounted at `/var/data` (recommended for logs + cached data).
6) Set environment variables (Render → Environment):
   - `DATA_DIR=/var/data`
   - `CONFIG_URL=<your remote config URL>`
   - `STUDENTS_SOURCE_URL=<google sheet csv export>`
   - `COUNSELORS_SOURCE_URL=<csv export>`
   - `COUNTRY_CURRENCY_SOURCE_URL=<json url>`
   - `STUDENTS_SYNC_MS=86400000`
   - `COUNSELORS_SYNC_MS=1728000000`
   - `COUNTRY_CURRENCY_SYNC_MS=86400000`
   - `ACCESS_ENABLED=true`
   - `ACCESS_CODE=<your access code>`
   - `AUDIT_LOG_MAX_BYTES=5242880`
   - `AUDIT_LOG_RETENTION=10`
7) Deploy.
8) Optional: set a cron/job (external or Render cron) to call:
   `POST /api/sync?targets=students,counselors,country_currency&min_age_minutes=30`

## 28-day validator
Funds rows include statement start/end; app marks each row eligible/ineligible and sums eligible funds.

## Keep-alive (Render free tier)
Add a GitHub Actions secret named `RENDER_URL` (e.g. `https://your-app.onrender.com`).
The workflow in `.github/workflows/keepalive.yml` pings `/healthz` every 10 minutes.

## Student + counselor lookups (optional)
Place files in `data/` (or override with env vars):

- `data/students.csv` (default, override with `STUDENTS_CSV_PATH`)
  Columns used: `AcknowledgementNumber` or `AcknowledgmentNumber`, `StudentName`,
  `ProgramName`, `University`, `STATUS`, `Intake`, `InYear`, `Intake InYear`,
  `ApplicationStageChangedOn`, `Assignee`, `AssigneeEmail`, `DOB`, `Gender`,
  `MaritalStatus`, `City`, `Country`, `Gross Tuition Fees`, `Scholarship`,
  `Deposit Paid`

- `data/counselors.csv` (default, override with `COUNSELORS_CSV_PATH`)
  Columns used: `Employee ID`, `Name`, `Email ID (Official)`, `Region`, `Sub Region`,
  `Designation`, `Roles`
