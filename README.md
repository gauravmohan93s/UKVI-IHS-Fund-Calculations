# UK Visa Calculator (v2) - IHS + Funds Required + Funds Available + PDF

## Run locally
Node.js 18+ required.

```bash
npm install
npm start
```
Open http://localhost:3000

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
$env:STUDENTS_SOURCE_URL="https://docs.google.com/spreadsheets/d/<id>/export?format=xlsx"
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
- `TEAMS_WEBHOOK_URL` (send sync status to Teams/Power Automate)

The app always serves from local files (`data/students.xlsx`, `data/counselors.csv`,
`data/country_currency.json`) and refreshes them in the background.

To trigger a manual sync (useful for Render cron), call:
```bash
curl -X POST "https://<your-app>.onrender.com/api/sync?targets=students"
```
If access control is enabled, send `X-Access-Code` header.


## 28-day validator
Funds rows include statement start/end; app marks each row eligible/ineligible and sums eligible funds.

## Keep-alive (Render free tier)
Add a GitHub Actions secret named `RENDER_URL` (e.g. `https://your-app.onrender.com`).
The workflow in `.github/workflows/keepalive.yml` pings `/healthz` every 10 minutes.

## Student + counselor lookups (optional)
Place files in `data/` (or override with env vars):

- `data/students.xlsx` (default, override with `STUDENTS_XLSX_PATH`)
  Columns used: `AcknowledgementNumber` or `AcknowledgmentNumber`, `StudentName`,
  `ProgramName`, `University`, `STATUS`, `Intake InYear`, `ApplicationStageChangedOn`,
  `Assignee`, `AssigneeEmail`, `DOB`, `Gender`, `MaritalStatus`, `City`, `Country`

- `data/counselors.csv` (default, override with `COUNSELORS_CSV_PATH`)
  Columns used: `Employee ID`, `Name`, `Email ID (Official)`, `Region`, `Sub Region`,
  `Designation`, `Roles`
