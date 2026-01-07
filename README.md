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
