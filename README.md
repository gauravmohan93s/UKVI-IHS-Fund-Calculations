# UK Visa Calculator (v2) – IHS + Funds Required + Funds Available + PDF

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
