KC Overseas - UK Student Visa IHS + Funds Calculator

Purpose
- Internal, team-proof calculator for IHS and UKVI funds checks with a client-facing PDF.
- Minimal manual input; auto rules; reduce formula edits and human error.

How it works (high level)
- Node/Express server serves static UI from `public/`.
- Config and rates live in `data/ukvi_config.json` (optionally merged with remote `CONFIG_URL`).
- Core calculations happen server-side in `server.js`:
  - IHS calculation
  - Funds required (tuition + maintenance + dependants)
  - Funds available (FX conversion + 28-day + 31-day checks)
  - PDF generation (client or internal view)
  - IHS grant/timeline logic aligned to Appendix Student ST 25.3 and decision-time defaults

Runtime
- Start: `npm start` (Node 18+)
- Dev: `npm run dev`
- Health: `/healthz`

Key inputs
- University (auto maps London vs outside London)
- Course start/end dates
- Visa application date (required for 31-day check)
- Intended travel date (used only when grant is < 1 month before course start)
- Visa service type + decision time (estimates grant date)
- Pre-sessional toggle (only relevant for courses under 6 months)
- Fees paid/waiver/dependants
- Bank statement rows (currency, balance, dates)

Key outputs
- IHS total
- Visa duration + 6-month chargeable blocks (IHS rounding basis)
- Funds required (GBP)
- Funds available (eligible vs all)
- Eligibility status + gap
- PDF report

Known issues (blockers)
- None confirmed after the 2026-01-07 fix pass; verify in staging if any edge cases remain.

Recommended fixes (first pass)
- Fix `getFx` signature to accept `manualFx`, and pass it from `/api/fx`, `/api/report`, and `/api/pdf`.
- Correct `/api/report` gap calculation to use `fundsAvail.summary.totalEligibleGbp` and/or `totalAllGbp`.
- Remove duplicate IDs in `public/index.html` and align `public/app.js` with the actual input elements.
- Remove duplicate/unused functions in `public/app.js` and standardize the university search flow.
- Normalize encoding (UTF-8) to remove stray characters in UI text.

Roadmap alignment (from brief)
- Phase 1: FX caching, retries, PDF layout stability, university list sync
- Phase 2: PDF audit stamps (version, FX timestamp, case ID), advisor confirmation
- Phase 3: Admin panel, auth gate, case history storage

Progress log (manual updates)
- 2026-01-07: Initial static review completed; blockers documented.
- 2026-01-07: Fixed FX/manual fallback, UI ID collisions, missing handlers, and encoding issues; ran `node --check` on `server.js` and `public/app.js`.
- 2026-01-08: Updated IHS calculation workflow to include grant-date estimation (working-day decision time), intended-travel clamping, and ST 25.3 rule path highlighting. Rebuilt the IHS page layout with timeline + scenario strip and added rule source link and wrap-up table visibility.


