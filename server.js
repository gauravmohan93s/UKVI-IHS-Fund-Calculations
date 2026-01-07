import express from "express";
import morgan from "morgan";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import PDFDocument from "pdfkit";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(morgan("dev"));
app.use(express.json({ limit: "2mb" }));

// --- Simple access-code gate (optional)
// Enable by setting config.access.enabled=true or ACCESS_ENABLED=true.
// Set ACCESS_CODE env var (preferred) or config.access.code.
// Frontend sends header: X-Access-Code.
function getAccessCode(config){
  const envCode = process.env.ACCESS_CODE;
  const cfgCode = config?.access?.code;
  const enabled = Boolean(process.env.ACCESS_ENABLED === "true" || config?.access?.enabled);
  const code = (envCode || cfgCode || "").trim();
  return { enabled, code };
}

app.use((req, res, next) => {
  if (!req.path.startsWith("/api/")) return next();
  if (req.path === "/api/config") return next(); // allow UI to load
  const config = readLocalConfig();
  const { enabled, code } = getAccessCode(config);
  if (!enabled) return next();

  const provided = String(req.header("x-access-code") || "").trim();
  if (code && provided === code) return next();

  return res.status(401).json({ error: "Unauthorized: invalid access code" });
});

app.use(express.static(path.join(__dirname, "public")));

const LOCAL_CONFIG_PATH = path.join(__dirname, "data", "ukvi_config.json");
const CONFIG_URL = process.env.CONFIG_URL || ""; // optional remote JSON

function readLocalConfig() {
  return JSON.parse(fs.readFileSync(LOCAL_CONFIG_PATH, "utf-8"));
}

async function fetchJson(url, timeoutMs = 8000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const res = await fetch(url, { signal: ctrl.signal, headers: { "accept": "application/json" } });
    if (!res.ok) throw new Error(`Fetch failed (${res.status})`);
    return await res.json();
  } finally {
    clearTimeout(t);
  }
}

// --- FX cache (daily) ---
const fxCache = new Map(); // key => data
function todayKey() { return new Date().toISOString().slice(0, 10); }

async function getFx(from, toCsv) {
  const key = `${todayKey()}|${from}|${toCsv}`;
  if (fxCache.has(key)) return fxCache.get(key);

  const url = `https://api.frankfurter.app/latest?from=${encodeURIComponent(from)}&to=${encodeURIComponent(toCsv)}`;
  const data = await fetchJson(url, 8000);
  fxCache.set(key, data);
  return data;
}

// --- Helpers ---
function safeNum(x){ const n = Number(x); return Number.isFinite(n) ? n : 0; }
function round2(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }
function round6(x){ return Math.round((x + Number.EPSILON) * 1e6) / 1e6; }
function daysBetween(d1, d2){ return Math.max(0, Math.ceil((new Date(d2) - new Date(d1)) / 86400000)); }

function monthsForMaintenance(startDate, endDate, maxMonths){
  const days = daysBetween(startDate, endDate);
  const m = Math.max(1, Math.ceil(days / 30.44));
  return Math.min(maxMonths, m);
}

function monthsForIHS(visaStart, visaEnd){
  const days = daysBetween(visaStart, visaEnd);
  return Math.max(0, Math.ceil(days / 30.44));
}

function calcIHSStudent(visaStart, visaEnd, ihsYearly, ihsHalf, applyingFrom="outside"){
  const months = monthsForIHS(visaStart, visaEnd);
  if (months <= 6) return applyingFrom === "outside" ? 0 : ihsHalf;
  if (months <= 12) return ihsYearly;

  const fullYears = Math.floor(months / 12);
  const rem = months % 12;
  let total = fullYears * ihsYearly;
  if (rem === 0) return total;
  total += (rem <= 6) ? ihsHalf : ihsYearly;
  return total;
}

function normalizeRegion(val){
  const v = String(val || "").toLowerCase();
  if (v.includes("london") && !v.includes("outside")) return "london";
  return "outside_london";
}

function calcFundsRequired(payload, config){
  const routeKey = payload.routeKey || "student";
  const route = config.routes[routeKey];
  if (!route) throw new Error(`Unknown routeKey: ${routeKey}`);

  const region = normalizeRegion(payload.region || payload.studyLocation || "outside_london");

  const start = payload.courseStart;
  const end = payload.courseEnd;
  if (!start || !end) throw new Error("courseStart and courseEnd required");

  const months = monthsForMaintenance(start, end, route.max_months);

  const tuitionTotal = safeNum(payload.tuitionFeeTotalGbp);
  const tuitionPaid = safeNum(payload.tuitionFeePaidGbp);
  const scholarship = safeNum(payload.scholarshipGbp);
  const tuitionDue = Math.max(0, tuitionTotal - tuitionPaid - scholarship);

  const studentMonthly = route.maintenance_monthly_gbp[region];
  const maintenanceStudent = months * studentMonthly;

  let dependants = Math.max(0, Math.floor(safeNum(payload.dependantsCount)));
  if (!route.dependants_allowed) dependants = 0;

  let maintenanceDependants = 0;
  if (route.dependants_allowed) {
    const depMonthly = route.dependant_monthly_gbp[region];
    maintenanceDependants = months * depMonthly * dependants;
  }

  const buffer = safeNum(payload.bufferGbp);
  const fundsRequired = tuitionDue + maintenanceStudent + maintenanceDependants + buffer;

  return {
    routeKey, region, monthsRequired: months,
    tuitionDueGbp: round2(tuitionDue),
    maintenanceStudentGbp: round2(maintenanceStudent),
    maintenanceDependantsGbp: round2(maintenanceDependants),
    bufferGbp: round2(buffer),
    fundsRequiredGbp: round2(fundsRequired),
    dependantsCountEffective: dependants
  };
}

async function calcFundsAvailable(payload, rules){
  // rows: [{accountType, source, currency, amount, statementStart, statementEnd}]
  const rows = Array.isArray(payload.fundsRows) ? payload.fundsRows : [];
  const applicationDate = payload.applicationDate ? new Date(payload.applicationDate) : null;

  const fundsHoldDays = Number(rules?.funds_hold_days ?? 28);
  const statementAgeDays = Number(rules?.statement_age_days ?? 31);

  let totalAllGbp = 0;
  let totalEligibleGbp = 0;

  const converted = [];

  const dayDiffInclusive = (startStr, endStr) => {
    if (!startStr || !endStr) return null;
    const s = new Date(startStr);
    const e = new Date(endStr);
    if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return null;
    // inclusive day count
    const ms = e - s;
    const days = Math.floor(ms / 86400000) + 1;
    return days;
  };

  const daysSince = (fromDate, toDate) => {
    const ms = toDate - fromDate;
    return Math.floor(ms / 86400000);
  };

  for (const r of rows) {
    const accountType = String(r.accountType || "Student");
    const source = String(r.source || "");
    const currency = String(r.currency || "GBP").toUpperCase().trim();
    const amount = safeNum(r.amount);
    const statementStart = r.statementStart || "";
    const statementEnd = r.statementEnd || "";

    if (amount <= 0) continue;

    // FX convert
    let gbp = amount;
    let rate = 1;
    if (currency !== "GBP") {
      const fx = await getFx(currency, "GBP");
      rate = safeNum(fx?.rates?.GBP);
      gbp = rate ? amount * rate : 0;
    }

    totalAllGbp += gbp;

    // eligibility checks
    const issues = [];
    const periodDays = dayDiffInclusive(statementStart, statementEnd);
    if (periodDays === null) {
      issues.push("Missing/invalid statement dates");
    } else {
      if (periodDays < fundsHoldDays) issues.push(`Statement period is ${periodDays} days (< ${fundsHoldDays})`);
    }

    if (applicationDate && statementEnd) {
      const endDt = new Date(statementEnd);
      if (!Number.isNaN(endDt.getTime())) {
        const age = daysSince(endDt, applicationDate);
        if (age < 0) issues.push("Statement end is after application date");
        else if (age > statementAgeDays) issues.push(`Statement end is ${age} days before application (> ${statementAgeDays})`);
      } else {
        issues.push("Invalid statement end date");
      }
    } else if (!applicationDate) {
      // not mandatory, but warn for full validation
      issues.push("No application date (31-day freshness not checked)");
    }

    const eligible = issues.length === 0 || (issues.length === 1 && issues[0].startsWith("No application date"));
    if (eligible) totalEligibleGbp += gbp;

    converted.push({
      accountType,
      source,
      currency,
      amount: round2(amount),
      statementStart,
      statementEnd,
      fxToGbp: round6(rate),
      amountGbp: round2(gbp),
      eligible,
      issues
    });
  }

  const summary = {
    totalAllGbp: round2(totalAllGbp),
    totalEligibleGbp: round2(totalEligibleGbp),
    fundsHoldDays,
    statementAgeDays,
    hasApplicationDate: !!applicationDate,
    anyRowMissingDates: converted.some(r => !r.statementStart || !r.statementEnd),
    anyIneligibleRows: converted.some(r => !r.eligible),
  };

  return { summary, rows: converted };
}

function computeIhsBlock(payload, config, dependantsEffective){
  const visaStart = payload.visaStart || payload.courseStart;
  const visaEnd = payload.visaEnd || payload.courseEnd;
  const applyingFrom = payload.applyingFrom || "outside";
  const ihsPerPerson = calcIHSStudent(visaStart, visaEnd, config.ihs.student_yearly_gbp, config.ihs.half_year_gbp, applyingFrom);
  const persons = 1 + Math.max(0, Math.floor(safeNum(payload.ihsDependantsCount ?? dependantsEffective ?? 0)));
  return { ihsPerPersonGbp: round2(ihsPerPerson), persons, ihsTotalGbp: round2(ihsPerPerson * persons) };
}

// --- APIs ---
app.get("/api/config", async (req, res) => {
  try {
    const local = readLocalConfig();
    if (!CONFIG_URL) return res.json({ config: local, source: "local" });

    const remote = await fetchJson(CONFIG_URL, 8000);
    const merged = {
      ...local,
      ...remote,
      routes: { ...local.routes, ...(remote.routes || {}) },
      rules: { ...local.rules, ...(remote.rules || {}) },
      ihs: { ...local.ihs, ...(remote.ihs || {}) },
      fx: { ...local.fx, ...(remote.fx || {}) },
      universities: remote.universities || local.universities
    };
    res.json({ config: merged, source: "remote", config_url: CONFIG_URL });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

app.get("/api/fx", async (req, res) => {
  try {
    const from = String(req.query.from || "GBP").toUpperCase();
    const to = String(req.query.to || "INR").toUpperCase();
    if (from === to) return res.json({ provider: "frankfurter.app", from, to, rate: 1 });

    const data = await getFx(from, to);
    const rate = safeNum(data?.rates?.[to]);
    res.json({ provider: "frankfurter.app", from, to, rate, date: data?.date });
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

app.post("/api/report", async (req, res) => {
  try {
    const config = readLocalConfig();
    const payload = req.body || {};
    const fundsReq = calcFundsRequired(payload, config);
    const fundsAvail = await calcFundsAvailable(payload, config.rules);
    const gap = round2(fundsAvail.totalAvailableGbp - fundsReq.fundsRequiredGbp);

    res.json({
      fundsRequired: fundsReq,
      fundsAvailable: fundsAvail,
      gapGbp: gap,
      gapEligibleOnlyGbp: round2(fundsAvail.summary.totalEligibleGbp - fundsReq.fundsRequiredGbp),
      rules: config.rules,
      ihs: computeIhsBlock(payload, config, fundsReq.dependantsCountEffective),
    });
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

// Server-side PDF generation
app.post("/api/pdf", async (req, res) => {
  try {
    const config = readLocalConfig();
    const payload = req.body || {};
    const fundsReq = calcFundsRequired(payload, config);
    const fundsAvail = await calcFundsAvailable(payload, config.rules);
    const ihs = computeIhsBlock(payload, config, fundsReq.dependantsCountEffective);

    const gapEligible = round2(fundsAvail.summary.totalEligibleGbp - fundsReq.fundsRequiredGbp);
    const gapAll = round2(fundsAvail.summary.totalAllGbp - fundsReq.fundsRequiredGbp);
    const ok = gapEligible >= 0;

    // Quote currency for client-friendly display
    const quote = String(payload.quoteCurrency || "GBP").toUpperCase();
    let gbpToQuote = 1;
    if (quote !== "GBP") {
      try {
        const fx = await getFx("GBP", quote);
        gbpToQuote = Number(fx?.rates?.[quote]) || 1;
      } catch (_) {
        gbpToQuote = 1;
      }
    }
    const fmt = (n) => Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const fmtGBP = (n) => `£${fmt(n)}`;
    const fmtQuote = (n) => (quote === "GBP" ? fmtGBP(n) : `${quote} ${fmt(n * gbpToQuote)}`);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="UK_Visa_IHS_Funds_Report.pdf"`);

    const doc = new PDFDocument({ size: "A4", margin: 42 });
    doc.pipe(res);

    // Header band (branded)
    const b = config.branding || {};
    const c1 = b.primary_color || "#0b5cab";
    const c2 = b.secondary_color || "#00a3a3";

    doc.rect(0, 0, 595.28, 92).fill(c1);
    doc.rect(0, 62, 595.28, 30).fill(c2);

    // Logo (PNG) if present
    try {
      const logoRel = (b.logo_pdf || "public/assets/kc_logo.png").replace(/^\/+/, "");
      const logoPath = path.join(__dirname, logoRel);
      doc.image(logoPath, 42, 18, { height: 34 });
    } catch (_) {
      // ignore logo errors
    }

    doc.fillColor("#ffffff").fontSize(18).text(b.product_name || "UK Visa Calculation Report", 42, 26, { width: 511, align: "center" });
    doc.fontSize(9).text(b.company_name || "", 42, 48, { width: 511, align: "center" });
    doc.fontSize(9).fillColor("#073344").text(`Generated: ${new Date().toLocaleString()}`, 42, 71);

    // Status stamp
    const stampText = ok ? "ELIGIBLE" : "NOT ELIGIBLE";
    doc.save();
    doc.rotate(-18, { origin: [470, 120] });
    doc.fontSize(28).fillColor(ok ? "#16a34a" : "#dc2626").opacity(0.25)
      .text(stampText, 355, 95, { width: 220, align: "center" });
    doc.opacity(1).restore();

    doc.fillColor("#000000");
    let y = 108;

    const hr = () => {
      doc.moveTo(42, y).lineTo(553, y).strokeColor("#e2e8f0").stroke();
      doc.strokeColor("#000000");
      y += 10;
    };
    const section = (title) => {
      doc.fillColor("#0f172a").fontSize(12).text(title, 42, y);
      y += 16;
      hr();
    };
    const kv = (k, v) => {
      doc.fillColor("#334155").fontSize(9).text(k, 42, y, { width: 210 });
      doc.fillColor("#0f172a").fontSize(10).text(String(v ?? "-"), 260, y, { width: 293 });
      y += 14;
    };
    const money = (k, v) => kv(k, `${fmtGBP(v)}  (${fmtQuote(v)})`);

    // Client-friendly summary box
    const boxH = 72;
    doc.roundedRect(42, y, 511, boxH, 12)
      .fill(ok ? "#dcfce7" : "#fee2e2");
    doc.fillColor(ok ? "#065f46" : "#7f1d1d").fontSize(13)
      .text(ok ? "Result: ELIGIBLE (Funds are sufficient)" : "Result: NOT ELIGIBLE (Funds are short)", 56, y + 12, { width: 483 });
    doc.fillColor("#0f172a").fontSize(10)
      .text(`Funds required: ${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})`, 56, y + 34, { width: 483 });
    doc.fillColor("#0f172a").fontSize(10)
      .text(`Eligible funds available: ${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})`, 56, y + 50, { width: 483 });

    y += boxH + 18;
    doc.fillColor("#000000");

    // Inputs
    section("Inputs");
    kv("Case / File ID", payload.caseId || "-");
    kv("Advisor", payload.advisorName || "-");
    kv("Applicant", payload.applicantName || "-");
    kv("University", payload.universityName || "-");
    kv("Study location (auto)", normalizeRegion(payload.region || "outside_london") === "london" ? "London" : "Outside London");
    kv("Course dates (CAS)", `${payload.courseStart || "-"}  to  ${payload.courseEnd || "-"}`);
    kv("Visa application date", payload.applicationDate || "-");
    kv("Display currency", quote);
    kv("Dependants", fundsReq.dependantsCountEffective);
    kv("Note", "Course/fees must match CAS and offer letter.");

    // Fees
    y += 6;
    section("Fees (GBP)");
    money("Tuition total", safeNum(payload.tuitionFeeTotalGbp));
    money("Tuition paid", safeNum(payload.tuitionFeePaidGbp));
    money("Scholarship/waiver", safeNum(payload.scholarshipGbp));
    money("Buffer (optional)", safeNum(payload.bufferGbp));

    // IHS
    y += 6;
    section("IHS (GBP)");
    money("IHS per person", ihs.ihsPerPersonGbp);
    kv("Persons counted", ihs.persons);
    money("IHS total", ihs.ihsTotalGbp);

    // Funds Required
    y += 6;
    section("Funds required (28-day) — GBP");
    kv("Months counted (cap applied)", fundsReq.monthsRequired);
    money("Tuition due", fundsReq.tuitionDueGbp);
    money("Maintenance (student)", fundsReq.maintenanceStudentGbp);
    money("Maintenance (dependants)", fundsReq.maintenanceDependantsGbp);
    money("Buffer", fundsReq.bufferGbp);

    // Total required callout
    doc.roundedRect(42, y, 511, 34, 10).fill("#0f172a");
    doc.fillColor("#ffffff").fontSize(13)
      .text(`TOTAL REQUIRED: ${fmtGBP(fundsReq.fundsRequiredGbp)}   (${fmtQuote(fundsReq.fundsRequiredGbp)})`, 56, y + 10);
    doc.fillColor("#000000");
    y += 50;

    // Funds Available
    section("Funds available (validated)");
    money("Total available (all rows)", fundsAvail.summary.totalAllGbp);
    money("Total eligible (meets checks)", fundsAvail.summary.totalEligibleGbp);

    // Table
    y += 6;
    doc.roundedRect(42, y, 511, 18, 6).fill("#f1f5f9");
    doc.fillColor("#334155").fontSize(8)
      .text("OK", 50, y + 5, { width: 25 })
      .text("Account", 80, y + 5, { width: 70 })
      .text("Currency", 155, y + 5, { width: 70 })
      .text("Amount", 225, y + 5, { width: 85 })
      .text("GBP", 320, y + 5, { width: 70 })
      .text("Issues", 395, y + 5, { width: 150 });
    y += 24;

    doc.fontSize(8);
    for (const r of fundsAvail.rows.slice(0, 12)) {
      const okTxt = r.eligible ? "OK" : "NO";
      doc.fillColor(r.eligible ? "#065f46" : "#9a3412").text(okTxt, 50, y);
      doc.fillColor("#0f172a").text(String(r.accountType || "-"), 80, y, { width: 70 });
      doc.text(String(r.currency || "-"), 155, y, { width: 70 });
      doc.text(String(r.amount || 0), 225, y, { width: 85 });
      doc.text(`£${String(r.amountGbp || 0)}`, 320, y, { width: 70 });
      doc.fillColor("#475569").text((r.issues || []).join("; "), 395, y, { width: 158 });
      y += 14;
      if (y > 720) break;
    }
    if (fundsAvail.rows.length > 12 && y <= 735) {
      doc.fillColor("#475569").fontSize(8).text(`(Showing first 12 rows of ${fundsAvail.rows.length})`, 42, y);
      y += 12;
    }

    // Gap
    y += 6;
    section("Gap summary");
    kv("Gap (Eligible − Required)", fmtGBP(gapEligible) + (quote === "GBP" ? "" : `  |  ${fmtQuote(gapEligible)}`));
    kv("Gap (All rows − Required)", fmtGBP(gapAll) + (quote === "GBP" ? "" : `  |  ${fmtQuote(gapAll)}`));

    y += 8;
    doc.fillColor("#475569").fontSize(8).text(
      `Rules reminder: Funds must be held for ${config.rules.funds_hold_days} consecutive days. ` +
      `Statement end date must be within ${config.rules.statement_age_days} days of the visa application date. ` +
      `If visa application date is missing, the 31-day freshness check is skipped.`,
      42, y, { width: 511 }
    );

    // Footer note
    if (b.footer_note) {
      doc.fillColor("#64748b").fontSize(7).text(String(b.footer_note), 42, 800, { width: 511, align: "center" });
    }

    doc.end();
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

app.get("/health", (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
