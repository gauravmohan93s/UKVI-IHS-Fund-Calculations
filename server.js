import express from "express";
import morgan from "morgan";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import PDFDocument from "pdfkit";
import XLSX from "xlsx";
import { parse as parseCsv } from "csv-parse/sync";

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
const STUDENTS_XLSX_PATH = process.env.STUDENTS_XLSX_PATH || path.join(__dirname, "data", "students.xlsx");
const COUNSELORS_CSV_PATH = process.env.COUNSELORS_CSV_PATH || path.join(__dirname, "data", "counselors.csv");

function readLocalConfig() {
  return JSON.parse(fs.readFileSync(LOCAL_CONFIG_PATH, "utf-8"));
}

function normalizeStr(val) {
  return String(val || "").trim();
}

const studentsCache = { mtimeMs: 0, rows: [] };
function normalizeKeyMap(row) {
  const out = {};
  Object.entries(row || {}).forEach(([k, v]) => {
    const nk = String(k || "")
      .replace(/,/g, " ")
      .trim()
      .replace(/\s+/g, " ")
      .toLowerCase();
    out[nk] = v;
  });
  return out;
}

function loadStudents() {
  if (!fs.existsSync(STUDENTS_XLSX_PATH)) return [];
  const stat = fs.statSync(STUDENTS_XLSX_PATH);
  if (stat.mtimeMs === studentsCache.mtimeMs) return studentsCache.rows;

  const wb = XLSX.readFile(STUDENTS_XLSX_PATH);
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  const mapped = rows.map((r) => {
    const n = normalizeKeyMap(r);
    return {
      ackNumber: normalizeStr(n.acknowledgementnumber || n.acknowledgmentnumber),
      studentName: normalizeStr(n.studentname),
      programName: normalizeStr(n.programname),
      university: normalizeStr(n.university),
      status: normalizeStr(n.status),
      intakeYear: normalizeStr(n["intake inyear"] || n["intake - inyear"] || n.intakeinyear),
      applicationStageChangedOn: normalizeStr(n.applicationstagechangedon),
      assignee: normalizeStr(n.assignee),
      assigneeEmail: normalizeStr(n.assigneeemail),
      dob: normalizeStr(n.dob),
      gender: normalizeStr(n.gender),
      maritalStatus: normalizeStr(n.maritalstatus),
      city: normalizeStr(n.city),
      country: normalizeStr(n.country),
    };
  }).filter((r) => r.ackNumber || r.studentName);

  studentsCache.mtimeMs = stat.mtimeMs;
  studentsCache.rows = mapped;
  return mapped;
}

const counselorsCache = { mtimeMs: 0, rows: [] };
function loadCounselors() {
  if (!fs.existsSync(COUNSELORS_CSV_PATH)) return [];
  const stat = fs.statSync(COUNSELORS_CSV_PATH);
  if (stat.mtimeMs === counselorsCache.mtimeMs) return counselorsCache.rows;

  const csv = fs.readFileSync(COUNSELORS_CSV_PATH, "utf-8");
  const rows = parseCsv(csv, { columns: true, skip_empty_lines: true, bom: true });
  const mapped = rows.map((r) => ({
    employeeId: normalizeStr(r["Employee ID"] || r.EmployeeID || r.EmployeeId),
    name: normalizeStr(r.Name),
    email: normalizeStr(r["Email ID (Official)"] || r.Email || r.EmailID),
    region: normalizeStr(r.Region),
    subRegion: normalizeStr(r["Sub Region"] || r.SubRegion),
    designation: normalizeStr(r.Designation),
    roles: normalizeStr(r.Roles),
  })).filter((r) => r.name || r.email || r.employeeId);

  counselorsCache.mtimeMs = stat.mtimeMs;
  counselorsCache.rows = mapped;
  return mapped;
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
const FX_TIMEOUT_MS = Number(process.env.FX_TIMEOUT_MS || 20000);
let lastFxFetchedAt = null;
function todayKey() { return new Date().toISOString().slice(0, 10); }

async function getFx(from, toCsv, manualFx) {
  const key = `${todayKey()}|${from}|${toCsv}`;
  if (fxCache.has(key)) return fxCache.get(key);

  const url = `https://api.frankfurter.app/latest?from=${encodeURIComponent(from)}&to=${encodeURIComponent(toCsv)}`;

  const attempt = async () => {
    const data = await fetchJson(url, FX_TIMEOUT_MS);
    data.fetchedAt = new Date().toISOString();
    lastFxFetchedAt = data.fetchedAt;
    fxCache.set(key, data);
    return data;
  };

  try {
    return await attempt();
  } catch (e) {
    // Manual FX fallback (INR-GBP only)
    const enabled = Boolean(manualFx && manualFx.enabled);
    const inrPerGbp = Number((manualFx && manualFx.inrPerGbp) || 0);
    if (enabled && inrPerGbp > 0) {
      const f = String(from).toUpperCase();
      const t = String(toCsv).toUpperCase();
      if (f === "INR" && t === "GBP") {
        return { rates: { GBP: 1 / inrPerGbp }, base: "INR", date: null, fetchedAt: new Date().toISOString(), fxSource: "manual" };
      }
      if (f === "GBP" && t === "INR") {
        return { rates: { INR: inrPerGbp }, base: "GBP", date: null, fetchedAt: new Date().toISOString(), fxSource: "manual" };
      }
    }

    try {
      return await attempt();
    } catch (e2) {
      throw new Error("FX rates could not be fetched (timeout). Please check internet/firewall or use the optional INR per GBP override.");
    }
  }
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
      const fx = await getFx(currency, "GBP", payload.manualFx);
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

app.get("/api/students", (req, res) => {
  const q = normalizeStr(req.query.q).toLowerCase();
  if (!q) return res.json({ items: [] });
  const rows = loadStudents();
  const items = rows.filter((r) =>
    String(r.ackNumber || "").toLowerCase().includes(q) ||
    String(r.studentName || "").toLowerCase().includes(q)
  ).slice(0, 10);
  res.json({ items });
});

app.get("/api/counselors", (req, res) => {
  const q = normalizeStr(req.query.q).toLowerCase();
  if (!q) return res.json({ items: [] });
  const rows = loadCounselors();
  const items = rows.filter((r) =>
    String(r.name || "").toLowerCase().includes(q) ||
    String(r.email || "").toLowerCase().includes(q) ||
    String(r.employeeId || "").toLowerCase().includes(q)
  ).slice(0, 10);
  res.json({ items });
});

app.get("/api/fx", async (req, res) => {
  try {
    const from = String(req.query.from || "GBP").toUpperCase();
    const to = String(req.query.to || "INR").toUpperCase();
    if (from === to) return res.json({ provider: "frankfurter.app", from, to, rate: 1 });

    const manualFx = {
      enabled: String(req.query.manual_enabled || "").toLowerCase() === "true",
      inrPerGbp: Number(req.query.manual_inr_per_gbp || 0)
    };
    const data = await getFx(from, to, manualFx);
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
    const pdfMode = String(payload.pdfMode || "internal");
    if (!payload.applicationDate) return res.status(400).json({ error: "Visa application date is required." });
    const fundsReq = calcFundsRequired(payload, config);
    const fundsAvail = await calcFundsAvailable(payload, config.rules);
    const gapEligible = round2(fundsAvail.summary.totalEligibleGbp - fundsReq.fundsRequiredGbp);
    const gapAll = round2(fundsAvail.summary.totalAllGbp - fundsReq.fundsRequiredGbp);

    res.json({
      meta: { version: "2.3.0", fxFetchedAt: lastFxFetchedAt },
      fundsRequired: fundsReq,
      fundsAvailable: fundsAvail,
      gapGbp: gapEligible,
      gapAllGbp: gapAll,
      gapEligibleOnlyGbp: gapEligible,
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
    const pdfMode = String(payload.pdfMode || "internal");
    if (!payload.applicationDate) return res.status(400).json({ error: "Visa application date is required." });
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
        const fx = await getFx("GBP", quote, payload.manualFx);
        gbpToQuote = Number(fx?.rates?.[quote]) || 1;
      } catch (_) {
        gbpToQuote = 1;
      }
    }
    const fmt = (n) => Number(n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const fmtGBP = (n) => `GBP ${fmt(n)}`;
    const fmtQuote = (n) => (quote === "GBP" ? fmtGBP(n) : `${quote} ${fmt(n * gbpToQuote)}`);

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="UK_Visa_IHS_Funds_Report.pdf"`);

    const doc = new PDFDocument({ size: "A4", margin: 42 });
    doc.pipe(res);

    const b = config.branding || {};
    const c1 = b.primary_color || "#0b5cab";
    const c2 = b.secondary_color || "#00a3a3";
    const margin = 42;
    const pageWidth = doc.page.width - margin * 2;
    const headerHeight = 70;
    const maxPages = pdfMode === "client" ? 1 : 2;
    let pageCount = 1;

    const fitText = (text, maxWidth) => {
      const raw = String(text ?? "-");
      if (doc.widthOfString(raw) <= maxWidth) return raw;
      let s = raw;
      while (s.length > 0 && doc.widthOfString(`${s}...`) > maxWidth) s = s.slice(0, -1);
      return s ? `${s}...` : "";
    };

    const drawValue = (text, x, yPos, width, baseSize = 9, minSize = 7) => {
      const raw = String(text ?? "-");
      let size = baseSize;
      doc.fontSize(size);
      while (size > minSize && doc.widthOfString(raw) > width) {
        size -= 0.5;
        doc.fontSize(size);
      }
      doc.text(raw, x, yPos, { width, lineBreak: false });
      doc.fontSize(baseSize);
    };

    const drawHeader = () => {
      doc.rect(0, 0, doc.page.width, headerHeight).fill(c1);
      doc.rect(0, headerHeight - 14, doc.page.width, 14).fill(c2);

      doc.roundedRect(margin, 12, 140, 38, 6).fill("#ffffff");
      try {
        const logoRel = (b.logo_pdf || "public/assets/kc_logo.png").replace(/^\/+/, "");
        const logoPath = path.join(__dirname, logoRel);
        doc.image(logoPath, margin + 10, 18, { height: 26 });
      } catch (_) {
        // ignore logo errors
      }

      doc.fillColor("#ffffff").fontSize(15).text(b.product_name || "UK Visa Calculation Report", margin, 16, { width: pageWidth, align: "center" });
      doc.fontSize(9).text(b.company_name || "", margin, 34, { width: pageWidth, align: "center" });
      doc.fontSize(8).fillColor("#073344").text(`Generated: ${new Date().toLocaleString()} - Mode: ${pdfMode === "client" ? "Client" : "Internal"}`, margin, headerHeight - 12, { width: pageWidth, align: "right" });

      const stampText = ok ? "ELIGIBLE" : "NOT ELIGIBLE";
      doc.save();
      doc.rotate(-18, { origin: [470, 120] });
      doc.fontSize(26).fillColor(ok ? "#16a34a" : "#dc2626").opacity(0.18)
        .text(stampText, 360, 95, { width: 200, align: "center" });
      doc.opacity(1).restore();

      return headerHeight + 14;
    };

    let y = drawHeader();
    const footerSpace = 28;
    const bottomLimit = () => doc.page.height - margin - footerSpace;

    const ensureSpace = (height) => {
      if (y + height <= bottomLimit()) return true;
      if (pageCount < maxPages) {
        doc.addPage();
        pageCount += 1;
        y = drawHeader();
        return true;
      }
      return false;
    };

    const drawSummary = () => {
      const boxH = 44;
      if (!ensureSpace(boxH + 6)) return;
      doc.roundedRect(margin, y, pageWidth, boxH, 10).fill(ok ? "#dcfce7" : "#fee2e2");
      doc.fillColor(ok ? "#065f46" : "#7f1d1d").fontSize(11)
        .text(ok ? "Result: ELIGIBLE (Funds are sufficient)" : "Result: NOT ELIGIBLE (Funds are short)", margin + 10, y + 8, { width: pageWidth - 20 });
      doc.fillColor("#0f172a").fontSize(9)
        .text(`Required: ${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})`, margin + 10, y + 24, { width: pageWidth - 20 });
      doc.fillColor("#0f172a").fontSize(9)
        .text(`Eligible available: ${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})`, margin + 10, y + 34, { width: pageWidth - 20 });
      y += boxH + 8;
      doc.fillColor("#000000");
    };

    const drawBox = (x, boxY, w, h, title, rows, columns = 1) => {
      doc.roundedRect(x, boxY, w, h, 8).fill("#f8fafc");
      doc.strokeColor("#e2e8f0").lineWidth(1).roundedRect(x, boxY, w, h, 8).stroke();
      doc.fillColor("#0f172a").fontSize(10).text(title, x + 8, boxY + 6);

      const paddingTop = 20;
      const rowHeight = 12;
      const colGap = 12;
      const cols = Math.max(1, columns);
      const colWidth = (w - 16 - colGap * (cols - 1)) / cols;
      const labelSize = 7;
      const valueSize = 9;
      rows.forEach((row, i) => {
        const col = i % cols;
        const rowIdx = Math.floor(i / cols);
        const rx = x + 8 + col * (colWidth + colGap);
        const ry = boxY + paddingTop + rowIdx * rowHeight;
        doc.fontSize(labelSize).fillColor("#64748b").text(`${row.label}: `, rx, ry, { continued: true });
        doc.fontSize(valueSize).fillColor("#0f172a");
        drawValue(fitText(row.value ?? "-", colWidth - 8), rx, ry, colWidth - 8, valueSize, 7);
      });
    };

    const drawInlineBox = (title, items) => {
      const boxH = 28;
      if (!ensureSpace(boxH + 6)) return;
      doc.roundedRect(margin, y, pageWidth, boxH, 8).fill("#f8fafc");
      doc.strokeColor("#e2e8f0").lineWidth(1).roundedRect(margin, y, pageWidth, boxH, 8).stroke();
      doc.fillColor("#0f172a").fontSize(9).text(title, margin + 8, y + 6);
      const colGap = 12;
      const startX = margin + 90;
      const colWidth = (pageWidth - 98 - colGap * (items.length - 1)) / items.length;
      items.forEach((item, idx) => {
        const x = startX + idx * (colWidth + colGap);
        doc.fontSize(7).fillColor("#64748b").text(`${item.label}:`, x, y + 6);
        doc.fontSize(8).fillColor("#0f172a");
        drawValue(fitText(item.value ?? "-", colWidth), x, y + 16, colWidth, 8, 7);
      });
      y += boxH + 8;
      doc.fillColor("#000000");
    };

    const drawFundsTable = (rows) => {
      const headerH = 16;
      const rowH = 10;
      const maxRows = 10;
      const cols = [
        { title: "OK", width: 24 },
        { title: "Account", width: 70 },
        { title: "Currency", width: 50 },
        { title: "Amount", width: 60 },
        { title: "GBP", width: 60 },
        { title: "Issues", width: pageWidth - (24 + 70 + 50 + 60 + 60) - 10 },
      ];

      if (!ensureSpace(headerH + rowH + 4)) return;
      doc.roundedRect(margin, y, pageWidth, headerH, 6).fill("#f1f5f9");
      doc.fillColor("#334155").fontSize(7);
      let x = margin + 6;
      cols.forEach((c) => {
        doc.text(c.title, x, y + 4, { width: c.width });
        x += c.width;
      });
      y += headerH + 2;

      doc.fontSize(7);
      rows.slice(0, maxRows).forEach((r) => {
        if (!ensureSpace(rowH + 2)) return;
        let cx = margin + 6;
        const okTxt = r.eligible ? "OK" : "NO";
        doc.fillColor(r.eligible ? "#065f46" : "#9a3412").text(okTxt, cx, y, { width: cols[0].width });
        cx += cols[0].width;
        doc.fillColor("#0f172a").text(fitText(r.accountType || "-", cols[1].width - 4), cx, y, { width: cols[1].width });
        cx += cols[1].width;
        doc.text(fitText(r.currency || "-", cols[2].width - 4), cx, y, { width: cols[2].width });
        cx += cols[2].width;
        doc.text(fitText(String(r.amount || 0), cols[3].width - 4), cx, y, { width: cols[3].width });
        cx += cols[3].width;
        doc.text(fitText(`GBP ${String(r.amountGbp || 0)}`, cols[4].width - 4), cx, y, { width: cols[4].width });
        cx += cols[4].width;
        doc.fillColor("#475569").text(fitText((r.issues || []).join("; "), cols[5].width - 4), cx, y, { width: cols[5].width });
        y += rowH;
      });
    };

    drawSummary();

    const studentCityCountry = [payload.studentCity, payload.studentCountry].filter(Boolean).join(", ");
    const courseDates = `${payload.courseStart || "-"} to ${payload.courseEnd || "-"}`;

    const colGap = 12;
    const halfWidth = (pageWidth - colGap) / 2;
    const rowTop = y;
    const boxH1 = 86;
    if (ensureSpace(boxH1)) {
      drawBox(margin, rowTop, halfWidth, boxH1, "Course details", [
        { label: "University", value: payload.universityName || "-" },
        { label: "Study location", value: normalizeRegion(payload.region || "outside_london") === "london" ? "London" : "Outside London" },
        { label: "Course dates", value: courseDates },
        { label: "Visa application", value: payload.applicationDate || "-" },
        { label: "Display currency", value: quote },
        { label: "Dependants", value: fundsReq.dependantsCountEffective },
      ], 1);
      drawBox(margin + halfWidth + colGap, rowTop, halfWidth, boxH1, "Student details", [
        { label: "Acknowledgement No", value: payload.studentAckNumber || "-" },
        { label: "Student name", value: payload.studentName || "-" },
        { label: "Program", value: payload.studentProgram || "-" },
        { label: "Status", value: payload.studentStatus || "-" },
        { label: "Intake - InYear", value: payload.studentIntakeYear || "-" },
        { label: "City / Country", value: studentCityCountry || "-" },
      ], 1);
      y = rowTop + boxH1 + 8;
    }

    drawInlineBox("Counselor details", [
      { label: "Name", value: payload.counselorName || "-" },
      { label: "Email", value: payload.counselorEmail || "-" },
      { label: "Region/Sub", value: [payload.counselorRegion, payload.counselorSubRegion].filter(Boolean).join(" / ") || "-" },
      { label: "Role", value: [payload.counselorDesignation, payload.counselorRoles].filter(Boolean).join(" / ") || "-" },
    ]);

    const boxH2 = 88;
    const rowTop2 = y;
    if (ensureSpace(boxH2)) {
      drawBox(margin, rowTop2, halfWidth, boxH2, "Fees and IHS", [
        { label: "Tuition total", value: `${fmtGBP(safeNum(payload.tuitionFeeTotalGbp))} (${fmtQuote(safeNum(payload.tuitionFeeTotalGbp))})` },
        { label: "Tuition paid", value: `${fmtGBP(safeNum(payload.tuitionFeePaidGbp))} (${fmtQuote(safeNum(payload.tuitionFeePaidGbp))})` },
        { label: "Scholarship", value: `${fmtGBP(safeNum(payload.scholarshipGbp))} (${fmtQuote(safeNum(payload.scholarshipGbp))})` },
        { label: "Buffer", value: `${fmtGBP(safeNum(payload.bufferGbp))} (${fmtQuote(safeNum(payload.bufferGbp))})` },
        { label: "IHS per person", value: `${fmtGBP(ihs.ihsPerPersonGbp)} (${fmtQuote(ihs.ihsPerPersonGbp)})` },
        { label: "IHS total", value: `${fmtGBP(ihs.ihsTotalGbp)} (${fmtQuote(ihs.ihsTotalGbp)})` },
      ], 1);
      drawBox(margin + halfWidth + colGap, rowTop2, halfWidth, boxH2, "Funds required (28-day)", [
        { label: "Tuition due", value: `${fmtGBP(fundsReq.tuitionDueGbp)} (${fmtQuote(fundsReq.tuitionDueGbp)})` },
        { label: "Maintenance (student)", value: `${fmtGBP(fundsReq.maintenanceStudentGbp)} (${fmtQuote(fundsReq.maintenanceStudentGbp)})` },
        { label: "Maintenance (dependants)", value: `${fmtGBP(fundsReq.maintenanceDependantsGbp)} (${fmtQuote(fundsReq.maintenanceDependantsGbp)})` },
        { label: "Buffer", value: `${fmtGBP(fundsReq.bufferGbp)} (${fmtQuote(fundsReq.bufferGbp)})` },
      ], 1);
      y = rowTop2 + boxH2 + 8;
    }

    drawInlineBox("Totals", [
      { label: "Total required", value: `${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})` },
      { label: "Eligible funds", value: `${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})` },
      { label: "Gap (Eligible-Required)", value: fmtGBP(gapEligible) + (quote === "GBP" ? "" : ` | ${fmtQuote(gapEligible)}`) },
    ]);

    if (pdfMode === "internal") {
      drawFundsTable(fundsAvail.rows);
      y += 6;
    }

    const ruleText =
      `Rules reminder: Funds must be held for ${config.rules.funds_hold_days} consecutive days. ` +
      `Statement end date must be within ${config.rules.statement_age_days} days of the visa application date. ` +
      `Visa application date is required for the statement freshness check.`;
    const rulesHeight = 22;
    if (y + rulesHeight > bottomLimit()) {
      y = bottomLimit() - rulesHeight;
    }
    doc.fillColor("#475569").fontSize(7).text(ruleText, margin, y, { width: pageWidth });
    y += 10;

    if (b.footer_note) {
      doc.fillColor("#64748b").fontSize(7).text(String(b.footer_note), margin, doc.page.height - 30, { width: pageWidth, align: "center" });
    }

    doc.end();
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

app.get("/healthz", (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

app.get("/health", (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
