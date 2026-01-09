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
const COUNTRY_CURRENCY_PATH = path.join(__dirname, "data", "country_currency.json");
const DATA_SYNC_ENABLED = String(process.env.DATA_SYNC_ENABLED || "true").toLowerCase() === "true";
const STUDENTS_SOURCE_URL = process.env.STUDENTS_SOURCE_URL || "";
const COUNSELORS_SOURCE_URL = process.env.COUNSELORS_SOURCE_URL || "";
const COUNTRY_CURRENCY_SOURCE_URL = process.env.COUNTRY_CURRENCY_SOURCE_URL || "";
const STUDENTS_SYNC_MS = Number(process.env.STUDENTS_SYNC_MS || 24 * 60 * 60 * 1000);
const COUNSELORS_SYNC_MS = Number(process.env.COUNSELORS_SYNC_MS || 20 * 24 * 60 * 60 * 1000);
const COUNTRY_CURRENCY_SYNC_MS = Number(process.env.COUNTRY_CURRENCY_SYNC_MS || 24 * 60 * 60 * 1000);
const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL || "";
const APP_VERSION = "2.3.0";

function readLocalConfig() {
  return JSON.parse(fs.readFileSync(LOCAL_CONFIG_PATH, "utf-8"));
}

function readCountryCurrency() {
  if (!fs.existsSync(COUNTRY_CURRENCY_PATH)) return {};
  return JSON.parse(fs.readFileSync(COUNTRY_CURRENCY_PATH, "utf-8"));
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
      intakeYear: normalizeStr(
        n.intake ||
        n["intake"] ||
        n["intake inyear"] ||
        n["intake - inyear"] ||
        n["intake in year"] ||
        n["intake year"] ||
        n["intake - in year"] ||
        n.intakeinyear ||
        n.intakeyear
      ),
      tuitionFeeTotalGbp: normalizeStr(
        n["tuition fee"] ||
        n["tuition fee gbp"] ||
        n["tuition fee total"] ||
        n["tuitionfeetotal"] ||
        n["course fee"] ||
        n["course fee gbp"] ||
        n["coursefee"] ||
        n["total fee"] ||
        n["total fee gbp"]
      ),
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

const syncState = new Map();
const syncMeta = new Map();
function formatIst(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "";
  const fmt = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "short",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit"
  });
  return `${fmt.format(d)} IST`;
}
async function sendTeamsNotification(message) {
  if (!TEAMS_WEBHOOK_URL) return;
  try {
    await fetch(TEAMS_WEBHOOK_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ text: message })
    });
  } catch (e) {
    console.warn(`Teams notify failed: ${e.message || e}`);
  }
}
async function fetchBinary(url, timeoutMs = 15000) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), timeoutMs);
  const state = syncState.get(url) || {};
  try {
    const headers = {};
    if (state.etag) headers["if-none-match"] = state.etag;
    if (state.lastModified) headers["if-modified-since"] = state.lastModified;
    const res = await fetch(url, { signal: ctrl.signal, headers });
    if (res.status === 304) return { notModified: true };
    if (!res.ok) throw new Error(`Fetch failed (${res.status})`);
    const arrayBuffer = await res.arrayBuffer();
    return {
      buffer: Buffer.from(arrayBuffer),
      etag: res.headers.get("etag") || "",
      lastModified: res.headers.get("last-modified") || ""
    };
  } finally {
    clearTimeout(t);
  }
}

async function downloadToFile(name, url, destPath) {
  const state = syncState.get(url) || {};
  if (state.running) return;
  state.running = true;
  syncState.set(url, state);
  const meta = syncMeta.get(name) || {};
  meta.lastAttemptAt = new Date().toISOString();
  syncMeta.set(name, meta);
  try {
    const res = await fetchBinary(url);
    if (res.notModified) {
      state.running = false;
      syncState.set(url, state);
      meta.lastCheckedAt = new Date().toISOString();
      syncMeta.set(name, meta);
      await sendTeamsNotification(
        `Sync OK (no changes): ${name}\nChecked: ${formatIst(meta.lastCheckedAt)}`
      );
      return;
    }
    const tmpPath = `${destPath}.tmp-${Date.now()}`;
    await fs.promises.mkdir(path.dirname(destPath), { recursive: true });
    await fs.promises.writeFile(tmpPath, res.buffer);
    await fs.promises.rename(tmpPath, destPath);
    syncState.set(url, {
      etag: res.etag,
      lastModified: res.lastModified,
      running: false
    });
    syncMeta.set(name, {
      lastAttemptAt: meta.lastAttemptAt,
      lastCheckedAt: new Date().toISOString(),
      lastSuccessAt: new Date().toISOString(),
      lastError: ""
    });
    console.log(`Synced ${url} -> ${destPath}`);
    await sendTeamsNotification(
      `Sync OK: ${name}\nUpdated: ${formatIst(new Date().toISOString())}`
    );
  } catch (e) {
    state.running = false;
    syncState.set(url, state);
    syncMeta.set(name, {
      lastAttemptAt: meta.lastAttemptAt,
      lastCheckedAt: new Date().toISOString(),
      lastSuccessAt: meta.lastSuccessAt || "",
      lastError: String(e?.message || e)
    });
    console.warn(`Sync failed for ${url}: ${e.message || e}`);
    await sendTeamsNotification(
      `Sync FAILED: ${name}\nChecked: ${formatIst(new Date().toISOString())}\nError: ${String(e?.message || e)}`
    );
  }
}

function scheduleSync(name, url, destPath, intervalMs) {
  if (!DATA_SYNC_ENABLED) return;
  if (!url) return;
  if (!Number.isFinite(intervalMs) || intervalMs <= 0) return;
  console.log(`Sync enabled for ${name} every ${Math.round(intervalMs / 60000)}m`);
  downloadToFile(name, url, destPath);
  const timer = setInterval(() => downloadToFile(name, url, destPath), intervalMs);
  if (timer.unref) timer.unref();
}

function fileStatus(pathname) {
  if (!fs.existsSync(pathname)) return { exists: false, updatedAt: "" };
  const stat = fs.statSync(pathname);
  return { exists: true, updatedAt: new Date(stat.mtimeMs).toISOString() };
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

  const tryManualOverride = () => {
    const overrides = manualFx && typeof manualFx === "object" ? manualFx.overrides : null;
    const f = String(from || "").toUpperCase();
    const t = String(toCsv || "").toUpperCase();
    if (overrides && overrides[f] && t === "GBP") {
      return { rates: { GBP: Number(overrides[f]) }, base: f, date: null, fetchedAt: new Date().toISOString(), fxSource: "manual_override" };
    }
    if (overrides && overrides[t] && f === "GBP") {
      return { rates: { [t]: 1 / Number(overrides[t]) }, base: "GBP", date: null, fetchedAt: new Date().toISOString(), fxSource: "manual_override" };
    }
    return null;
  };

  try {
    return await attempt();
  } catch (e) {
    const override = tryManualOverride();
    if (override) return override;
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

function addMonthsExcel(dateStr, months){
  if (!dateStr) return null;
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return null;
  const year = d.getFullYear();
  const month = d.getMonth();
  const day = d.getDate();
  const target = new Date(year, month + Number(months || 0), 1);
  const lastDay = new Date(target.getFullYear(), target.getMonth() + 1, 0).getDate();
  target.setDate(Math.min(day, lastDay));
  return target;
}

function datedifMonths(startDate, endDate){
  if (!startDate || !endDate) return 0;
  const s = new Date(startDate);
  const e = new Date(endDate);
  if (Number.isNaN(s.getTime()) || Number.isNaN(e.getTime())) return 0;
  let months = (e.getFullYear() - s.getFullYear()) * 12 + (e.getMonth() - s.getMonth());
  if (e.getDate() < s.getDate()) months -= 1;
  return Math.max(0, months);
}

function formatDateISO(d){
  if (!d || Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

function todayISOInIST(){
  const parts = new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Kolkata", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(new Date());
  const y = parts.find(p => p.type === "year")?.value || "0000";
  const m = parts.find(p => p.type === "month")?.value || "01";
  const d = parts.find(p => p.type === "day")?.value || "01";
  return `${y}-${m}-${d}`;
}

const dateFormatter = new Intl.DateTimeFormat("en-IN", { day: "2-digit", month: "short", year: "numeric", timeZone: "Asia/Kolkata" });
const dateTimeFormatter = new Intl.DateTimeFormat("en-IN", {
  day: "2-digit",
  month: "short",
  year: "numeric",
  hour: "numeric",
  minute: "2-digit",
  hour12: true,
  timeZone: "Asia/Kolkata"
});
function formatDateDisplay(value){
  if (!value) return "-";
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return "-";
  return dateFormatter.format(d);
}
function formatDateTimeDisplay(value){
  if (!value) return "-";
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return "-";
  return dateTimeFormatter.format(d);
}

function sanitizeFilenamePart(val){
  return String(val || "")
    .replace(/[^\w\s()-]/g, "")
    .trim()
    .replace(/\s+/g, "_");
}

function buildPdfFilename(payload){
  const name = sanitizeFilenamePart(payload.studentName) || "Student";
  const uni = sanitizeFilenamePart(payload.universityName) || "University";
  const intake = sanitizeFilenamePart(payload.studentIntakeYear) || "Intake";
  const ack = sanitizeFilenamePart(payload.studentAckNumber) || "ACK";
  return `${name}_${uni}_${intake}_(${ack}).pdf`;
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
  // rows: [{fundType, accountType, source, currency, amount, statementStart, statementEnd, fdMaturity, loanDisbursement}]
  const rows = Array.isArray(payload.fundsRows) ? payload.fundsRows : [];
  const applicationDate = payload.applicationDate ? new Date(payload.applicationDate) : null;

  const fundsHoldDays = Number(rules?.funds_hold_days ?? 28);
  const statementAgeDays = Number(rules?.statement_age_days ?? 31);
  const loanLetterMaxAgeDays = Number(rules?.loan_letter_max_age_days ?? 180);
  const skipFunds = Boolean(payload.fundsSkip);

  if (skipFunds) {
    return {
      summary: {
        totalAllGbp: 0,
        totalEligibleGbp: 0,
        fundsHoldDays,
        statementAgeDays,
        loanLetterMaxAgeDays,
        hasApplicationDate: !!applicationDate,
        anyRowMissingDates: false,
        anyIneligibleRows: false,
        skipped: true,
        skipReason: "User marked funds as not held yet"
      },
      rows: []
    };
  }

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

  const isAppDateWarning = (msg) => msg && msg.startsWith("No application date");

  for (const r of rows) {
    const fundType = String(r.fundType || "bank").toLowerCase();
    const accountType = String(r.accountType || "Student");
    const source = String(r.source || "");
    const currency = String(r.currency || "GBP").toUpperCase().trim();
    const amount = safeNum(r.amount);
    const statementStart = r.statementStart || "";
    const statementEnd = r.statementEnd || "";
    const fdMaturity = r.fdMaturity || "";
    const loanDisbursement = r.loanDisbursement || "";

    if (amount <= 0) continue;

    // FX convert
    let gbp = amount;
    let rate = 1;
    if (currency !== "GBP") {
      try {
        const fx = await getFx(currency, "GBP", payload.manualFx);
        rate = safeNum(fx?.rates?.GBP);
        gbp = rate ? amount * rate : 0;
        if (!rate) issues.push(`FX rate unavailable for ${currency}`);
      } catch (e) {
        rate = 0;
        gbp = 0;
        issues.push(`FX unavailable for ${currency}`);
      }
    }

    totalAllGbp += gbp;

    // eligibility checks
    const issues = [];
    let dateLabel = "";
    let dateValue = "";

    if (fundType === "fd") {
      dateLabel = "FD maturity";
      dateValue = formatDateDisplay(fdMaturity);
      const mDt = fdMaturity ? new Date(fdMaturity) : null;
      if (!mDt || Number.isNaN(mDt.getTime())) {
        issues.push("Missing/invalid FD maturity date");
      }
    } else if (fundType === "loan") {
      dateLabel = "Loan disbursement";
      dateValue = formatDateDisplay(loanDisbursement);
      const dDt = loanDisbursement ? new Date(loanDisbursement) : null;
      if (!dDt || Number.isNaN(dDt.getTime())) {
        issues.push("Missing/invalid loan disbursement letter date");
      } else if (applicationDate && loanLetterMaxAgeDays > 0) {
        const age = daysSince(dDt, applicationDate);
        if (age < 0) issues.push("Loan letter date is after application date");
        else if (age > loanLetterMaxAgeDays) issues.push(`Loan letter is ${age} days before application (> ${loanLetterMaxAgeDays})`);
      } else if (!applicationDate) {
        issues.push("No application date (loan letter age not checked)");
      }
    } else {
      dateLabel = "Statement";
      const startLabel = formatDateDisplay(statementStart);
      const endLabel = formatDateDisplay(statementEnd);
      dateValue = `${startLabel} - ${endLabel}`;
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
    }

    const eligible = issues.length === 0 || issues.every(isAppDateWarning);
    if (eligible) totalEligibleGbp += gbp;

    converted.push({
      fundType,
      accountType,
      source,
      currency,
      amount: round2(amount),
      statementStart,
      statementEnd,
      fdMaturity,
      loanDisbursement,
      dateLabel,
      dateValue,
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
    loanLetterMaxAgeDays,
    hasApplicationDate: !!applicationDate,
    anyRowMissingDates: converted.some((r) => {
      if (r.fundType === "fd") return !r.fdMaturity;
      if (r.fundType === "loan") return !r.loanDisbursement;
      return !r.statementStart || !r.statementEnd;
    }),
    anyIneligibleRows: converted.some(r => !r.eligible),
    skipped: false
  };

  return { summary, rows: converted };
}

function computeIhsBlock(payload, config, dependantsEffective){
  const courseStart = payload.courseStart;
  const courseEnd = payload.courseEnd;
  const visaEndDate = addMonthsExcel(courseEnd, 4);
  const visaEndIso = formatDateISO(visaEndDate);
  const applyingFrom = payload.applyingFrom || "outside";

  let totalStayDays = 0;
  if (courseStart && visaEndDate) {
    const cs = new Date(courseStart);
    if (!Number.isNaN(cs.getTime())) {
      totalStayDays = Math.floor((visaEndDate - cs) / 86400000) + 1;
    }
  }
  const datedifM = datedifMonths(courseStart, visaEndIso);
  const startDay = courseStart ? new Date(courseStart).getDate() : 0;
  const endDay = visaEndDate ? visaEndDate.getDate() : 0;
  const totalStayMonths = Math.max(0, datedifM + (endDay > startDay ? 1 : 0));

  const chargeableUnits = Math.max(0, Math.ceil(totalStayMonths / 6));
  const fullYears = Math.floor(chargeableUnits / 2);
  const halfYearCharges = chargeableUnits % 2;
  const yearlyCharges = fullYears;

  const ihsPerPerson = round2((fullYears * config.ihs.student_yearly_gbp) + (halfYearCharges * config.ihs.half_year_gbp));
  const persons = 1 + Math.max(0, Math.floor(safeNum(payload.ihsDependantsCount ?? dependantsEffective ?? 0)));

  return {
    ihsPerPersonGbp: round2(ihsPerPerson),
    persons,
    ihsTotalGbp: round2(ihsPerPerson * persons),
    visaEndDate: visaEndIso,
    totalStayDays: Math.max(0, totalStayDays),
    totalStayMonths,
    chargeableUnits,
    fullYears,
    yearlyCharges,
    halfYearCharges,
    rateYearlyGbp: round2(config.ihs.student_yearly_gbp),
    rateHalfGbp: round2(config.ihs.half_year_gbp),
    applyingFrom
  };
}

// --- APIs ---
app.get("/api/config", async (req, res) => {
  try {
    const local = readLocalConfig();
    const countryCurrency = readCountryCurrency();
    if (!CONFIG_URL) {
      return res.json({ config: { ...local, country_currency: countryCurrency }, source: "local" });
    }

    const remote = await fetchJson(CONFIG_URL, 8000);
    const merged = {
      ...local,
      ...remote,
      routes: { ...local.routes, ...(remote.routes || {}) },
      rules: { ...local.rules, ...(remote.rules || {}) },
      ihs: { ...local.ihs, ...(remote.ihs || {}) },
      fees: { ...local.fees, ...(remote.fees || {}) },
      fx: { ...local.fx, ...(remote.fx || {}) },
      universities: remote.universities || local.universities,
      country_currency: { ...(local.country_currency || {}), ...(remote.country_currency || {}), ...countryCurrency }
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

app.get("/api/sync-status", (req, res) => {
  const studentsMeta = syncMeta.get("students") || {};
  const counselorsMeta = syncMeta.get("counselors") || {};
  const countryMeta = syncMeta.get("country_currency") || {};
  res.json({
    serverTime: new Date().toISOString(),
    students: { ...fileStatus(STUDENTS_XLSX_PATH), ...studentsMeta },
    counselors: { ...fileStatus(COUNSELORS_CSV_PATH), ...counselorsMeta },
    country_currency: { ...fileStatus(COUNTRY_CURRENCY_PATH), ...countryMeta }
  });
});

app.post("/api/sync", async (req, res) => {
  try {
    const targetsRaw = normalizeStr(req.query.targets || "all").toLowerCase();
    const wantAll = targetsRaw === "all";
    const wants = new Set(targetsRaw.split(",").map((t) => t.trim()).filter(Boolean));
    const tasks = [];
    const picked = [];

    if (STUDENTS_SOURCE_URL && (wantAll || wants.has("students"))) {
      tasks.push(downloadToFile("students", STUDENTS_SOURCE_URL, STUDENTS_XLSX_PATH));
      picked.push("students");
    }
    if (COUNSELORS_SOURCE_URL && (wantAll || wants.has("counselors"))) {
      tasks.push(downloadToFile("counselors", COUNSELORS_SOURCE_URL, COUNSELORS_CSV_PATH));
      picked.push("counselors");
    }
    if (COUNTRY_CURRENCY_SOURCE_URL && (wantAll || wants.has("country_currency"))) {
      tasks.push(downloadToFile("country_currency", COUNTRY_CURRENCY_SOURCE_URL, COUNTRY_CURRENCY_PATH));
      picked.push("country_currency");
    }

    if (!tasks.length) {
      return res.status(400).json({ error: "No valid sync targets configured." });
    }

    await Promise.all(tasks);
    return res.json({ ok: true, targets: picked });
  } catch (e) {
    return res.status(500).json({ error: String(e) });
  }
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

app.post("/api/ihs", (req, res) => {
  try {
    const config = readLocalConfig();
    const payload = req.body || {};
    if (!payload.courseStart || !payload.courseEnd) {
      return res.status(400).json({ error: "courseStart and courseEnd required." });
    }
    const ihs = computeIhsBlock(payload, config, 0);
    res.json({ ihs });
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

app.post("/api/report", async (req, res) => {
  try {
    const config = readLocalConfig();
    const payload = req.body || {};
    if (!payload.applicationDate) {
      payload.applicationDate = todayISOInIST();
      payload.applicationDateDefaulted = true;
    }
    const pdfMode = String(payload.pdfMode || "internal");
    const fundsReq = calcFundsRequired(payload, config);
    const fundsAvail = await calcFundsAvailable(payload, config.rules);
    const gapEligible = round2(fundsAvail.summary.totalEligibleGbp - fundsReq.fundsRequiredGbp);
    const gapAll = round2(fundsAvail.summary.totalAllGbp - fundsReq.fundsRequiredGbp);

    res.json({
      meta: { version: APP_VERSION, fxFetchedAt: lastFxFetchedAt },
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
    if (!payload.applicationDate) {
      payload.applicationDate = todayISOInIST();
      payload.applicationDateDefaulted = true;
    }
    const pdfMode = String(payload.pdfMode || "internal");
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
    const currencyLocale = (cur) => {
      const c = String(cur || "").toUpperCase();
      if (c === "INR") return "en-IN";
      if (c === "GBP") return "en-GB";
      if (c === "EUR") return "de-DE";
      return "en-US";
    };
    const fmtCurrency = (cur, n) => new Intl.NumberFormat(currencyLocale(cur), {
      style: "currency",
      currency: String(cur || "GBP").toUpperCase(),
      currencySign: "accounting",
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    }).format(Number(n || 0));
    const fmtGBP = (n) => fmtCurrency("GBP", n);
    const fmtQuote = (n) => (quote === "GBP" ? fmtGBP(n) : fmtCurrency(quote, n * gbpToQuote));

    res.setHeader("Content-Type", "application/pdf");
    const pdfName = buildPdfFilename(payload);
    res.setHeader("Content-Disposition", `attachment; filename="${pdfName}"`);

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

    const applyPdfFont = () => {
      const fontRel = (b.pdf_font_path || "public/assets/fonts/NotoSans-Regular.ttf").replace(/^\/+/, "");
      const fontPath = path.join(__dirname, fontRel);
      if (fs.existsSync(fontPath)) {
        doc.registerFont("body", fontPath);
        doc.font("body");
      }
    };
    applyPdfFont();

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
      const fxStamp = lastFxFetchedAt ? `FX: ${formatDateTimeDisplay(lastFxFetchedAt)}` : "FX: -";
      const metaStamp = `v${APP_VERSION} | ${fxStamp}`;
      doc.fontSize(8).fillColor("#073344")
        .text(`Generated: ${formatDateTimeDisplay(new Date())} - Mode: ${pdfMode === "client" ? "Client" : "Internal"}`, margin, headerHeight - 22, { width: pageWidth, align: "right" });
      doc.fontSize(8).fillColor("#073344")
        .text(metaStamp, margin, headerHeight - 10, { width: pageWidth, align: "right" });

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
        .text(ok ? "Result: ELIGIBLE (Funds are sufficient)" : "Result: NOT ELIGIBLE (Funds are short)", margin + 10, y + 6, { width: pageWidth - 20 });

      const colW = (pageWidth - 20) / 2;
      const leftX = margin + 10;
      const rightX = margin + 10 + colW;
      doc.fillColor("#0f172a").fontSize(9)
        .text(`Required: ${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})`, leftX, y + 24, { width: colW - 6 });
      doc.fillColor("#0f172a").fontSize(9)
        .text(`Eligible available: ${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})`, rightX, y + 24, { width: colW - 6 });
      y += boxH + 8;
      doc.fillColor("#000000");
    };

    const measureKeyValueBoxHeight = (rows, w, labelRatio = 0.58) => {
      const headerH = 18;
      const paddingX = 10;
      const labelW = Math.floor((w - paddingX * 2) * labelRatio);
      const valueW = (w - paddingX * 2) - labelW;
      let height = headerH + 6;
      rows.forEach((row) => {
        const label = String(row.label ?? "-");
        const value = String(row.value ?? "-");
        const lh = doc.heightOfString(label, { width: labelW });
        const vh = doc.heightOfString(value, { width: valueW, align: "right" });
        height += Math.max(lh, vh) + 4;
      });
      return height;
    };

    const drawKeyValueBox = (x, boxY, w, h, title, rows, labelRatio = 0.58) => {
      doc.roundedRect(x, boxY, w, h, 8).fill("#f8fafc");
      doc.strokeColor("#e2e8f0").lineWidth(1).roundedRect(x, boxY, w, h, 8).stroke();
      doc.fillColor("#0f172a").fontSize(10).text(title, x + 10, boxY + 6);

      const paddingX = 10;
      const headerH = 18;
      const labelW = Math.floor((w - paddingX * 2) * labelRatio);
      const valueW = (w - paddingX * 2) - labelW;
      let ry = boxY + headerH;
      rows.forEach((row) => {
        const label = String(row.label ?? "-");
        const value = String(row.value ?? "-");
        doc.fontSize(7).fillColor("#64748b").text(label, x + paddingX, ry, { width: labelW });
        doc.fontSize(9).fillColor(row.valueColor || "#0f172a").text(value, x + paddingX + labelW, ry, { width: valueW, align: "right" });
        const lh = doc.heightOfString(label, { width: labelW });
        const vh = doc.heightOfString(value, { width: valueW, align: "right" });
        ry += Math.max(lh, vh) + 4;
      });
    };

    const measureTwoColBoxHeight = (w, rows) => {
      const headerH = 18;
      const paddingX = 10;
      const colGap = 12;
      const colW = (w - paddingX * 2 - colGap) / 2;
      let height = headerH + 6;
      rows.forEach((r) => {
        const leftText = `${r.leftLabel}: ${r.leftValue || "-"}`;
        const rightText = `${r.rightLabel}: ${r.rightValue || "-"}`;
        const lh = doc.heightOfString(leftText, { width: colW });
        const rh = doc.heightOfString(rightText, { width: colW });
        height += Math.max(lh, rh) + 4;
      });
      return height;
    };

    const drawTwoColBox = (x, boxY, w, title, rows) => {
      const headerH = 18;
      const paddingX = 10;
      const colGap = 12;
      const colW = (w - paddingX * 2 - colGap) / 2;
      const height = measureTwoColBoxHeight(w, rows);

      doc.roundedRect(x, boxY, w, height, 8).fill("#f8fafc");
      doc.strokeColor("#e2e8f0").lineWidth(1).roundedRect(x, boxY, w, height, 8).stroke();
      doc.fillColor("#0f172a").fontSize(10).text(title, x + 10, boxY + 6);

      let ry = boxY + headerH;
      rows.forEach((r) => {
        const leftText = `${r.leftLabel}: ${r.leftValue || "-"}`;
        const rightText = `${r.rightLabel}: ${r.rightValue || "-"}`;
        doc.fontSize(8).fillColor("#0f172a").text(leftText, x + paddingX, ry, { width: colW });
        doc.text(rightText, x + paddingX + colW + colGap, ry, { width: colW });
        const lh = doc.heightOfString(leftText, { width: colW });
        const rh = doc.heightOfString(rightText, { width: colW });
        ry += Math.max(lh, rh) + 4;
      });
    };

    const drawFundsTable = (rows) => {
      const headerH = 14;
      const rowH = 9;
      const maxRows = 10;
      const cols = [
        { title: "OK", width: 20 },
        { title: "Type", width: 70 },
        { title: "Owner", width: 70 },
        { title: "Amount", width: 90 },
        { title: "GBP", width: 90 },
        { title: "Date", width: 140 },
        { title: "Issues", width: pageWidth - (20 + 70 + 70 + 90 + 90 + 140) - 10 },
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

      const fundTypeLabel = (t) => {
        const ft = String(t || "bank").toLowerCase();
        if (ft === "fd") return "FD";
        if (ft === "loan") return "Loan";
        return "Bank";
      };

      doc.fontSize(7);
      rows.slice(0, maxRows).forEach((r) => {
        if (!ensureSpace(rowH + 2)) return;
        let cx = margin + 6;
        const okTxt = r.eligible ? "OK" : "NO";
        doc.fillColor(r.eligible ? "#065f46" : "#9a3412").text(okTxt, cx, y, { width: cols[0].width, align: "center" });
        cx += cols[0].width;
        doc.fillColor("#0f172a").text(fitText(fundTypeLabel(r.fundType), cols[1].width - 4), cx, y, { width: cols[1].width, align: "center" });
        cx += cols[1].width;
        doc.text(fitText(r.accountType || "-", cols[2].width - 4), cx, y, { width: cols[2].width, align: "center" });
        cx += cols[2].width;
        doc.text(fitText(fmtCurrency(r.currency || "GBP", r.amount || 0), cols[3].width - 4), cx, y, { width: cols[3].width, align: "center" });
        cx += cols[3].width;
        doc.text(fitText(fmtGBP(r.amountGbp || 0), cols[4].width - 4), cx, y, { width: cols[4].width, align: "center" });
        cx += cols[4].width;
        doc.text(fitText(r.dateValue || "-", cols[5].width - 4), cx, y, { width: cols[5].width, align: "center" });
        cx += cols[5].width;
        doc.fillColor("#475569").text(fitText((r.issues || []).join("; "), cols[6].width - 4), cx, y, { width: cols[6].width });
        y += rowH;
      });
    };

    drawSummary();

    const studentCityCountry = [payload.studentCity, payload.studentCountry].filter(Boolean).join(", ");
    const courseDates = `${formatDateDisplay(payload.courseStart)} to ${formatDateDisplay(payload.courseEnd)}`;

    const colGap = 12;
    const halfWidth = (pageWidth - colGap) / 2;
    const rowTop = y;
    const courseRows = [
      { label: "University", value: payload.universityName || "-" },
      { label: "Study location", value: normalizeRegion(payload.region || "outside_london") === "london" ? "London" : "Outside London" },
      { label: "Course dates", value: courseDates },
      { label: "Visa application", value: formatDateDisplay(payload.applicationDate) },
      { label: "Display currency", value: quote },
    ];
    const studentRows = [
      { label: "Acknowledgement No", value: payload.studentAckNumber || "-" },
      { label: "Student name", value: payload.studentName || "-" },
      { label: "Program", value: payload.studentProgram || "-" },
      { label: "Status", value: payload.studentStatus || "-" },
      { label: "Intake", value: payload.studentIntakeYear || "-" },
      { label: "City / Country", value: studentCityCountry || "-" },
    ];
    const boxH1 = Math.max(
      measureKeyValueBoxHeight(courseRows, halfWidth),
      measureKeyValueBoxHeight(studentRows, halfWidth)
    );
    if (ensureSpace(boxH1)) {
      drawKeyValueBox(margin, rowTop, halfWidth, boxH1, "Course details", courseRows);
      drawKeyValueBox(margin + halfWidth + colGap, rowTop, halfWidth, boxH1, "Student details", studentRows);
      y = rowTop + boxH1 + 8;
    }

    const counselorRows = [
      { leftLabel: "Name", leftValue: payload.counselorName || "-", rightLabel: "Email", rightValue: payload.counselorEmail || "-" },
      { leftLabel: "Role", leftValue: [payload.counselorDesignation, payload.counselorRoles].filter(Boolean).join(" / ") || "-", rightLabel: "Region/Sub", rightValue: [payload.counselorRegion, payload.counselorSubRegion].filter(Boolean).join(" / ") || "-" },
    ];
    const counselorH = measureTwoColBoxHeight(pageWidth, counselorRows);
    if (ensureSpace(counselorH)) {
      drawTwoColBox(margin, y, pageWidth, "Counselor details", counselorRows);
      y += counselorH + 8;
    }

    const ihsParts = [];
    if (ihs.yearlyCharges > 0) ihsParts.push(`${fmtGBP(ihs.rateYearlyGbp)} x ${ihs.yearlyCharges} year${ihs.yearlyCharges > 1 ? "s" : ""}`);
    if (ihs.halfYearCharges > 0) ihsParts.push(`${fmtGBP(ihs.rateHalfGbp)} x ${ihs.halfYearCharges} half-year${ihs.halfYearCharges > 1 ? "s" : ""}`);
    const ihsCalcText = ihsParts.length ? `${ihsParts.join(" + ")} = ${fmtGBP(ihs.ihsPerPersonGbp)} per person` : "No IHS";

    const rowTop2 = y;
    const visaFeeGbp = safeNum(config.fees?.visa_application_fee_gbp);
    const feesRows = [
      { label: "Tuition total", value: `${fmtGBP(safeNum(payload.tuitionFeeTotalGbp))} (${fmtQuote(safeNum(payload.tuitionFeeTotalGbp))})` },
      { label: "Tuition paid", value: `${fmtGBP(safeNum(payload.tuitionFeePaidGbp))} (${fmtQuote(safeNum(payload.tuitionFeePaidGbp))})` },
      { label: "Scholarship", value: `${fmtGBP(safeNum(payload.scholarshipGbp))} (${fmtQuote(safeNum(payload.scholarshipGbp))})` },
      { label: "Buffer", value: `${fmtGBP(safeNum(payload.bufferGbp))} (${fmtQuote(safeNum(payload.bufferGbp))})` },
      { label: "Visa application fee", value: `${fmtGBP(visaFeeGbp)} (${fmtQuote(visaFeeGbp)})` },
      { label: "Visa end date", value: formatDateDisplay(ihs.visaEndDate) },
      { label: "Total stay", value: `${ihs.totalStayMonths} months (${ihs.totalStayDays} days)` },
      { label: "IHS calculation", value: ihsCalcText },
      { label: "IHS total", value: `${fmtGBP(ihs.ihsTotalGbp)} (${fmtQuote(ihs.ihsTotalGbp)})` },
    ];
    const fundsReqRows = [
      { label: "Tuition due", value: `${fmtGBP(fundsReq.tuitionDueGbp)} (${fmtQuote(fundsReq.tuitionDueGbp)})` },
      { label: "Maintenance (student)", value: `${fmtGBP(fundsReq.maintenanceStudentGbp)} (${fmtQuote(fundsReq.maintenanceStudentGbp)})` },
      { label: "Maintenance (dependants)", value: `${fmtGBP(fundsReq.maintenanceDependantsGbp)} (${fmtQuote(fundsReq.maintenanceDependantsGbp)})` },
      { label: "Buffer", value: `${fmtGBP(fundsReq.bufferGbp)} (${fmtQuote(fundsReq.bufferGbp)})` },
    ];
    const feesH = measureKeyValueBoxHeight(feesRows, halfWidth);
    const fundsReqH = measureKeyValueBoxHeight(fundsReqRows, halfWidth);

    const gapLabel = gapEligible >= 0 ? "Funds are sufficient" : "Additional funds required";
    const gapValue = gapEligible >= 0 ? fmtGBP(gapEligible) : fmtGBP(Math.abs(gapEligible));
    const gapValueQuote = quote === "GBP" ? "" : ` | ${gapEligible >= 0 ? fmtQuote(gapEligible) : fmtQuote(Math.abs(gapEligible))}`;
    const totalsRows = [
      { label: "Total required", value: `${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})` },
      { label: "Eligible funds", value: `${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})` },
      { label: gapLabel, value: `${gapValue}${gapValueQuote}`, valueColor: gapEligible >= 0 ? "#166534" : "#b91c1c" },
    ];
    const totalsH = measureKeyValueBoxHeight(totalsRows, halfWidth, 0.5);

    if (ensureSpace(Math.max(feesH, fundsReqH + totalsH + 8))) {
      drawKeyValueBox(margin, rowTop2, halfWidth, feesH, "Fees, IHS & Visa fee", feesRows, 0.5);
      drawKeyValueBox(margin + halfWidth + colGap, rowTop2, halfWidth, fundsReqH, "Funds required (28-day)", fundsReqRows, 0.5);
      drawKeyValueBox(margin + halfWidth + colGap, rowTop2 + fundsReqH + 8, halfWidth, totalsH, "Totals", totalsRows, 0.5);
      y = rowTop2 + Math.max(feesH, fundsReqH + totalsH + 8) + 8;
    }

    if (!fundsAvail.summary.skipped) {
      const issueCounts = new Map();
      fundsAvail.rows.forEach((r) => {
        if (!r.issues || !r.issues.length) return;
        if (r.eligible) return;
        r.issues.forEach((i) => issueCounts.set(i, (issueCounts.get(i) || 0) + 1));
      });
      const topIssues = [...issueCounts.entries()].sort((a, b) => b[1] - a[1]);
      if (topIssues.length) {
        const issueText = topIssues.map(([msg, count]) => `- ${msg} (${count})`).join("\n");
        const issueBoxH = Math.max(28, doc.heightOfString(issueText, { width: pageWidth - 20 }) + 18);
        if (ensureSpace(issueBoxH)) {
          doc.roundedRect(margin, y, pageWidth, issueBoxH, 8).fill("#fff7ed");
          doc.strokeColor("#fed7aa").lineWidth(1).roundedRect(margin, y, pageWidth, issueBoxH, 8).stroke();
          doc.fillColor("#9a3412").fontSize(9).text("Issue summary", margin + 10, y + 6);
          doc.fillColor("#7c2d12").fontSize(8).text(issueText, margin + 10, y + 18, { width: pageWidth - 20 });
          y += issueBoxH + 8;
        }
      }
    }

    if (pdfMode === "internal") {
      if (fundsAvail.summary.skipped) {
        const noteH = 22;
        if (ensureSpace(noteH + 6)) {
          doc.roundedRect(margin, y, pageWidth, noteH, 6).fill("#fff7ed");
          doc.strokeColor("#fed7aa").lineWidth(1).roundedRect(margin, y, pageWidth, noteH, 6).stroke();
          doc.fillColor("#9a3412").fontSize(8).text("Funds section skipped: student marked as not holding funds yet.", margin + 8, y + 6, { width: pageWidth - 16 });
          y += noteH + 6;
        }
      } else {
        if (ensureSpace(14)) {
          doc.fillColor("#0f172a").fontSize(10).text("Funds available (details)", margin, y);
          y += 12;
        }
        drawFundsTable(fundsAvail.rows);
        y += 6;
      }
    }

    const ruleText =
      `Bank statements: funds must be held for ${config.rules.funds_hold_days} consecutive days and end within ${config.rules.statement_age_days} days of the visa application date. ` +
      `FDs: maturity date required (28/31-day checks not applied). ` +
      `Education loans: disbursement letter should be within ${config.rules.loan_letter_max_age_days ?? 180} days of the application. ` +
      `Visa application date defaults to today if not provided.`;
    const rulesHeight = 22;
    if (y + rulesHeight > bottomLimit()) {
      y = bottomLimit() - rulesHeight;
    }
    doc.fillColor("#475569").fontSize(7).text(ruleText, margin, y, { width: pageWidth });
    y += 10;

    // PDF footer intentionally omitted per requirements.

    doc.end();
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

app.get("/healthz", (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

app.get("/health", (req, res) => res.json({ ok: true }));

scheduleSync("students", STUDENTS_SOURCE_URL, STUDENTS_XLSX_PATH, STUDENTS_SYNC_MS);
scheduleSync("counselors", COUNSELORS_SOURCE_URL, COUNSELORS_CSV_PATH, COUNSELORS_SYNC_MS);
scheduleSync("country_currency", COUNTRY_CURRENCY_SOURCE_URL, COUNTRY_CURRENCY_PATH, COUNTRY_CURRENCY_SYNC_MS);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
