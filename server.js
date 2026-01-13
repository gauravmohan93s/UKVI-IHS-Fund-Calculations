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
const STUDENTS_SYNC_MIN_AGE_MINUTES = Number(process.env.STUDENTS_SYNC_MIN_AGE_MINUTES || 0);
const COUNSELORS_SYNC_MIN_AGE_MINUTES = Number(process.env.COUNSELORS_SYNC_MIN_AGE_MINUTES || 0);
const COUNTRY_CURRENCY_SYNC_MIN_AGE_MINUTES = Number(process.env.COUNTRY_CURRENCY_SYNC_MIN_AGE_MINUTES || 0);
const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL || "";
const APP_VERSION = (() => {
  try {
    const pkgPath = path.join(__dirname, "package.json");
    const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf-8"));
    return String(pkg.version || "0.0.0");
  } catch (_) {
    return "0.0.0";
  }
})();

const localConfigCache = { mtimeMs: 0, data: null };
function readLocalConfig() {
  const stat = fs.statSync(LOCAL_CONFIG_PATH);
  if (stat.mtimeMs === localConfigCache.mtimeMs && localConfigCache.data) {
    return localConfigCache.data;
  }
  const parsed = JSON.parse(fs.readFileSync(LOCAL_CONFIG_PATH, "utf-8"));
  localConfigCache.mtimeMs = stat.mtimeMs;
  localConfigCache.data = parsed;
  return parsed;
}

function readCountryCurrency() {
  if (!fs.existsSync(COUNTRY_CURRENCY_PATH)) return {};
  return JSON.parse(fs.readFileSync(COUNTRY_CURRENCY_PATH, "utf-8"));
}

function normalizeStr(val) {
  return String(val || "").trim();
}
function normalizeLower(val) {
  return normalizeStr(val).toLowerCase();
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
    const ackNumber = normalizeStr(
      n.acknowledgementnumber || n.acknowledgmentnumber
    );
    const studentName = normalizeStr(n.studentname);
    const university = normalizeStr(n.university);
    const status = normalizeStr(n.status);
    const intakeYear = normalizeStr(
      n.intake ||
      n["intake"] ||
      n["intake inyear"] ||
      n["intake - inyear"] ||
      n["intake in year"] ||
      n["intake year"] ||
      n["intake - in year"] ||
      n.intakeinyear ||
      n.intakeyear
    );
    const tuitionFeeTotalGbp = normalizeStr(
      n["gross tuition fees"] ||
      n["gross tuition fee"] ||
      n["tuition fee"] ||
      n["tuition fee gbp"] ||
      n["tuition fee total"] ||
      n["tuitionfeetotal"] ||
      n["course fee"] ||
      n["course fee gbp"] ||
      n["coursefee"] ||
      n["total fee"] ||
      n["total fee gbp"]
    );
    const tuitionFeePaidGbp = normalizeStr(
      n["deposit paid"] ||
      n["tuition paid"] ||
      n["tuition paid gbp"] ||
      n["tuitionfeepaid"]
    );
    const scholarshipGbp = normalizeStr(
      n["scholarship"] ||
      n["scholarship gbp"] ||
      n["waiver"]
    );
    const assignee = normalizeStr(n.assignee);
    const assigneeEmail = normalizeStr(n.assigneeemail);
    const city = normalizeStr(n.city);
    const country = normalizeStr(n.country);
    return {
      ackNumber,
      ackLower: normalizeLower(ackNumber),
      studentId: normalizeStr(n["student id."] || n["student id"] || n.studentid || n["studentid"]),
      studentName,
      studentNameLower: normalizeLower(studentName),
      programName: normalizeStr(n.programname),
      university,
      status,
      intakeYear,
      tuitionFeeTotalGbp,
      tuitionFeePaidGbp,
      scholarshipGbp,
      applicationStageChangedOn: normalizeStr(n.applicationstagechangedon),
      assignee,
      assigneeEmail,
      assigneeLower: normalizeLower(assignee),
      assigneeEmailLower: normalizeLower(assigneeEmail),
      dob: normalizeStr(n.dob),
      gender: normalizeStr(n.gender),
      maritalStatus: normalizeStr(n.maritalstatus),
      city,
      country,
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
  const mapped = rows.map((r) => {
    const employeeId = normalizeStr(
      r["Employee ID"] || r.EmployeeID || r.EmployeeId
    );
    const name = normalizeStr(r.Name);
    const email = normalizeStr(r["Email ID (Official)"] || r.Email || r.EmailID);
    return {
      employeeId,
      employeeIdLower: normalizeLower(employeeId),
      name,
      nameLower: normalizeLower(name),
      email,
      emailLower: normalizeLower(email),
      region: normalizeStr(r.Region),
      subRegion: normalizeStr(r["Sub Region"] || r.SubRegion),
      designation: normalizeStr(r.Designation),
      roles: normalizeStr(r.Roles),
    };
  }).filter((r) => r.name || r.email || r.employeeId);

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
    const card = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        { type: "TextBlock", text: "KC Data Sync", weight: "Bolder", size: "Medium" },
        { type: "TextBlock", text: String(message || "-"), wrap: true }
      ]
    };
    await fetch(TEAMS_WEBHOOK_URL, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(card)
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

function scheduleSync(name, url, destPath, intervalMs, minAgeMinutes = 0) {
  if (!DATA_SYNC_ENABLED) return;
  if (!url) return;
  if (!Number.isFinite(intervalMs) || intervalMs <= 0) return;
  const minAgeMs = Number.isFinite(minAgeMinutes) && minAgeMinutes > 0 ? minAgeMinutes * 60 * 1000 : 0;
  const run = () => {
    if (shouldSync(name, destPath, minAgeMs)) {
      downloadToFile(name, url, destPath);
    }
  };
  console.log(`Sync enabled for ${name} every ${Math.round(intervalMs / 60000)}m`);
  run();
  const timer = setInterval(run, intervalMs);
  if (timer.unref) timer.unref();
}

function fileStatus(pathname) {
  if (!fs.existsSync(pathname)) return { exists: false, updatedAt: "" };
  const stat = fs.statSync(pathname);
  return { exists: true, updatedAt: new Date(stat.mtimeMs).toISOString() };
}

function getLastSuccess(name, pathname) {
  const meta = syncMeta.get(name);
  if (meta?.lastSuccessAt) return meta.lastSuccessAt;
  const status = fileStatus(pathname);
  return status.updatedAt || "";
}

function shouldSync(name, pathname, minAgeMs) {
  if (!Number.isFinite(minAgeMs) || minAgeMs <= 0) return true;
  const last = getLastSuccess(name, pathname);
  if (!last) return true;
  const lastMs = new Date(last).getTime();
  if (!Number.isFinite(lastMs)) return true;
  return (Date.now() - lastMs) >= minAgeMs;
}

// --- FX cache (daily) ---
const fxCache = new Map(); // key => data
const FX_TIMEOUT_MS = Number(process.env.FX_TIMEOUT_MS || 20000);
const FX_FALLBACK_URL = process.env.FX_FALLBACK_URL || "https://open.er-api.com/v6/latest";
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
    data.fxSource = "frankfurter.app";
    fxCache.set(key, data);
    return data;
  };
  const attemptFallback = async () => {
    const fbUrl = `${FX_FALLBACK_URL}/${encodeURIComponent(from)}`;
    const data = await fetchJson(fbUrl, FX_TIMEOUT_MS);
    data.fetchedAt = new Date().toISOString();
    lastFxFetchedAt = data.fetchedAt;
    data.fxSource = "open.er-api.com";
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

  const validateRate = (data) => {
    const rate = safeNum(data?.rates?.[toCsv]);
    return Number.isFinite(rate) && rate > 0;
  };

  try {
    const data = await attempt();
    if (validateRate(data)) return data;
    throw new Error("FX rate unavailable");
  } catch (e) {
    const override = tryManualOverride();
    if (override) return override;
    try {
      const fallback = await attemptFallback();
      if (validateRate(fallback)) return fallback;
    } catch (_) {
      // ignore fallback error
    }
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
    return null;
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

function addDays(dateStr, days){
  if (!dateStr) return null;
  const d = dateStr instanceof Date ? new Date(dateStr) : new Date(dateStr);
  if (Number.isNaN(d.getTime())) return null;
  const out = new Date(d.getTime());
  out.setDate(out.getDate() + Number(days || 0));
  return out;
}

function addWorkingDays(dateStr, days){
  if (!dateStr) return null;
  const d = dateStr instanceof Date ? new Date(dateStr) : new Date(dateStr);
  if (Number.isNaN(d.getTime())) return null;
  let remaining = Math.max(0, Math.floor(Number(days || 0)));
  const out = new Date(d.getTime());
  while (remaining > 0) {
    out.setDate(out.getDate() + 1);
    const day = out.getDay();
    if (day !== 0 && day !== 6) remaining -= 1;
  }
  return out;
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
    tuitionTotalGbp: round2(tuitionTotal),
    tuitionPaidGbp: round2(tuitionPaid),
    scholarshipGbp: round2(scholarship),
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

    const issues = [];
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
  const applyingFrom = payload.applyingFrom || "outside";
  const isPreSessional = Boolean(payload.isPreSessional);
  const applicationDate = payload.applicationDate || todayISOInIST();
  const applicationDateDefaulted = Boolean(payload.applicationDateDefaulted) || !payload.applicationDate;
  const visaServiceType = String(payload.visaServiceType || "standard").toLowerCase();
  const decisionDefaults = {
    standard: 20,
    priority: 5,
    "super-priority": 1,
    "super priority": 1,
    superpriority: 1,
    super_priority: 1
  };
  const decisionDefault = decisionDefaults[visaServiceType] ?? 20;
  const decisionDaysRaw = Number(payload.visaDecisionDays);
  const decisionDays = (Number.isFinite(decisionDaysRaw) && decisionDaysRaw > 0)
    ? Math.floor(decisionDaysRaw)
    : decisionDefault;
  const decisionDaysSafe = Math.max(1, decisionDays || decisionDefault || 1);
  const decisionDaysDefaulted = !(decisionDaysRaw > 0);

  const cs = courseStart ? new Date(courseStart) : null;
  const ce = courseEnd ? new Date(courseEnd) : null;
  const courseStartOk = cs && !Number.isNaN(cs.getTime());
  const courseEndOk = ce && !Number.isNaN(ce.getTime());
  const courseDays = (courseStartOk && courseEndOk)
    ? (Math.floor((ce - cs) / 86400000) + 1)
    : 0;
  const courseMonthsBase = datedifMonths(courseStart, courseEnd);
  const courseMonths = (courseStartOk && courseEndOk)
    ? Math.max(0, courseMonthsBase + (ce.getDate() >= cs.getDate() ? 1 : 0))
    : 0;

  let preWrapMonths = 0;
  let preWrapDays = 0;
  let postWrapMonths = 0;
  let postWrapDays = 0;
  let courseCategory = "Unknown";
  if (courseMonths >= 12) {
    preWrapMonths = 1;
    postWrapMonths = 4;
    courseCategory = "Course of 12 months or longer";
  } else if (courseMonths >= 6) {
    preWrapMonths = 1;
    postWrapMonths = 2;
    courseCategory = "Course of 6 months or longer but shorter than 12 months";
  } else if (isPreSessional) {
    preWrapMonths = 1;
    postWrapMonths = 1;
    courseCategory = "Pre-sessional course of less than 6 months";
  } else {
    preWrapDays = 7;
    postWrapDays = 7;
    courseCategory = "Course of less than 6 months (not pre-sessional)";
  }

  const wrapLabel = (months, days) => {
    if (months && months > 0) return `${months} month${months > 1 ? "s" : ""}`;
    return `${days} day${days > 1 ? "s" : ""}`;
  };
  const preWrapLabel = wrapLabel(preWrapMonths, preWrapDays);
  const postWrapLabel = wrapLabel(postWrapMonths, postWrapDays);

  const grantBase = new Date(applicationDate);
  const grantDateOk = !Number.isNaN(grantBase.getTime());
  const grantDate = grantDateOk ? addWorkingDays(grantBase, decisionDaysSafe) : null;
  const grantDateIso = grantDate ? formatDateISO(grantDate) : "";

  let intendedTravelDate = null;
  let intendedTravelDefaulted = false;
  let intendedTravelAdjusted = false;
  let intendedTravelAdjustedReason = "";
  if (payload.intendedTravelDate) {
    const candidate = new Date(payload.intendedTravelDate);
    if (!Number.isNaN(candidate.getTime())) {
      intendedTravelDate = candidate;
    } else {
      intendedTravelDefaulted = true;
    }
  }
  if (!intendedTravelDate && courseStartOk) {
    intendedTravelDefaulted = true;
    intendedTravelDate = preWrapDays > 0
      ? addDays(cs, -preWrapDays)
      : addMonthsExcel(courseStart, -1);
  }
  if (intendedTravelDate && grantDate && intendedTravelDate < grantDate) {
    intendedTravelDate = new Date(grantDate.getTime());
    intendedTravelAdjusted = true;
    intendedTravelAdjustedReason = "Adjusted to grant date";
  }
  const intendedTravelIso = intendedTravelDate ? formatDateISO(intendedTravelDate) : "";

  const oneMonthBeforeCourse = courseStartOk ? addMonthsExcel(courseStart, -1) : null;
  const sevenDaysBeforeTravel = intendedTravelDate ? addDays(intendedTravelDate, -7) : null;

  let visaStartDate = null;
  let visaStartRule = "ST 25.3(a)";
  if (grantDateOk && oneMonthBeforeCourse && grantDate <= oneMonthBeforeCourse) {
    visaStartDate = preWrapDays > 0 ? addDays(cs, -preWrapDays) : addMonthsExcel(courseStart, -preWrapMonths);
    visaStartRule = "ST 25.3(a): grant >= 1 month before course start";
  } else if (grantDateOk && sevenDaysBeforeTravel && grantDate <= sevenDaysBeforeTravel) {
    visaStartDate = sevenDaysBeforeTravel;
    visaStartRule = "ST 25.3(b): grant < 1 month before course start";
  } else if (grantDateOk) {
    visaStartDate = grantDate;
    visaStartRule = "ST 25.3(c): grant < 7 days before intended travel";
  } else if (courseStartOk) {
    visaStartDate = preWrapDays > 0 ? addDays(cs, -preWrapDays) : addMonthsExcel(courseStart, -preWrapMonths);
    visaStartRule = "ST 25.3(a): grant date unavailable";
  }

  const visaEndDate = courseEndOk
    ? (postWrapDays > 0 ? addDays(ce, postWrapDays) : addMonthsExcel(courseEnd, postWrapMonths))
    : null;
  const visaStartIso = formatDateISO(visaStartDate);
  const visaEndIso = formatDateISO(visaEndDate);

  let totalStayDays = 0;
  if (visaStartDate && visaEndDate) {
    totalStayDays = Math.max(0, Math.floor((visaEndDate - visaStartDate) / 86400000) + 1);
  }
  const ihsMonthsRaw = monthsForIHS(visaStartIso, visaEndIso);
  const totalStayMonths = Number.isFinite(ihsMonthsRaw) ? ihsMonthsRaw : 0;

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
    courseStart,
    courseEnd,
    visaStartDate: visaStartIso,
    visaEndDate: visaEndIso,
    visaStartRule,
    intendedTravelDate: intendedTravelIso,
    intendedTravelDefaulted,
    intendedTravelAdjusted,
    intendedTravelAdjustedReason,
    grantDate: grantDateIso,
    grantDateDefaulted: applicationDateDefaulted || !grantDateOk,
    grantDateEstimated: true,
    visaServiceType,
    decisionDays,
    decisionDaysDefaulted,
    courseCategory,
    courseMonths,
    courseDays,
    preWrapLabel,
    postWrapLabel,
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
    try {
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
      return res.json({ config: merged, source: "remote", config_url: CONFIG_URL });
    } catch (e) {
      return res.json({
        config: { ...local, country_currency: countryCurrency },
        source: "local-fallback",
        config_url: CONFIG_URL,
        error: String(e?.message || e)
      });
    }
  } catch (e) {
    res.status(500).json({ error: String(e) });
  }
});

app.get("/api/students", (req, res) => {
  const q = normalizeStr(req.query.q).toLowerCase();
  if (!q) return res.json({ items: [] });
  const rows = loadStudents();
  const items = rows.filter((r) =>
    String(r.ackLower || "").includes(q) ||
    String(r.studentNameLower || "").includes(q)
  ).slice(0, 10);
  res.json({ items });
});

app.get("/api/counselors", (req, res) => {
  const q = normalizeStr(req.query.q).toLowerCase();
  if (!q) return res.json({ items: [] });
  const rows = loadCounselors();
  const items = rows.filter((r) =>
    String(r.nameLower || "").includes(q) ||
    String(r.emailLower || "").includes(q) ||
    String(r.employeeIdLower || "").includes(q)
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
    const minAgeMinutes = Number(req.query.min_age_minutes || 0);
    const minAgeMs = Number.isFinite(minAgeMinutes) && minAgeMinutes > 0 ? minAgeMinutes * 60 * 1000 : 0;
    const tasks = [];
    const picked = [];
    const skipped = [];

    if (STUDENTS_SOURCE_URL && (wantAll || wants.has("students"))) {
      if (shouldSync("students", STUDENTS_XLSX_PATH, minAgeMs)) {
        tasks.push(downloadToFile("students", STUDENTS_SOURCE_URL, STUDENTS_XLSX_PATH));
        picked.push("students");
      } else {
        skipped.push("students");
      }
    }
    if (COUNSELORS_SOURCE_URL && (wantAll || wants.has("counselors"))) {
      if (shouldSync("counselors", COUNSELORS_CSV_PATH, minAgeMs)) {
        tasks.push(downloadToFile("counselors", COUNSELORS_SOURCE_URL, COUNSELORS_CSV_PATH));
        picked.push("counselors");
      } else {
        skipped.push("counselors");
      }
    }
    if (COUNTRY_CURRENCY_SOURCE_URL && (wantAll || wants.has("country_currency"))) {
      if (shouldSync("country_currency", COUNTRY_CURRENCY_PATH, minAgeMs)) {
        tasks.push(downloadToFile("country_currency", COUNTRY_CURRENCY_SOURCE_URL, COUNTRY_CURRENCY_PATH));
        picked.push("country_currency");
      } else {
        skipped.push("country_currency");
      }
    }

    if (!tasks.length) {
      if (skipped.length) {
        return res.json({ ok: true, targets: [], skipped, min_age_minutes: minAgeMinutes });
      }
      return res.status(400).json({ error: "No valid sync targets configured." });
    }

    await Promise.all(tasks);
    return res.json({ ok: true, targets: picked, skipped, min_age_minutes: minAgeMinutes });
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
    let quote = String(payload.quoteCurrency || "GBP").toUpperCase();
    let gbpToQuote = 1;
    let fxAvailable = true;
    if (quote !== "GBP") {
      try {
        const fx = await getFx("GBP", quote, payload.manualFx);
        const rate = Number(fx?.rates?.[quote]) || 0;
        if (!rate) throw new Error("FX unavailable");
        gbpToQuote = rate;
      } catch (_) {
        fxAvailable = false;
        quote = "GBP";
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

    const doc = new PDFDocument({ size: "A4", margin: 30 });
    doc.pipe(res);

    const b = config.branding || {};
    const c1 = b.primary_color || "#0b5cab";
    const c2 = b.secondary_color || "#00a3a3";
    const margin = 30;
    const pageWidth = doc.page.width - margin * 2;
    const headerHeight = 64;
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

    const drawValue = (text, x, yPos, width, baseSize = 8, minSize = 6) => {
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

      doc.fillColor("#ffffff").fontSize(14).text(b.product_name || "UK Visa Calculation Report", margin, 16, { width: pageWidth, align: "center" });
      doc.fontSize(8).text(b.company_name || "", margin, 34, { width: pageWidth, align: "center" });
      const fxStamp = lastFxFetchedAt ? `FX: ${formatDateTimeDisplay(lastFxFetchedAt)}` : "FX: -";
      const metaStamp = `v${APP_VERSION} | ${fxStamp}`;
      doc.fontSize(7).fillColor("#e2e8f0")
        .text(`Generated: ${formatDateTimeDisplay(new Date())}`, margin, headerHeight - 22, { width: pageWidth, align: "right" });
      doc.fontSize(7).fillColor("#e2e8f0")
        .text(metaStamp, margin, headerHeight - 10, { width: pageWidth, align: "right" });

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
      const boxH = 40;
      if (!ensureSpace(boxH + 6)) return;
      doc.roundedRect(margin, y, pageWidth, boxH, 10).fill(ok ? "#dcfce7" : "#fee2e2");
      doc.fillColor(ok ? "#065f46" : "#7f1d1d").fontSize(10)
        .text(ok ? "Result: ELIGIBLE (Funds are sufficient)" : "Result: NOT ELIGIBLE (Funds are short)", margin + 10, y + 6, { width: pageWidth - 20 });

      const colW = (pageWidth - 20) / 2;
      const leftX = margin + 10;
      const rightX = margin + 10 + colW;
      doc.fillColor("#0f172a").fontSize(8)
        .text(`Required: ${fmtGBP(fundsReq.fundsRequiredGbp)} (${fmtQuote(fundsReq.fundsRequiredGbp)})`, leftX, y + 24, { width: colW - 6 });
      doc.fillColor("#0f172a").fontSize(8)
        .text(`Eligible available: ${fmtGBP(fundsAvail.summary.totalEligibleGbp)} (${fmtQuote(fundsAvail.summary.totalEligibleGbp)})`, rightX, y + 24, { width: colW - 6 });
      y += boxH + 8;
      doc.fillColor("#000000");
    };

    const measureKeyValueBoxHeight = (rows, w, labelRatio = 0.58) => {
      const headerH = 14;
      const paddingX = 10;
      const labelW = Math.floor((w - paddingX * 2) * labelRatio);
      const valueW = (w - paddingX * 2) - labelW;
      let height = headerH + 6;
      rows.forEach((row) => {
        const label = String(row.label ?? "-");
        if (row.section) {
          doc.fontSize(7);
          const sh = doc.heightOfString(label, { width: w - paddingX * 2 });
          height += sh + 6;
          return;
        }
        const value = String(row.value ?? "-");
        doc.fontSize(6);
        const lh = doc.heightOfString(label, { width: labelW });
        if (row.valueParts && row.valueParts.singleLine) {
          doc.fontSize(8);
          const vh = doc.heightOfString(row.valueParts.primary || "-", { width: valueW, align: "right" });
          height += Math.max(lh, vh) + 4;
          return;
        }
        if (value.includes("\n")) {
          const [v1, v2] = value.split("\n");
          doc.fontSize(8);
          const h1 = doc.heightOfString(v1 || "-", { width: valueW, align: "right" });
          doc.fontSize(7);
          const h2 = doc.heightOfString(v2 || "-", { width: valueW, align: "right" });
          height += Math.max(lh, h1 + h2 + 2) + 4;
          return;
        }
        doc.fontSize(8);
        const vh = doc.heightOfString(value, { width: valueW, align: "right" });
        height += Math.max(lh, vh) + 4;
      });
      return height;
    };

    const drawKeyValueBox = (x, boxY, w, h, title, rows, labelRatio = 0.58) => {
      doc.roundedRect(x, boxY, w, h, 8).fill("#f8fafc");
      doc.strokeColor("#e2e8f0").lineWidth(1).roundedRect(x, boxY, w, h, 8).stroke();
      doc.fillColor("#0f172a").fontSize(9).text(title, x + 10, boxY + 4, { width: w - 20, align: "left" });

      const paddingX = 10;
      const headerH = 14;
      const labelW = Math.floor((w - paddingX * 2) * labelRatio);
      const valueW = (w - paddingX * 2) - labelW;
      let ry = boxY + headerH;
      rows.forEach((row) => {
        const label = String(row.label ?? "-");
        if (row.section) {
          doc.fontSize(7).fillColor("#334155").text(label, x + paddingX, ry, { width: w - paddingX * 2, align: "right" });
          const sh = doc.heightOfString(label, { width: w - paddingX * 2 });
          ry += sh + 4;
          doc.strokeColor("#e2e8f0").lineWidth(0.5).moveTo(x + paddingX, ry).lineTo(x + w - paddingX, ry).stroke();
          ry += 4;
          return;
        }
        const value = String(row.value ?? "-");
        doc.fontSize(6).fillColor("#64748b").text(label, x + paddingX, ry, { width: labelW });
        if (row.valueParts && row.valueParts.singleLine) {
          const primary = String(row.valueParts.primary || "-");
          const secondary = String(row.valueParts.secondary || "-");
          const sep = " | ";
          doc.fontSize(8);
          const w1 = doc.widthOfString(primary);
          doc.fontSize(7);
          const w2 = doc.widthOfString(secondary);
          const wSep = doc.widthOfString(sep);
          const totalW = w1 + wSep + w2;
          const startX = x + paddingX + labelW + Math.max(0, valueW - totalW);
          doc.fontSize(8).fillColor(row.valueColor || "#0f172a").text(primary, startX, ry, { lineBreak: false });
          doc.fontSize(7).fillColor("#64748b").text(sep + secondary, startX + w1, ry, { lineBreak: false });
          const lh = doc.heightOfString(label, { width: labelW });
          const vh = doc.heightOfString(primary, { width: valueW, align: "right" });
          ry += Math.max(lh, vh) + 4;
          return;
        }
        if (value.includes("\n")) {
          const [v1, v2] = value.split("\n");
          doc.fontSize(8).fillColor(row.valueColor || "#0f172a").text(v1 || "-", x + paddingX + labelW, ry, { width: valueW, align: "right" });
          const h1 = doc.heightOfString(v1 || "-", { width: valueW, align: "right" });
          doc.fontSize(7).fillColor("#64748b").text(v2 || "-", x + paddingX + labelW, ry + h1 + 2, { width: valueW, align: "right" });
          const h2 = doc.heightOfString(v2 || "-", { width: valueW, align: "right" });
          const lh = doc.heightOfString(label, { width: labelW });
          ry += Math.max(lh, h1 + h2 + 2) + 4;
          return;
        }
        doc.fontSize(8).fillColor(row.valueColor || "#0f172a").text(value, x + paddingX + labelW, ry, { width: valueW, align: "right" });
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
      doc.fillColor("#0f172a").fontSize(9).text(title, x + 10, boxY + 6);

      let ry = boxY + headerH;
      rows.forEach((r) => {
        const leftText = `${r.leftLabel}: ${r.leftValue || "-"}`;
        const rightText = `${r.rightLabel}: ${r.rightValue || "-"}`;
        doc.fontSize(7).fillColor("#0f172a").text(leftText, x + paddingX, ry, { width: colW });
        doc.text(rightText, x + paddingX + colW + colGap, ry, { width: colW });
        const lh = doc.heightOfString(leftText, { width: colW });
        const rh = doc.heightOfString(rightText, { width: colW });
        ry += Math.max(lh, rh) + 4;
      });
    };

    const drawFundsTable = (rows) => {
      const headerH = 12;
      const rowH = 8;
      const cols = [
        { title: "OK", width: 20, align: "center" },
        { title: "Type", width: 60, align: "center" },
        { title: "Owner", width: 60, align: "center" },
        { title: "Amount", width: 80, align: "center" },
        { title: "GBP", width: 80, align: "center" },
        { title: "Date", width: 90, align: "center" },
        { title: "Issues", width: pageWidth - (20 + 60 + 60 + 80 + 80 + 90) - 10, align: "left" },
      ];

      if (!ensureSpace(headerH + rowH + 4)) return;
      doc.roundedRect(margin, y, pageWidth, headerH, 6).fill("#f1f5f9");
      doc.fillColor("#334155").fontSize(6);
      let x = margin + 6;
      cols.forEach((c) => {
        doc.text(c.title, x, y + 4, { width: c.width, align: c.align || "left" });
        x += c.width;
      });
      y += headerH + 2;

      const fundTypeLabel = (t) => {
        const ft = String(t || "bank").toLowerCase();
        if (ft === "fd") return "FD";
        if (ft === "loan") return "Loan";
        return "Bank";
      };

      doc.fontSize(6);
      rows.forEach((r) => {
        const okTxt = r.eligible ? "OK" : "NO";
        const issuesText = (r.issues || []).join("; ") || "-";
        const dateText = r.dateValue || "-";
        const cells = [
          { text: okTxt, width: cols[0].width, align: cols[0].align, color: r.eligible ? "#065f46" : "#9a3412" },
          { text: fundTypeLabel(r.fundType), width: cols[1].width, align: cols[1].align, color: "#0f172a" },
          { text: r.accountType || "-", width: cols[2].width, align: cols[2].align, color: "#0f172a" },
          { text: fmtCurrency(r.currency || "GBP", r.amount || 0), width: cols[3].width, align: cols[3].align, color: "#0f172a" },
          { text: fmtGBP(r.amountGbp || 0), width: cols[4].width, align: cols[4].align, color: "#0f172a" },
          { text: dateText, width: cols[5].width, align: cols[5].align, color: "#0f172a" },
          { text: issuesText, width: cols[6].width, align: cols[6].align, color: "#475569" },
        ];

        const heights = cells.map((c) =>
          doc.heightOfString(String(c.text || "-"), { width: c.width - 4, align: c.align })
        );
        const rowHeight = Math.max(rowH, ...heights) + 2;
        if (!ensureSpace(rowHeight + 2)) return;

        let cx = margin + 6;
        cells.forEach((c) => {
          doc.fillColor(c.color || "#0f172a").text(String(c.text || "-"), cx, y, {
            width: c.width - 2,
            align: c.align
          });
          cx += c.width;
        });
        y += rowHeight;
      });
    };

    drawSummary();

    const studentCityCountry = [payload.studentCity, payload.studentCountry].filter(Boolean).join(", ");
    const courseDates = `${formatDateDisplay(payload.courseStart)} to ${formatDateDisplay(payload.courseEnd)}`;

    const colGap = 12;
    const halfWidth = (pageWidth - colGap) / 2;
    const rowTop = y;
    const locationLabel = normalizeRegion(payload.region || "outside_london") === "london"
      ? "London"
      : "Outside London";
    const universityValue = payload.universityName ? `${payload.universityName} (${locationLabel})` : "-";
    const programValue = payload.studentProgram
      ? `${payload.studentProgram}${payload.studentIntakeYear ? ` (${payload.studentIntakeYear})` : ""}`
      : "-";
    const studentRows = [
      { label: "Acknowledgement No", value: payload.studentAckNumber || "-" },
      { label: "Student name", value: payload.studentName || "-" },
      { label: "Student ID", value: payload.studentId || "-" },
      { label: "University", value: universityValue },
      { label: "Program", value: programValue },
      { label: "Course dates", value: courseDates },
      { label: "City / Country", value: studentCityCountry || "-" },
      { label: "Display currency", value: fxAvailable ? quote : "GBP (FX unavailable)" },
    ];
    const serviceLabel = ihs.visaServiceType ? String(ihs.visaServiceType).replace("-", " ") : "-";
    const decisionLabel = ihs.decisionDays ? `${ihs.decisionDays} working days` : "-";
    const serviceValue = serviceLabel !== "-"
      ? (decisionLabel !== "-" ? `${serviceLabel} (${decisionLabel})` : serviceLabel)
      : "-";
    const visaStartLabel = formatDateDisplay(ihs.visaStartDate);
    const visaEndLabel = formatDateDisplay(ihs.visaEndDate);
    const visaDateRange = (visaStartLabel !== "-" || visaEndLabel !== "-")
      ? `${visaStartLabel} - ${visaEndLabel}`
      : "-";
    const visaRows = [
      { label: "Visa application", value: formatDateDisplay(payload.applicationDate) },
      { label: "Visa service type", value: serviceValue },
      { label: "Intended travel", value: formatDateDisplay(payload.intendedTravelDate || ihs.intendedTravelDate) },
      { label: "Pre-sessional course", value: payload.isPreSessional ? "Yes" : "No" },
      { label: "Visa dates", value: visaDateRange },
      { label: "Total stay", value: `${ihs.totalStayMonths} months (${ihs.totalStayDays} days)` },
    ];
    const boxH1 = Math.max(
      measureKeyValueBoxHeight(studentRows, halfWidth),
      measureKeyValueBoxHeight(visaRows, halfWidth)
    );
    if (ensureSpace(boxH1)) {
      drawKeyValueBox(margin, rowTop, halfWidth, boxH1, "Student details", studentRows);
      drawKeyValueBox(margin + halfWidth + colGap, rowTop, halfWidth, boxH1, "Visa details", visaRows);
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
    const baseVisaFee = safeNum(config.fees?.visa_application_fee_gbp);
    const visaType = String(ihs.visaServiceType || "").toLowerCase();
    const visaFeeGbp = baseVisaFee + (visaType === "priority" ? 500 : (visaType === "super-priority" || visaType === "super priority" ? 1000 : 0));
    const dual = (n) => `${fmtGBP(n)}\n${fmtQuote(n)}`;
    const dualInline = (n) => ({ primary: fmtGBP(n), secondary: fmtQuote(n), singleLine: true });
    const feesRows = [
      { section: true, label: "Fees" },
      { label: "Tuition total", valueParts: dualInline(safeNum(payload.tuitionFeeTotalGbp)) },
      { label: "Tuition paid", valueParts: dualInline(safeNum(payload.tuitionFeePaidGbp)) },
      { label: "Scholarship", valueParts: dualInline(safeNum(payload.scholarshipGbp)) },
      { label: "Buffer", valueParts: dualInline(safeNum(payload.bufferGbp)) },
      { section: true, label: "Visa" },
      { label: "Visa application fee", valueParts: dualInline(visaFeeGbp) },
      { section: true, label: "IHS" },
      { label: "IHS calculation", value: ihsCalcText },
      { label: "IHS total", valueParts: dualInline(ihs.ihsTotalGbp) },
    ];
    const fundsReqRows = [
      { label: "Tuition due", value: dual(fundsReq.tuitionDueGbp) },
      { label: "Maintenance (student)", value: dual(fundsReq.maintenanceStudentGbp) },
      { label: "Maintenance (dependants)", value: dual(fundsReq.maintenanceDependantsGbp) },
      { label: "Buffer", value: dual(fundsReq.bufferGbp) },
    ];
    const feesH = measureKeyValueBoxHeight(feesRows, halfWidth);
    const fundsReqH = measureKeyValueBoxHeight(fundsReqRows, halfWidth);

    const gapLabel = gapEligible >= 0 ? "Funds are sufficient" : "Additional funds required";
    const gapValue = gapEligible >= 0 ? fmtGBP(gapEligible) : fmtGBP(Math.abs(gapEligible));
    const gapValueQuote = gapEligible >= 0 ? fmtQuote(gapEligible) : fmtQuote(Math.abs(gapEligible));
    const totalsRows = [
      { label: "Total required", value: dual(fundsReq.fundsRequiredGbp) },
      { label: "Eligible funds", value: dual(fundsAvail.summary.totalEligibleGbp) },
      { label: gapLabel, value: `${gapValue}\n${gapValueQuote}`, valueColor: gapEligible >= 0 ? "#166534" : "#b91c1c" },
    ];
    const totalsH = measureKeyValueBoxHeight(totalsRows, halfWidth, 0.5);

    if (ensureSpace(Math.max(feesH, fundsReqH + totalsH + 8))) {
      drawKeyValueBox(margin, rowTop2, halfWidth, feesH, "Fees, IHS & Visa", feesRows, 0.5);
      const rightTotalH = fundsReqH + totalsH + 8;
      const targetH = Math.max(feesH, rightTotalH);
      const extra = Math.max(0, targetH - rightTotalH);
      const totalsHAdj = totalsH + extra;
      drawKeyValueBox(margin + halfWidth + colGap, rowTop2, halfWidth, fundsReqH, "Funds required (28-day)", fundsReqRows, 0.5);
      drawKeyValueBox(margin + halfWidth + colGap, rowTop2 + fundsReqH + 8, halfWidth, totalsHAdj, "Totals", totalsRows, 0.5);
      y = rowTop2 + targetH + 8;
    }

    if (fundsAvail.rows.length) {
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
          doc.fillColor("#9a3412").fontSize(8).text("Issue summary", margin + 10, y + 6);
          doc.fillColor("#7c2d12").fontSize(7).text(issueText, margin + 10, y + 18, { width: pageWidth - 20 });
          y += issueBoxH + 8;
        }
      }
    }

    if (pdfMode === "internal") {
      if (fundsAvail.rows.length) {
        if (ensureSpace(14)) {
          doc.fillColor("#0f172a").fontSize(10).text("Funds available (details)", margin, y);
          y += 12;
        }
        drawFundsTable(fundsAvail.rows);
        y += 6;
      } else {
        const emptyRows = [{ label: "Status", value: "Funds details not provided." }];
        const emptyH = measureKeyValueBoxHeight(emptyRows, pageWidth);
        if (ensureSpace(emptyH)) {
          drawKeyValueBox(margin, y, pageWidth, emptyH, "Funds available (details)", emptyRows, 0.5);
          y += emptyH + 6;
        }
      }
    }

    const ruleText =
      `Bank statements: funds must be held for ${config.rules.funds_hold_days} consecutive days and end within ${config.rules.statement_age_days} days of the visa application date.\n` +
      `FDs: maturity date required (28/31-day checks not applied). Education loans: disbursement letter should be within ${config.rules.loan_letter_max_age_days ?? 180} days of the application.\n` +
      `Visa application date defaults to today if not provided.`;
    y += 8;
    const rulesHeight = doc.heightOfString(ruleText, { width: pageWidth });
    if (y + rulesHeight > bottomLimit()) {
      y = bottomLimit() - rulesHeight;
    }
    doc.fillColor("#475569").fontSize(6).text(ruleText, margin, y, { width: pageWidth });
    y += 10;

    // PDF footer intentionally omitted per requirements.

    doc.end();
  } catch (e) {
    res.status(400).json({ error: String(e) });
  }
});

app.get("/healthz", (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

app.get("/health", (req, res) => res.json({ ok: true }));

scheduleSync("students", STUDENTS_SOURCE_URL, STUDENTS_XLSX_PATH, STUDENTS_SYNC_MS, STUDENTS_SYNC_MIN_AGE_MINUTES);
scheduleSync("counselors", COUNSELORS_SOURCE_URL, COUNSELORS_CSV_PATH, COUNSELORS_SYNC_MS, COUNSELORS_SYNC_MIN_AGE_MINUTES);
scheduleSync("country_currency", COUNTRY_CURRENCY_SOURCE_URL, COUNTRY_CURRENCY_PATH, COUNTRY_CURRENCY_SYNC_MS, COUNTRY_CURRENCY_SYNC_MIN_AGE_MINUTES);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
