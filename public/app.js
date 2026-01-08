const $ = (id) => document.getElementById(id);

const els = {
  fxPanel: $("fxPanel"),
  fxFallbackNote: $("fxFallbackNote"),
  manualInrPerGbp: $("manualInrPerGbp"),
  useManualFx: $("useManualFx"),
  universityInput: $("universityInput"),
  uniSuggest: $("uniSuggest"),
  formErrors: $("formErrors"),
  accessModal: $("accessModal"),
  accessCodeInput: $("accessCodeInput"),
  accessCodeSave: $("accessCodeSave"),
  accessCodeError: $("accessCodeError"),

  universitySelect: $("universitySelect"),
  courseStart: $("courseStart"),
  courseEnd: $("courseEnd"),
  applicationDate: $("applicationDate"),
  regionDisplay: $("regionDisplay"),
  quote: $("quote"),
  studentAck: $("studentAck"),
  studentSuggest: $("studentSuggest"),
  studentName: $("studentName"),
  studentProgram: $("studentProgram"),
  studentStatus: $("studentStatus"),
  studentIntakeYear: $("studentIntakeYear"),
  studentCityCountry: $("studentCityCountry"),
  counselorInput: $("counselorInput"),
  counselorSuggest: $("counselorSuggest"),
  counselorEmail: $("counselorEmail"),
  counselorRegion: $("counselorRegion"),
  counselorDesignation: $("counselorDesignation"),
  tuitionTotal: $("tuitionTotal"),
  tuitionPaid: $("tuitionPaid"),
  scholarship: $("scholarship"),
  dependants: $("dependants"),
  buffer: $("buffer"),
  rateHint: $("rateHint"),
  ratePanel: $("ratePanel"),

  addRow: $("addRow"),
  clearRows: $("clearRows"),
  fundsTbody: document.querySelector("#fundsTable tbody"),
  fundsSkip: $("fundsSkip"),
  fundsSkipNote: $("fundsSkipNote"),

  btnCalc: $("btnCalc"),
  btnPdf: $("btnPdf"),
  btnReset: $("btnReset"),

  ihsGbp: $("ihsGbp"),
  ihsCalc: $("ihsCalc"),
  ihsFx: $("ihsFx"),
  fundsReqGbp: $("fundsReqGbp"),
  fundsReqFx: $("fundsReqFx"),
  fundsAvailGbp: $("fundsAvailGbp"),
  fundsAvailFx: $("fundsAvailFx"),
  gapGbp: $("gapGbp"),
  gapFx: $("gapFx"),
  gapLabel: $("gapLabel"),

  bTuition: $("bTuition"),
  bStudent: $("bStudent"),
  bDeps: $("bDeps"),
  bBuffer: $("bBuffer"),
  bTotal: $("bTotal"),
  fundsAvailBreakBody: document.querySelector("#fundsAvailBreak tbody"),
  validationNote: $("validationNote"),
  issueSummary: $("issueSummary"),
  rulesNote: $("rulesNote"),
  eligibilityStatus: $("eligibilityStatus"),
  sourcesNote: $("sourcesNote"),
};

let CONFIG = null;
let SELECTED_REGION = "outside_london";
let CURRENT_STUDENT = null;
let CURRENT_COUNSELOR = null;
let QUOTE_MANUAL = false;
const currencyLocale = (cur) => {
  const c = String(cur || "").toUpperCase();
  if (c === "INR") return "en-IN";
  if (c === "GBP") return "en-GB";
  if (c === "EUR") return "de-DE";
  return "en-US";
};
const fmtNumber = (n) => (Number.isFinite(n) ? n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "-");
const fmtMoney = (cur, n) => new Intl.NumberFormat(currencyLocale(cur), {
  style: "currency",
  currency: String(cur || "GBP").toUpperCase(),
  currencySign: "accounting",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2
}).format(Number(n || 0));
const fmtGBP = (n) => fmtMoney("GBP", n);

function setActiveTab(id){
  document.querySelectorAll(".tab").forEach(b => b.classList.toggle("active", b.dataset.tab === id));
  document.querySelectorAll(".tabpane").forEach(p => p.classList.toggle("active", p.id === id));
  window.scrollTo({ top: 0, behavior: "smooth" });
}

document.querySelectorAll(".tab").forEach(b => b.addEventListener("click", () => setActiveTab(b.dataset.tab)));
document.querySelectorAll("[data-next]").forEach(b => b.addEventListener("click", () => setActiveTab(b.dataset.next)));
document.querySelectorAll("[data-prev]").forEach(b => b.addEventListener("click", () => setActiveTab(b.dataset.prev)));


function applyBranding(){
  const b = CONFIG?.branding || {};
  const logoImg = document.getElementById("brandLogoImg");
  const logoText = document.getElementById("brandLogo");

  if (b.company_name) document.getElementById("brandName").textContent = b.company_name;
  if (b.product_name) document.getElementById("brandProduct").textContent = b.product_name;

  if (logoImg && b.logo_web) {
    logoImg.src = b.logo_web;
    logoImg.style.display = "block";
    if (logoText) logoText.style.display = "none";
  } else if (logoText) {
    logoText.textContent = b.logo_text || "KC";
    logoText.style.display = "flex";
    if (logoImg) logoImg.style.display = "none";
  }

  if (b.footer_note) document.getElementById("brandFooter").innerHTML = `<p>${escapeHtml(b.footer_note)}</p>`;

  const root = document.documentElement;
  if (b.primary_color) root.style.setProperty("--brand", b.primary_color);
  if (b.secondary_color) root.style.setProperty("--brand2", b.secondary_color);
  if (b.accent_color) root.style.setProperty("--brand3", b.accent_color);
}



function normalizeRegionLabel(r){
  return (r === "london") ? "London" : "Outside London";
}

function formatDateDisplay(value){
  if (!value) return "-";
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return "-";
  return d.toLocaleDateString("en-IN", { day: "2-digit", month: "short", year: "numeric" });
}

function applyCountryCurrency(country){
  if (!country || QUOTE_MANUAL) return;
  const map = CONFIG?.country_currency || {};
  const cur = map[country];
  if (!cur) return;
  if (els.quote) els.quote.value = cur;
}

function setUniversityByName(name){
  const list = CONFIG?.universities || [];
  const match = list.find(u => (u.name || "").toLowerCase() === (name || "").toLowerCase());
  if (!match) return false;
  els.universitySelect.value = match.name;
  SELECTED_REGION = (match.region === "london") ? "london" : "outside_london";
  els.regionDisplay.value = normalizeRegionLabel(SELECTED_REGION);
  if (els.universityInput) els.universityInput.value = match.name;
  return true;
}

function updateUniSuggest(){
  if (!els.universityInput || !els.uniSuggest) return;
  const q = (els.universityInput.value || "").toLowerCase().trim();
  const list = (CONFIG?.universities || []);
  if (!q){
    els.uniSuggest.style.display = "none";
    els.uniSuggest.innerHTML = "";
    return;
  }
  const matches = list
    .filter(u => (u.name || "").toLowerCase().includes(q))
    .slice(0, 5);

  if (!matches.length){
    els.uniSuggest.style.display = "none";
    els.uniSuggest.innerHTML = "";
    return;
  }

  els.uniSuggest.innerHTML = matches.map(u => {
    const loc = normalizeRegionLabel(u.region);
    return `<div class="item" data-name="${escapeHtml(u.name)}">
      ${escapeHtml(u.name)} <span class="meta">${loc}</span>
    </div>`;
  }).join("");

  els.uniSuggest.style.display = "block";
}

function initUniversitySearch(){
  if (!els.universityInput) return;
  els.universityInput.addEventListener("input", updateUniSuggest);
  els.universityInput.addEventListener("blur", () => {
    setTimeout(()=>{ if (els.uniSuggest) els.uniSuggest.style.display="none"; }, 150);
  });
  if (els.uniSuggest){
    els.uniSuggest.addEventListener("mousedown", (e) => {
      const item = e.target.closest(".item");
      if (!item) return;
      const name = item.getAttribute("data-name") || "";
      setUniversityByName(name);
      els.uniSuggest.style.display = "none";
      showErrors([]);
    });
  }
}

let studentSuggestMap = new Map();
const studentCache = new Map();
let studentSearchTimer = null;
async function searchStudents(query){
  const q = String(query || "").trim();
  if (q.length < 2) return [];
  if (studentCache.has(q)) return studentCache.get(q);
  const out = await apiFetch(`/api/students?q=${encodeURIComponent(q)}`).then(r=>r.json());
  const items = Array.isArray(out.items) ? out.items : [];
  studentCache.set(q, items);
  return items;
}

function getCachedStudentMatches(query){
  const q = String(query || "").trim().toLowerCase();
  if (!q || q.length < 2) return [];
  let best = null;
  for (const [key, items] of studentCache.entries()){
    if (!key) continue;
    const lk = key.toLowerCase();
    if (q.startsWith(lk) && (!best || lk.length > best.key.length)){
      best = { key: lk, items };
    }
  }
  if (!best) return [];
  return best.items.filter((r) =>
    String(r.ackNumber || "").toLowerCase().includes(q) ||
    String(r.studentName || "").toLowerCase().includes(q)
  );
}

function renderStudentSuggest(items){
  if (!els.studentSuggest) return;
  studentSuggestMap = new Map();
  if (!items.length){
    els.studentSuggest.style.display = "none";
    els.studentSuggest.innerHTML = "";
    return;
  }
  items.forEach(i => studentSuggestMap.set(i.ackNumber, i));
  els.studentSuggest.innerHTML = items.map(i => {
    const uni = i.university ? ` - ${escapeHtml(i.university)}` : "";
    const sub = i.studentName ? ` | ${escapeHtml(i.studentName)}` : "";
    return `<div class="item" data-ack="${escapeHtml(i.ackNumber)}">${escapeHtml(i.ackNumber)}${sub}${uni}</div>`;
  }).join("");
  els.studentSuggest.style.display = "block";
}

async function applyStudent(item){
  CURRENT_STUDENT = item || null;
  if (!item) return;
  if (els.studentAck) els.studentAck.value = item.ackNumber || "";
  if (els.studentName) els.studentName.value = item.studentName || "";
  if (els.studentProgram) els.studentProgram.value = item.programName || "";
  if (els.studentStatus) els.studentStatus.value = item.status || "";
  if (els.studentIntakeYear) els.studentIntakeYear.value = item.intakeYear || "";
  if (els.studentCityCountry) {
    const cc = [item.city, item.country].filter(Boolean).join(", ");
    els.studentCityCountry.value = cc;
  }
  if (item.country) applyCountryCurrency(item.country);

  if (item.university){
    const ok = setUniversityByName(item.university);
    if (!ok && els.universityInput) els.universityInput.value = item.university;
  }

  if (els.counselorInput) els.counselorInput.value = item.assignee || "";
  if (els.counselorEmail) els.counselorEmail.value = item.assigneeEmail || "";
  if (els.studentIntakeYear && item.intakeYear) els.studentIntakeYear.value = item.intakeYear;

  if (item.assigneeEmail || item.assignee){
    const q = item.assigneeEmail || item.assignee;
    const res = await searchCounselors(q);
    if (res.length) applyCounselor(res[0]);
  }
}

function initStudentSearch(){
  if (!els.studentAck) return;
  els.studentAck.addEventListener("input", () => {
    CURRENT_STUDENT = null;
    if (els.studentName) els.studentName.value = "";
    if (els.studentProgram) els.studentProgram.value = "";
    if (els.studentStatus) els.studentStatus.value = "";
    if (els.studentIntakeYear) els.studentIntakeYear.value = "";
    if (els.studentCityCountry) els.studentCityCountry.value = "";
    const cached = getCachedStudentMatches(els.studentAck.value);
    if (cached.length) renderStudentSuggest(cached);
    if (studentSearchTimer) clearTimeout(studentSearchTimer);
    studentSearchTimer = setTimeout(async () => {
      const items = await searchStudents(els.studentAck.value);
      renderStudentSuggest(items);
    }, 120);
  });
  els.studentAck.addEventListener("focus", async () => {
    const items = await searchStudents(els.studentAck.value);
    renderStudentSuggest(items);
  });
  document.addEventListener("click", (e) => {
    const t = e.target;
    if (!(t instanceof HTMLElement)) return;
    const item = t.closest("#studentSuggest .item");
    if (item && item.dataset.ack){
      const picked = studentSuggestMap.get(item.dataset.ack);
      applyStudent(picked);
      els.studentSuggest.style.display = "none";
      return;
    }
    if (!t.closest("#studentSuggest") && !t.closest("#studentAck")){
      if (els.studentSuggest) els.studentSuggest.style.display = "none";
    }
  });
}

let counselorSuggestMap = new Map();
async function searchCounselors(query){
  const q = String(query || "").trim();
  if (q.length < 2) return [];
  const out = await apiFetch(`/api/counselors?q=${encodeURIComponent(q)}`).then(r=>r.json());
  return Array.isArray(out.items) ? out.items : [];
}

function renderCounselorSuggest(items){
  if (!els.counselorSuggest) return;
  counselorSuggestMap = new Map();
  if (!items.length){
    els.counselorSuggest.style.display = "none";
    els.counselorSuggest.innerHTML = "";
    return;
  }
  items.forEach(i => counselorSuggestMap.set(i.email || i.name, i));
  els.counselorSuggest.innerHTML = items.map(i => {
    const sub = i.email ? ` | ${escapeHtml(i.email)}` : "";
    return `<div class="item" data-key="${escapeHtml(i.email || i.name)}">${escapeHtml(i.name || i.email)}${sub}</div>`;
  }).join("");
  els.counselorSuggest.style.display = "block";
}

function applyCounselor(item){
  CURRENT_COUNSELOR = item || null;
  if (!item) return;
  if (els.counselorInput) els.counselorInput.value = item.name || "";
  if (els.counselorEmail) els.counselorEmail.value = item.email || "";
  if (els.counselorRegion) {
    const reg = [item.region, item.subRegion].filter(Boolean).join(" / ");
    els.counselorRegion.value = reg;
  }
  if (els.counselorDesignation) {
    const des = [item.designation, item.roles].filter(Boolean).join(" / ");
    els.counselorDesignation.value = des;
  }
}

function initCounselorSearch(){
  if (!els.counselorInput) return;
  els.counselorInput.addEventListener("input", async () => {
    CURRENT_COUNSELOR = null;
    if (els.counselorEmail) els.counselorEmail.value = "";
    if (els.counselorRegion) els.counselorRegion.value = "";
    if (els.counselorDesignation) els.counselorDesignation.value = "";
    const items = await searchCounselors(els.counselorInput.value);
    renderCounselorSuggest(items);
  });
  els.counselorInput.addEventListener("focus", async () => {
    const items = await searchCounselors(els.counselorInput.value);
    renderCounselorSuggest(items);
  });
  document.addEventListener("click", (e) => {
    const t = e.target;
    if (!(t instanceof HTMLElement)) return;
    const item = t.closest("#counselorSuggest .item");
    if (item && item.dataset.key){
      const picked = counselorSuggestMap.get(item.dataset.key);
      applyCounselor(picked);
      els.counselorSuggest.style.display = "none";
      return;
    }
    if (!t.closest("#counselorSuggest") && !t.closest("#counselorInput")){
      if (els.counselorSuggest) els.counselorSuggest.style.display = "none";
    }
  });
}

async function loadConfig(){
  const out = await apiFetch("/api/config").then(r=>r.json());
  CONFIG = out.config;
  applyBranding();
  initUniversitySearch();
  // universities dropdown
  for (const u of (CONFIG.universities || [])){
    const opt = document.createElement("option");
    opt.value = u.name;
    const loc = u.location || (u.region === "london" ? "London" : "Outside London");
    opt.textContent = loc ? `${u.name} - ${loc}` : u.name;
    opt.dataset.region = u.region || "outside_london";
    els.universitySelect.appendChild(opt);
  }

  // rate hint
  const s = CONFIG.routes.student;
  const cs = CONFIG.routes.child_16_17_independent;
  if (els.ratePanel) {
    els.ratePanel.innerHTML = `
      <div class="rate-card">
        <div class="rate-title">Maintenance rates (per month)</div>
        <div class="rate-list">
          <div><span class="rate-strong">Student (London):</span> ${fmtMoney("GBP", s.maintenance_monthly_gbp.london)}</div>
          <div><span class="rate-strong">Student (Outside):</span> ${fmtMoney("GBP", s.maintenance_monthly_gbp.outside_london)}</div>
          <div><span class="rate-strong">Dependants (London):</span> ${fmtMoney("GBP", s.dependant_monthly_gbp.london)}</div>
          <div><span class="rate-strong">Dependants (Outside):</span> ${fmtMoney("GBP", s.dependant_monthly_gbp.outside_london)}</div>
          <div><span class="rate-strong">Child 16–17 (London):</span> ${fmtMoney("GBP", cs.maintenance_monthly_gbp.london)}</div>
          <div><span class="rate-strong">Child 16–17 (Outside):</span> ${fmtMoney("GBP", cs.maintenance_monthly_gbp.outside_london)}</div>
          <div><span class="rate-strong">Cap:</span> ${s.max_months} months</div>
        </div>
      </div>
      <div class="rate-card">
        <div class="rate-title">IHS rate</div>
        <div class="rate-list">
          <div><span class="rate-strong">Yearly:</span> ${fmtMoney("GBP", CONFIG.ihs.student_yearly_gbp)}</div>
          <div><span class="rate-strong">Half‑year:</span> ${fmtMoney("GBP", CONFIG.ihs.half_year_gbp)}</div>
        </div>
      </div>
    `;
  }

  els.sourcesNote.textContent = `Rates sourced from GOV.UK guidance (configured server-side). FX via frankfurter.app.`;
}


els.universitySelect.addEventListener("change", () => {
  const opt = els.universitySelect.selectedOptions[0];
  const region = opt?.dataset?.region || "outside_london";
  SELECTED_REGION = (region === "london") ? "london" : "outside_london";
  els.regionDisplay.value = (SELECTED_REGION === "london") ? "London" : "Outside London";
});

if (els.quote) {
  els.quote.addEventListener("change", () => { QUOTE_MANUAL = true; });
}

function getFundsRows(){
  const rows = [];
  els.fundsTbody.querySelectorAll("tr").forEach(tr=>{
    const fundType = tr.querySelector("[data-role='fundType']")?.value || "bank";
    const accountType = tr.querySelector("[data-role='accountType']")?.value || "Student";
    const source = tr.querySelector("[data-role='source']")?.value || "";
    const currency = tr.querySelector("[data-role='currency']")?.value || "GBP";
    const amount = Number(tr.querySelector("[data-role='amount']")?.value || 0);
    const statementStart = tr.querySelector("[data-field='statementStart']")?.value || "";
    const statementEnd = tr.querySelector("[data-field='statementEnd']")?.value || "";
    const fdMaturity = tr.querySelector("[data-field='fdMaturity']")?.value || "";
    const loanDisbursement = tr.querySelector("[data-field='loanDisbursement']")?.value || "";
    rows.push({
      fundType,
      accountType,
      source,
      currency,
      amount,
      statementStart,
      statementEnd,
      fdMaturity,
      loanDisbursement
    });
  });
  return rows;
}

function setFundFieldState(input, enabled){
  if (!input) return;
  input.disabled = !enabled;
  input.classList.toggle("muted-field", !enabled);
  if (!enabled) input.value = "";
}

function updateFundRow(tr){
  const fundType = tr.querySelector("[data-role='fundType']")?.value || "bank";
  const isBank = fundType === "bank";
  const isFd = fundType === "fd";
  const isLoan = fundType === "loan";
  setFundFieldState(tr.querySelector("[data-field='statementStart']"), isBank);
  setFundFieldState(tr.querySelector("[data-field='statementEnd']"), isBank);
  setFundFieldState(tr.querySelector("[data-field='fdMaturity']"), isFd);
  setFundFieldState(tr.querySelector("[data-field='loanDisbursement']"), isLoan);
}

function addFundsRow(pref={}){
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td>
      <select data-role="fundType">
        <option value="bank">Bank statement</option>
        <option value="fd">Fixed deposit (FD)</option>
        <option value="loan">Education loan</option>
      </select>
    </td>
    <td>
      <select data-role="accountType">
        <option value="Student">Student</option>
        <option value="Parent">Parent</option>
        <option value="Sponsor">Sponsor</option>
      </select>
    </td>
    <td><input data-role="source" type="text" placeholder="Bank / Sponsor / Account notes" value="${pref.source || ""}"></td>
    <td>
      <select data-role="currency">
        <option value="GBP">GBP - British Pound (GBP)</option>
        <option value="INR">INR - Indian Rupee (INR)</option>
        <option value="USD">USD - US Dollar (USD)</option>
        <option value="EUR">EUR - Euro (EUR)</option>
        <option value="BDT">BDT - Bangladeshi Taka (BDT)</option>
        <option value="PKR">PKR - Pakistani Rupee (PKR)</option>
        <option value="LKR">LKR - Sri Lankan Rupee (LKR)</option>
        <option value="NGN">NGN - Nigerian Naira (NGN)</option>
        <option value="KES">KES - Kenyan Shilling (KES)</option>
        <option value="GHS">GHS - Ghanaian Cedi (GHS)</option>
        <option value="UGX">UGX - Ugandan Shilling (UGX)</option>
        <option value="ZAR">ZAR - South African Rand (ZAR)</option>
        <option value="AED">AED - UAE Dirham (AED)</option>
        <option value="SAR">SAR - Saudi Riyal (SAR)</option>
      </select>
    </td>
    <td><input data-role="amount" type="number" min="0" step="0.01" placeholder="Closing balance" value="${pref.amount || ""}"></td>
    <td><input data-field="statementStart" type="date" value="${pref.statementStart || ""}"></td>
    <td><input data-field="statementEnd" type="date" value="${pref.statementEnd || ""}"></td>
    <td><input data-field="fdMaturity" type="date" value="${pref.fdMaturity || ""}"></td>
    <td><input data-field="loanDisbursement" type="date" value="${pref.loanDisbursement || ""}"></td>
    <td><button class="icon-btn" title="Remove">X</button></td>
  `;
  const fundTypeSel = tr.querySelector("[data-role='fundType']");
  const acctSel = tr.querySelector("[data-role='accountType']");
  const curSel = tr.querySelector("[data-role='currency']");
  fundTypeSel.value = pref.fundType || "bank";
  acctSel.value = pref.accountType || "Student";
  curSel.value = pref.currency || "GBP";

  fundTypeSel.addEventListener("change", () => updateFundRow(tr));
  tr.querySelector(".icon-btn").addEventListener("click", ()=> tr.remove());
  updateFundRow(tr);
  els.fundsTbody.appendChild(tr);
}

function toggleFundsSkip(){
  const skip = Boolean(els.fundsSkip && els.fundsSkip.checked);
  if (els.fundsSkipNote) els.fundsSkipNote.style.display = skip ? "block" : "none";
  if (els.addRow) els.addRow.disabled = skip;
  if (els.clearRows) els.clearRows.disabled = skip;
  els.fundsTbody.querySelectorAll("input,select,button").forEach(el => {
    el.disabled = skip;
  });
  if (!skip) {
    els.fundsTbody.querySelectorAll("tr").forEach(updateFundRow);
  }
}

els.addRow.addEventListener("click", () => addFundsRow({ fundType:"bank", accountType:"Student", currency:"INR" }));
els.clearRows.addEventListener("click", () => { els.fundsTbody.innerHTML = ""; });
if (els.fundsSkip) els.fundsSkip.addEventListener("change", toggleFundsSkip);

async function fx(from, to){
  if (from === to) return 1;
  if (!fx._cache) fx._cache = new Map();
  const key = `${from}|${to}`;
  const cached = fx._cache.get(key);
  const now = Date.now();
  if (cached && (now - cached.ts) < 10 * 60 * 1000) return cached.rate;
  const manualEnabled = Boolean(els.useManualFx && els.useManualFx.checked);
  const manualRate = Number(els.manualInrPerGbp && els.manualInrPerGbp.value || 0);
  const manualQuery = (manualEnabled && manualRate > 0)
    ? `&manual_enabled=true&manual_inr_per_gbp=${encodeURIComponent(manualRate)}`
    : "";
  const out = await apiFetch(`/api/fx?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}${manualQuery}`).then(r=>r.json());
  const rate = Number(out.rate);
  if (!Number.isFinite(rate) || rate <= 0) {
    if (els.fxFallbackNote) els.fxFallbackNote.style.display = "block";
    return 1;
  }
  if (els.fxFallbackNote) els.fxFallbackNote.style.display = "none";
  fx._cache.set(key, { rate, ts: now });
  return rate;
}

function payload(){
  return {
    routeKey: "student",
    universityName: els.universitySelect.value || "",
    region: SELECTED_REGION,
    courseStart: els.courseStart.value,
    courseEnd: els.courseEnd.value,
    applicationDate: els.applicationDate.value || null,

    studentAckNumber: (els.studentAck && els.studentAck.value) || (CURRENT_STUDENT && CURRENT_STUDENT.ackNumber) || "",
    studentName: (els.studentName && els.studentName.value) || (CURRENT_STUDENT && CURRENT_STUDENT.studentName) || "",
    studentProgram: (els.studentProgram && els.studentProgram.value) || (CURRENT_STUDENT && CURRENT_STUDENT.programName) || "",
    studentStatus: (els.studentStatus && els.studentStatus.value) || (CURRENT_STUDENT && CURRENT_STUDENT.status) || "",
    studentIntakeYear: (els.studentIntakeYear && els.studentIntakeYear.value) || (CURRENT_STUDENT && CURRENT_STUDENT.intakeYear) || "",
    studentCity: (CURRENT_STUDENT && CURRENT_STUDENT.city) || "",
    studentCountry: (CURRENT_STUDENT && CURRENT_STUDENT.country) || "",
    counselorName: (els.counselorInput && els.counselorInput.value) || (CURRENT_COUNSELOR && CURRENT_COUNSELOR.name) || "",
    counselorEmail: (els.counselorEmail && els.counselorEmail.value) || (CURRENT_COUNSELOR && CURRENT_COUNSELOR.email) || "",
    counselorRegion: (CURRENT_COUNSELOR && CURRENT_COUNSELOR.region) || "",
    counselorSubRegion: (CURRENT_COUNSELOR && CURRENT_COUNSELOR.subRegion) || "",
    counselorDesignation: (CURRENT_COUNSELOR && CURRENT_COUNSELOR.designation) || "",
    counselorRoles: (CURRENT_COUNSELOR && CURRENT_COUNSELOR.roles) || "",
    counselorEmployeeId: (CURRENT_COUNSELOR && CURRENT_COUNSELOR.employeeId) || "",

    quoteCurrency: els.quote.value,

    tuitionFeeTotalGbp: Number(els.tuitionTotal.value || 0),
    tuitionFeePaidGbp: Number(els.tuitionPaid.value || 0),
    scholarshipGbp: Number(els.scholarship.value || 0),
    dependantsCount: Number(els.dependants.value || 0),
    bufferGbp: Number(els.buffer.value || 0),

    fundsSkip: Boolean(els.fundsSkip && els.fundsSkip.checked),
    fundsRows: (els.fundsSkip && els.fundsSkip.checked) ? [] : getFundsRows(),

    manualFx: {
      enabled: Boolean(els.useManualFx && els.useManualFx.checked),
      inrPerGbp: Number(els.manualInrPerGbp && els.manualInrPerGbp.value || 0)
    }
  };
}


function showErrors(list){
  if (!els.formErrors) return;
  if (!list || !list.length) { els.formErrors.style.display = "none"; els.formErrors.innerHTML=""; return; }
  els.formErrors.style.display = "block";
  els.formErrors.innerHTML = "<strong>Please fix:</strong><ul>" + list.map(x=>`<li>${escapeHtml(x)}</li>`).join("") + "</ul>";
}

function markField(el, isBad){
  if (!el) return;
  if (isBad) el.classList.add("field-error");
  else el.classList.remove("field-error");
}

function validateTab(tabId){
  const errs = [];
  // clear previous highlights
  ["universityInput","courseStart","courseEnd","applicationDate"].forEach(id => markField($(id), false));
  if (tabId === "t1"){
    const uniOk = Boolean(els.universitySelect.value);
    markField(els.universityInput, !uniOk);
    if (!uniOk) errs.push("Select a university.");
    

    const sOk = Boolean(els.courseStart.value);
    const eOk = Boolean(els.courseEnd.value);
    if (!sOk) errs.push("Enter course start date (as per CAS).");
    if (!eOk) errs.push("Enter course end date (as per CAS).");
    markField(els.courseStart, !sOk);
    markField(els.courseEnd, !eOk);

    if (sOk && eOk){
      const sd = new Date(els.courseStart.value);
      const ed = new Date(els.courseEnd.value);
      if (sd > ed) { errs.push("Course end date must be after start date."); markField(els.courseEnd, true); }
    }

    const appOk = Boolean(els.applicationDate.value);
    if (!appOk) errs.push("Enter visa application date (required).");
    markField(els.applicationDate, !appOk);
  }
  if (tabId === "t2"){
    // Fees can be 0 but must be numbers; no mandatory besides dependants default 0
  }
  if (tabId === "t3"){
    // At least one funds row required
    const skip = Boolean(els.fundsSkip && els.fundsSkip.checked);
    const rows = getFundsRows();
    if (!skip && !rows.length) errs.push("Add at least one fund row or use the skip option.");
  }
  showErrors(errs);
  return errs.length === 0;
}
function validateCore(){
  if (!els.universitySelect.value) return "Please choose a university from the suggestions.";
  if (!els.courseStart.value || !els.courseEnd.value) return "Please enter course start and end dates.";
  if (new Date(els.courseEnd.value) <= new Date(els.courseStart.value)) return "Course end date must be after the start date.";
  if (!els.applicationDate.value) return "Please enter visa application date (mandatory).";
  return null;
}

function buildIhsCalc(ihs){
  if (!ihs) return "-";
  const parts = [];
  if (ihs.yearlyCharges > 0) parts.push(`${fmtMoney("GBP", ihs.rateYearlyGbp)} x ${ihs.yearlyCharges} year${ihs.yearlyCharges > 1 ? "s" : ""}`);
  if (ihs.halfYearCharges > 0) parts.push(`${fmtMoney("GBP", ihs.rateHalfGbp)} x ${ihs.halfYearCharges} half-year${ihs.halfYearCharges > 1 ? "s" : ""}`);
  const perPerson = parts.length ? `${parts.join(" + ")} = ${fmtMoney("GBP", ihs.ihsPerPersonGbp)} per person` : "No IHS";
  const persons = ihs.persons ? `; x ${ihs.persons} person${ihs.persons > 1 ? "s" : ""}` : "";
  const visaDate = ihs.visaEndDate ? new Date(ihs.visaEndDate) : null;
  const visaLabel = visaDate && !Number.isNaN(visaDate.getTime())
    ? formatDateDisplay(visaDate)
    : (ihs.visaEndDate || "-");
  return `Visa end: ${visaLabel} (${ihs.totalStayMonths} months), ${perPerson}${persons}`;
}

async function calculate(){
  const err = validateCore();
  if (err) { alert(err); return; }

  const out = await apiFetch("/api/report", { method:"POST", headers:{ "content-type":"application/json" }, body: JSON.stringify(payload()) }).then(r=>r.json());
  if (out.error) { alert(out.error); return; }

  const quote = els.quote.value;
  const gbpToQuote = await fx("GBP", quote);

  const gapEligible = Number(out.gapEligibleOnlyGbp) || 0;
  const gapLabel = gapEligible >= 0 ? "Funds are sufficient" : "Additional funds required";
  const gapDisplay = gapEligible >= 0 ? gapEligible : Math.abs(gapEligible);

  if (els.gapLabel) els.gapLabel.textContent = gapLabel;
  els.ihsGbp.textContent = fmtMoney("GBP", out.ihs.ihsTotalGbp);
  if (els.ihsCalc) els.ihsCalc.textContent = buildIhsCalc(out.ihs);
  els.ihsFx.textContent = `${fmtMoney(quote, out.ihs.ihsTotalGbp * gbpToQuote)} (1 GBP = ${gbpToQuote.toFixed(4)} ${quote})`;

  els.fundsReqGbp.textContent = fmtMoney("GBP", out.fundsRequired.fundsRequiredGbp);
  els.fundsReqFx.textContent = fmtMoney(quote, out.fundsRequired.fundsRequiredGbp * gbpToQuote);

  els.fundsAvailGbp.textContent = fmtMoney("GBP", out.fundsAvailable.summary.totalEligibleGbp);
  els.fundsAvailFx.textContent = fmtMoney(quote, out.fundsAvailable.summary.totalEligibleGbp * gbpToQuote);

  els.gapGbp.textContent = fmtMoney("GBP", gapDisplay);
  els.gapFx.textContent = fmtMoney(quote, gapDisplay * gbpToQuote);

  // Eligibility status (simple for students/parents)
  if (els.eligibilityStatus){
    const ok = Number(out.gapEligibleOnlyGbp) >= 0;
    els.eligibilityStatus.textContent = ok ? "Status: ELIGIBLE (Funds are sufficient)" : "Status: NOT ELIGIBLE (Funds are short)";
    els.eligibilityStatus.classList.remove("eligible","noteligible","neutral");
    els.eligibilityStatus.classList.add(ok ? "eligible" : "noteligible");
  }

  els.bTuition.textContent = fmtMoney("GBP", out.fundsRequired.tuitionDueGbp);
  els.bStudent.textContent = fmtMoney("GBP", out.fundsRequired.maintenanceStudentGbp);
  els.bDeps.textContent = fmtMoney("GBP", out.fundsRequired.maintenanceDependantsGbp);
  els.bBuffer.textContent = fmtMoney("GBP", out.fundsRequired.bufferGbp);
  els.bTotal.textContent = fmtMoney("GBP", out.fundsRequired.fundsRequiredGbp);

  els.fundsAvailBreakBody.innerHTML = "";
  const fundTypeLabel = (t) => {
    const ft = String(t || "bank").toLowerCase();
    if (ft === "fd") return "Fixed deposit";
    if (ft === "loan") return "Education loan";
    return "Bank statement";
  };
  if (out.fundsAvailable.summary.skipped) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="9">Funds section skipped by user.</td>`;
    els.fundsAvailBreakBody.appendChild(tr);
  } else {
    out.fundsAvailable.rows.forEach(r=>{
      const tr = document.createElement("tr");
      tr.className = r.eligible ? "" : "row-bad";
      const dateText = r.dateLabel ? `${r.dateLabel}: ${r.dateValue || "-"}` : (r.dateValue || "-");
      tr.innerHTML = `
        <td>${r.eligible ? "OK" : "NOT OK"}</td>
        <td>${escapeHtml(fundTypeLabel(r.fundType))}</td>
        <td>${escapeHtml(r.accountType || "")}</td>
        <td>${escapeHtml(r.source || "")}</td>
        <td>${escapeHtml(fmtMoney(r.currency, r.amount))}</td>
        <td>${escapeHtml(r.currency || "")}</td>
        <td>${escapeHtml(fmtMoney("GBP", r.amountGbp))}</td>
        <td>${escapeHtml(dateText)}</td>
        <td>${escapeHtml((r.issues || []).join("; "))}</td>
      `;
      els.fundsAvailBreakBody.appendChild(tr);
    });
  }

  // Validation/Warn panel
  const s = out.fundsAvailable.summary;
  const warns = [];
  if (s.skipped) warns.push("Funds section skipped - available funds treated as GBP 0.");
  if (!s.hasApplicationDate) warns.push("Application date not provided - freshness checks are NOT verified.");
  warns.push("31-day statement freshness is checked only for bank statements when the visa application date is entered.");
  if (s.anyRowMissingDates) warns.push("Some fund rows are missing required dates - those rows are treated as NOT eligible.");
  if (s.anyIneligibleRows) warns.push("Some fund rows are NOT eligible due to rule violations (see breakdown).");
  if (s.totalEligibleGbp < out.fundsRequired.fundsRequiredGbp) warns.push("Eligible funds are LESS than required funds - student may be ineligible (based on entered data).");

  els.validationNote.innerHTML = warns.length ? ("<strong>Validation warnings:</strong><br>" + warns.map(w=>`- ${escapeHtml(w)}`).join("<br>")) : "<strong>Validation:</strong> All entered fund rows meet the checks (based on provided dates).";

  if (els.issueSummary){
    const issues = new Map();
    out.fundsAvailable.rows.forEach((r) => {
      if (!r.issues || !r.issues.length) return;
      if (r.eligible) return;
      r.issues.forEach((i) => issues.set(i, (issues.get(i) || 0) + 1));
    });
    const top = [...issues.entries()].sort((a, b) => b[1] - a[1]);
    if (s.skipped) {
      els.issueSummary.style.display = "block";
      els.issueSummary.innerHTML = "<strong>Issue summary:</strong> Funds section skipped by user.";
    } else if (top.length) {
      els.issueSummary.style.display = "block";
      els.issueSummary.innerHTML = "<strong>Issue summary:</strong><br>" + top.map(([msg, count]) => `- ${escapeHtml(msg)} (${count})`).join("<br>");
    } else {
      els.issueSummary.style.display = "none";
      els.issueSummary.textContent = "";
    }
  }

  els.rulesNote.innerHTML =
    `<strong>Bank statements:</strong> Funds must be held for <strong>${out.rules.funds_hold_days} consecutive days</strong> and end within <strong>${out.rules.statement_age_days} days</strong> of the visa application date. ` +
    `<strong>FDs:</strong> Maturity date is required (28/31-day rule does not apply). ` +
    `<strong>Education loans:</strong> Disbursement letter should be within <strong>${out.rules.loan_letter_max_age_days || 180} days</strong> of application. ` +
    `Maintenance months cap applied: <strong>${out.fundsRequired.monthsRequired}</strong>.`;
}

function escapeHtml(str){
  return String(str).replace(/[&<>"']/g, s => ({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;" }[s]));
}

async function downloadPdf(){
  const err = validateCore();
  if (err) { alert(err); return; }

  const res = await apiFetch("/api/pdf", { method:"POST", headers:{ "content-type":"application/json" }, body: JSON.stringify(payload()) });
  if (!res.ok) {
    const t = await res.text();
    alert("PDF error: " + t);
    return;
  }
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "UK_Visa_IHS_Funds_Report.pdf";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function resetAll(){
  // Clear inputs
  els.courseStart.value = "";
  els.courseEnd.value = "";
  els.applicationDate.value = "";
  if (els.universityInput) els.universityInput.value = "";
  els.universitySelect.value = "";
  SELECTED_REGION = "outside_london";
  els.regionDisplay.value = "Outside London";
  els.quote.value = "INR";
  QUOTE_MANUAL = false;
  CURRENT_STUDENT = null;
  CURRENT_COUNSELOR = null;

  // Fees
  els.tuitionTotal.value = 0;
  els.tuitionPaid.value = 0;
  els.scholarship.value = 0;
  els.dependants.value = 0;
  els.buffer.value = 0;

  // Funds rows
  els.fundsTbody.innerHTML = "";
  if (els.fundsSkip) els.fundsSkip.checked = false;

  // Results placeholders
  ["ihsGbp","ihsCalc","ihsFx","fundsReqGbp","fundsReqFx","fundsAvailGbp","fundsAvailFx","gapGbp","gapFx","bTuition","bStudent","bDeps","bBuffer","bTotal"]
    .forEach(id => $(id).textContent = "-");
  els.fundsAvailBreakBody.innerHTML = "";
  els.validationNote.textContent = "";
  if (els.issueSummary) {
    els.issueSummary.textContent = "";
    els.issueSummary.style.display = "none";
  }
  els.rulesNote.textContent = "";
  if (els.eligibilityStatus){
    els.eligibilityStatus.textContent = "Status: -";
    els.eligibilityStatus.className = "statusbadge neutral";
  }
  if (els.gapLabel) els.gapLabel.textContent = "Gap (Eligible - Required)";
  if (els.uniSuggest) els.uniSuggest.style.display = "none";
  if (els.studentAck) els.studentAck.value = "";
  if (els.studentName) els.studentName.value = "";
  if (els.studentProgram) els.studentProgram.value = "";
  if (els.studentStatus) els.studentStatus.value = "";
  if (els.studentIntakeYear) els.studentIntakeYear.value = "";
  if (els.studentCityCountry) els.studentCityCountry.value = "";
  if (els.studentSuggest) els.studentSuggest.style.display = "none";
  if (els.counselorInput) els.counselorInput.value = "";
  if (els.counselorEmail) els.counselorEmail.value = "";
  if (els.counselorRegion) els.counselorRegion.value = "";
  if (els.counselorDesignation) els.counselorDesignation.value = "";
  if (els.counselorSuggest) els.counselorSuggest.style.display = "none";

  showErrors([]);
  toggleFundsSkip();
}

els.btnCalc.addEventListener("click", calculate);
els.btnPdf.addEventListener("click", downloadPdf);
els.btnReset.addEventListener("click", resetAll);

loadConfig().then(()=>{ els.regionDisplay.value = "Outside London"; }).catch(()=> { if (els.ratePanel) els.ratePanel.textContent = "Could not load configuration."; });

async function apiFetch(url, opts = {}) {
  const code = localStorage.getItem("kc_access_code") || "";
  const headers = { ...(opts.headers || {}), ...(code ? { "x-access-code": code } : {}) };
  const res = await fetch(url, { ...opts, headers });
  if (res.status === 401) {
    // show modal
    if (els.accessModal) els.accessModal.style.display = "flex";
  }
  return res;
}

function initAccessCode(){
  if (!els.accessCodeSave || !els.accessCodeInput) return;
  const showError = (msg) => {
    if (!els.accessCodeError) return;
    els.accessCodeError.textContent = msg;
    els.accessCodeError.style.display = msg ? "block" : "none";
  };
  const save = () => {
    const code = String(els.accessCodeInput.value || "").trim();
    if (!code) {
      showError("Access code is required.");
      return;
    }
    localStorage.setItem("kc_access_code", code);
    showError("");
    if (els.accessModal) els.accessModal.style.display = "none";
  };
  els.accessCodeSave.addEventListener("click", save);
  els.accessCodeInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") save();
  });
}




initAccessCode();
initStudentSearch();
initCounselorSearch();
toggleFundsSkip();


// Step navigation guard
document.querySelectorAll("[data-next]").forEach(btn=>{
  btn.addEventListener("click", (e)=>{
    const next = btn.getAttribute("data-next");
    const active = document.querySelector(".tabpane.active")?.id || "t1";
    if (!validateTab(active)) { e.preventDefault(); return; }
    // existing nav handler already switches; if not, switch here:
  });
});

function captureNextGuard(){
  document.addEventListener("click", (e)=>{
    const t = e.target;
    if (!(t instanceof HTMLElement)) return;
    const btn = t.closest("[data-next]");
    if (!btn) return;
    const active = document.querySelector(".tabpane.active")?.id || "t1";
    if (!validateTab(active)){
      e.preventDefault();
      e.stopPropagation();
    }
  }, true);
}
captureNextGuard();
