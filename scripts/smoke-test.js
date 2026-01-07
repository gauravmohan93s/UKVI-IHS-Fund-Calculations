import { spawn } from "child_process";

const SERVER_URL = "http://localhost:3000";
const START_TIMEOUT_MS = 15000;
const FETCH_TIMEOUT_MS = 15000;

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchJson(url, opts = {}) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), FETCH_TIMEOUT_MS);
  try {
    const res = await fetch(url, { ...opts, signal: ctrl.signal });
    const text = await res.text();
    const data = text ? JSON.parse(text) : null;
    return { res, data, text };
  } finally {
    clearTimeout(t);
  }
}

async function waitForHealth() {
  const deadline = Date.now() + START_TIMEOUT_MS;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`${SERVER_URL}/healthz`);
      if (res.ok) return true;
    } catch (_) {
      // ignore until timeout
    }
    await delay(300);
  }
  throw new Error("Server did not become healthy in time.");
}

async function run() {
  const child = spawn("node", ["server.js"], {
    stdio: "inherit",
    env: process.env,
  });

  try {
    await waitForHealth();

    const config = await fetchJson(`${SERVER_URL}/api/config`);
    if (!config.res.ok || !config.data?.config) {
      throw new Error("Config endpoint failed.");
    }

    const fx = await fetchJson(
      `${SERVER_URL}/api/fx?from=GBP&to=INR&manual_enabled=true&manual_inr_per_gbp=100`
    );
    if (!fx.res.ok || !Number.isFinite(fx.data?.rate)) {
      throw new Error("FX endpoint failed.");
    }

    const payload = {
      routeKey: "student",
      universityName: "University of Manchester",
      region: "outside_london",
      courseStart: "2026-01-01",
      courseEnd: "2026-12-31",
      applicationDate: "2026-01-15",
      tuitionFeeTotalGbp: 18000,
      tuitionFeePaidGbp: 4000,
      scholarshipGbp: 0,
      dependantsCount: 0,
      bufferGbp: 0,
      fundsRows: [
        {
          accountType: "Student",
          source: "Test Bank",
          currency: "GBP",
          amount: 20000,
          statementStart: "2025-12-15",
          statementEnd: "2026-01-12",
        },
      ],
      manualFx: { enabled: false, inrPerGbp: 0 },
    };

    const report = await fetchJson(`${SERVER_URL}/api/report`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!report.res.ok || report.data?.error) {
      throw new Error(`Report endpoint failed: ${report.text}`);
    }

    const pdfRes = await fetch(`${SERVER_URL}/api/pdf`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!pdfRes.ok) throw new Error("PDF endpoint failed.");
    const pdfBuf = await pdfRes.arrayBuffer();
    if (pdfBuf.byteLength < 4) throw new Error("PDF response is empty.");

    console.log("Smoke test OK");
  } finally {
    if (!child.killed) child.kill();
  }
}

run().catch((err) => {
  console.error(err);
  process.exit(1);
});
