import fs from "fs";
import path from "path";

const root = process.cwd();
const url = process.env.STUDENTS_SOURCE_URL || "";
const destPath = process.env.STUDENTS_CSV_PATH || path.join(root, "data", "students.csv");

if (!url) {
  console.error("STUDENTS_SOURCE_URL is not set.");
  process.exit(1);
}

async function fetchCsv(u) {
  const res = await fetch(u, { headers: { "accept": "text/csv" } });
  if (!res.ok) throw new Error(`Fetch failed (${res.status})`);
  return Buffer.from(await res.arrayBuffer());
}

async function main() {
  const tmpPath = `${destPath}.tmp-${Date.now()}`;
  await fs.promises.mkdir(path.dirname(destPath), { recursive: true });
  const buf = await fetchCsv(url);
  await fs.promises.writeFile(tmpPath, buf);
  await fs.promises.rename(tmpPath, destPath);
  console.log(`Synced students CSV -> ${destPath}`);
}

main().catch((e) => {
  console.error(e?.message || e);
  process.exit(1);
});
