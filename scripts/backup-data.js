import fs from "fs";
import path from "path";

const root = process.cwd();
const dataDir = path.join(root, "data");
const backupRoot = process.env.BACKUP_ROOT || path.join(root, "backups");

function ymd() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

async function main() {
  if (!fs.existsSync(dataDir)) {
    console.error("data/ folder not found.");
    process.exit(1);
  }
  const dest = path.join(backupRoot, ymd());
  await fs.promises.mkdir(dest, { recursive: true });
  await fs.promises.cp(dataDir, dest, { recursive: true });
  console.log(`Backup created: ${dest}`);
}

main().catch((e) => {
  console.error(e?.message || e);
  process.exit(1);
});
