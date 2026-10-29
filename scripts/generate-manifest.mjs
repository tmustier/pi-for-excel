import fs from "node:fs";
import path from "node:path";

import { DEV_BASE_URL } from "./generate-dev-manifest.mjs";

function fail(msg) {
  console.error(`[pi-for-excel] ${msg}`);
  process.exit(1);
}

const baseUrlRaw = process.env.ADDIN_BASE_URL;
if (!baseUrlRaw) {
  fail("Missing ADDIN_BASE_URL. Example: ADDIN_BASE_URL=\"https://pi-for-excel.vercel.app\"");
}

const baseUrl = baseUrlRaw.trim().replace(/\/+$/, "");

if (!/^https:\/\//i.test(baseUrl)) {
  fail(`ADDIN_BASE_URL must be https:// (got: ${baseUrlRaw})`);
}

let parsed;
try {
  parsed = new URL(baseUrl);
} catch {
  fail(`ADDIN_BASE_URL is not a valid URL: ${baseUrlRaw}`);
}

if (!parsed.hostname) {
  fail(`ADDIN_BASE_URL missing hostname: ${baseUrlRaw}`);
}

const repoRoot = path.resolve(process.cwd());
const inPath = path.join(repoRoot, "manifest.xml");
const outPath = path.join(repoRoot, process.env.OUT || "manifest.prod.xml");
const publicOutPath = path.join(repoRoot, "public", "manifest.prod.xml");

if (!fs.existsSync(inPath)) {
  fail(`Missing input manifest at ${inPath}`);
}

const xml = fs.readFileSync(inPath, "utf-8");
if (!xml.includes(DEV_BASE_URL)) {
  fail(`Input manifest does not include expected dev base URL ${DEV_BASE_URL}`);
}

const replaced = xml.split(DEV_BASE_URL).join(baseUrl);

fs.writeFileSync(outPath, replaced);
console.log(`[pi-for-excel] Wrote ${outPath}`);

fs.mkdirSync(path.dirname(publicOutPath), { recursive: true });
fs.writeFileSync(publicOutPath, replaced);
console.log(`[pi-for-excel] Wrote ${publicOutPath}`);
