#!/usr/bin/env node

/**
 * Generate a production manifest from the dev manifest.
 *
 * Usage:
 *   ADDIN_BASE_URL="https://pi-for-excel.vercel.app" node scripts/generate-manifest.mjs
 *   ADDIN_BASE_URL="https://pi-for-excel.example.com" OUT=manifest.prod.xml node scripts/generate-manifest.mjs
 *
 * Replaces all occurrences of the dev base URL (https://localhost:3000) with ADDIN_BASE_URL.
 *
 * Notes:
 * - Office add-ins require HTTPS.
 * - Keep the production SourceLocation stable so hosted builds can update automatically.
 */

import fs from "node:fs";
import path from "node:path";

const DEV_BASE_URL = "https://localhost:3000";

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
