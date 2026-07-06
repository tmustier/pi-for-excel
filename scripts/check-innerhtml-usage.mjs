#!/usr/bin/env node
import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";

const root = process.cwd();
const searchRoots = ["src", "tests"];
const allowedFiles = new Set(["src/utils/html.ts"]);
const offenders = [];

function collectFiles(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name === "dist") continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.isFile() && full.endsWith(".ts")) out.push(full);
  }
  return out;
}

for (const dir of searchRoots) {
  for (const file of collectFiles(join(root, dir))) {
    const rel = relative(root, file);
    if (allowedFiles.has(rel)) continue;
    const lines = readFileSync(file, "utf8").split("\n");
    for (const [index, line] of lines.entries()) {
      if (/\.innerHTML\b/u.test(line)) {
        offenders.push(`${rel}:${index + 1}: ${line.trim()}`);
      }
    }
  }
}

if (offenders.length > 0) {
  console.error("innerHTML usage check failed.");
  console.error("Use DOM APIs when practical, or setSafeInnerHTML(...) from src/utils/html.ts with escaped dynamic values and a local safety reason.");
  for (const offender of offenders) console.error(`- ${offender}`);
  process.exit(1);
}

console.log("✓ innerHTML usage check passed.");
