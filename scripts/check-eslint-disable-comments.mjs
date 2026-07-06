#!/usr/bin/env node
import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";

const root = process.cwd();
const searchRoots = ["src", "tests", "scripts"];
const offenders = [];

function collectFiles(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name === "dist") continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.isFile() && /\.(ts|mjs)$/.test(entry.name)) out.push(full);
  }
  return out;
}

function disableDirectiveBeforeReason(line) {
  const match = line.match(/eslint-disable(?:-next-line|-line)?\s*([^\n]*?)(?:\s--\s|\*\/|$)/u);
  return match?.[1]?.trim() ?? "";
}

for (const dir of searchRoots) {
  for (const file of collectFiles(join(root, dir))) {
    const rel = relative(root, file);
    if (rel === "scripts/check-eslint-disable-comments.mjs") continue;
    const lines = readFileSync(file, "utf8").split("\n");
    for (const [index, line] of lines.entries()) {
      if (!line.includes("eslint-disable")) continue;
      const ruleList = disableDirectiveBeforeReason(line);
      const hasSpecificRule = ruleList.length > 0 && !ruleList.startsWith("--");
      const reason = line.split(/\s--\s/u)[1]?.replace(/\*\//u, "").trim() ?? "";

      if (!hasSpecificRule) {
        offenders.push(`${rel}:${index + 1} disables ESLint without naming a specific rule`);
        continue;
      }
      if (reason.length < 12) {
        offenders.push(`${rel}:${index + 1} eslint-disable needs a concrete reason after \`--\``);
      }
    }
  }
}

if (offenders.length > 0) {
  console.error("ESLint-disable comment check failed.");
  console.error("Suppressions must be narrow and explain the local safety/interop invariant for future agents.");
  for (const offender of offenders) console.error(`- ${offender}`);
  process.exit(1);
}

console.log("✓ ESLint-disable comment check passed.");
