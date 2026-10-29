import { readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";

const root = process.cwd();
const baselinePath = "scripts/comment-burndown-baseline.json";
const searchRoots = ["src", "tests", "scripts"];
const extraFiles = ["vite.config.ts"];
const publicApiFiles = new Set(["src/commands/extension-api-types.ts"]);
const exemptLine =
  /\/\/\/\s*<reference|oxlint-disable|@ts-expect-error|@vite-ignore|prettier-ignore|@license|SPDX-License|Copyright|©|Vendored from|\bMIT\b/u;

function collectFiles(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name === "dist") continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.isFile() && /\.(ts|mts|mjs|js|cjs|css)$/u.test(entry.name)) out.push(full);
  }
  return out;
}

function commentProseLines(source) {
  const lines = [];
  let block = null;
  const flushBlock = () => {
    if (block && !block.exempt) lines.push(...block.lines);
    block = null;
  };
  for (const [index, raw] of source.split("\n").entries()) {
    const line = raw.trim();
    if (block) {
      const text = line.replace(/^\*+/u, "").replace(/\*\/.*$/u, "").trim();
      if (text.length > 0) block.lines.push(index + 1);
      if (exemptLine.test(line)) block.exempt = true;
      if (line.includes("*/")) flushBlock();
    } else if (line.startsWith("/*")) {
      const text = line.replace(/^\/\*+/u, "").replace(/\*\/.*$/u, "").trim();
      block = { lines: text.length > 0 ? [index + 1] : [], exempt: exemptLine.test(line) };
      if (line.includes("*/")) flushBlock();
    } else if (line.startsWith("//")) {
      const text = line.replace(/^\/+/u, "").trim();
      if (text.length > 0 && !exemptLine.test(line)) lines.push(index + 1);
    }
  }
  flushBlock();
  return lines;
}

const counts = {};
const files = [...searchRoots.flatMap((dir) => collectFiles(join(root, dir))), ...extraFiles.map((file) => join(root, file))];
for (const file of files) {
  const rel = relative(root, file).replace(/\\/g, "/");
  if (publicApiFiles.has(rel)) continue;
  const found = commentProseLines(readFileSync(file, "utf8"));
  if (found.length > 0) counts[rel] = found;
}

const sortedCounts = Object.fromEntries(
  Object.keys(counts).sort().map((file) => [file, counts[file].length]),
);

if (process.argv.includes("--write")) {
  writeFileSync(join(root, baselinePath), `${JSON.stringify(sortedCounts, null, 2)}\n`);
  console.log(`Wrote ${baselinePath} (${Object.keys(sortedCounts).length} files).`);
  process.exit(0);
}

const baseline = JSON.parse(readFileSync(join(root, baselinePath), "utf8"));
const regressions = Object.entries(sortedCounts)
  .filter(([file, count]) => count > (baseline[file] ?? 0))
  .map(([file, count]) => `${file}: ${count} comment lines (baseline ${baseline[file] ?? 0}) at lines ${counts[file].join(", ")}`);

if (regressions.length > 0) {
  console.error("✗ comment burn-down regressed:");
  for (const regression of regressions) console.error(`  ${regression}`);
  console.error("  Make the code say it instead: rename, extract, type, or test. A comment survives only for");
  console.error("  behavior forced by an external platform, vendor, or dependency, a license header, or a public");
  console.error(`  API contract; raise that file's entry in ${baselinePath} and say why in the PR.`);
  process.exit(1);
}

const total = Object.values(sortedCounts).reduce((sum, count) => sum + count, 0);
const baselineTotal = Object.values(baseline).reduce((sum, count) => sum + Number(count), 0);
const stale = Object.keys(baseline).filter((file) => (sortedCounts[file] ?? 0) < baseline[file]);
if (stale.length > 0) {
  console.log(`✓ comment burn-down: ${total} prose lines in ${Object.keys(sortedCounts).length} files; ${stale.length} baseline entries can be lowered with \`node ${relative(root, process.argv[1])} --write\`.`);
} else {
  console.log(`✓ comment burn-down: ${total} prose lines in ${Object.keys(sortedCounts).length} files, at baseline (${baselineTotal}).`);
}
