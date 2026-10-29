import { spawnSync } from "node:child_process";

const BASELINE = 417;

const result = spawnSync(
  "npx",
  ["oxlint", "-c", ".oxlintrc.unknown-burndown.json", "--format", "json", "src"],
  { encoding: "utf8", shell: process.platform === "win32" },
);
if (result.error) {
  console.error(`Unknown burn-down check could not run oxlint: ${result.error.message}`);
  process.exit(1);
}

const jsonStart = result.stdout.indexOf("{");
if (jsonStart < 0) {
  console.error("Unknown burn-down check: oxlint produced no JSON report.");
  console.error(result.stdout, result.stderr);
  process.exit(1);
}
const report = JSON.parse(result.stdout.slice(jsonStart));
const diagnostics = Array.isArray(report.diagnostics) ? report.diagnostics : [];
const RULE_PATTERN = /^anti-slop\((no-unknown-[a-z-]+)\)$/;
const perRule = new Map();
let total = 0;
for (const diagnostic of diagnostics) {
  const match = RULE_PATTERN.exec(String(diagnostic.code ?? ""));
  if (!match) continue;
  perRule.set(match[1], (perRule.get(match[1]) ?? 0) + 1);
  total += 1;
}
const breakdown = [...perRule.entries()].map(([rule, count]) => `${rule}: ${count}`).join(", ");

if (total > BASELINE) {
  console.error(`✗ unknown burn-down regressed: ${total} sites (${breakdown}), baseline ${BASELINE}.`);
  console.error("  Parse the value into a domain type at the seam instead of passing unknown inward.");
  console.error(`  Inspect with: npx oxlint -c .oxlintrc.unknown-burndown.json src`);
  process.exit(1);
}
if (total < BASELINE) {
  console.log(`✓ unknown burn-down: ${total} sites (${breakdown}); baseline ${BASELINE} can be lowered in scripts/check-unknown-burndown.mjs.`);
} else {
  console.log(`✓ unknown burn-down: ${total} sites (${breakdown}), at baseline.`);
}
