#!/usr/bin/env node
/**
 * `unknown` burn-down ratchet.
 *
 * `unknown` is allowed only at the seam that decodes a raw external value;
 * the anti-slop `no-unknown-*` rules flag every parameter, return and alias
 * that lets it travel further. `src` had far too many when the
 * `DynamicValue` alias was deleted to enable those rules as errors, so this
 * check runs them as warnings and fails when the count rises above the
 * committed baseline. Lower the baseline whenever a change reduces the
 * count; never raise it.
 */
import { spawnSync } from "node:child_process";

const BASELINE = 422; // 2026-09-11: 356 no-unknown-parameters, 66 no-unknown-returns, 0 no-unknown-type-aliases

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
// oxlint also reports its default rules; only the no-unknown-* rules count.
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
