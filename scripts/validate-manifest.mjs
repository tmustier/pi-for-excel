#!/usr/bin/env node

import { spawnSync } from "node:child_process";

const manifestPath = process.argv[2] ?? "manifest.xml";
const maxAttempts = 3;
const retryDelayMs = 5_000;

function sleep(ms) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
}

function isValidationServiceFailure(output) {
  return (
    /Unable to validate the manifest\.\s*5\d\d\b/i.test(output) ||
    /Unable to contact the manifest validation service/i.test(output)
  );
}

let lastOutput = "";

for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
  const result = spawnSync(
    "npx",
    ["office-addin-manifest", "validate", manifestPath],
    { encoding: "utf8" },
  );

  const output = `${result.stdout ?? ""}${result.stderr ?? ""}`;
  lastOutput = output;
  process.stdout.write(result.stdout ?? "");
  process.stderr.write(result.stderr ?? "");

  if (result.status === 0) {
    process.exit(0);
  }

  if (!isValidationServiceFailure(output)) {
    process.exit(result.status ?? 1);
  }

  if (attempt < maxAttempts) {
    console.warn(
      `Manifest validation service failed on attempt ${attempt}/${maxAttempts}; retrying in ${retryDelayMs / 1000}s...`,
    );
    sleep(retryDelayMs);
  }
}

console.warn(
  "::warning::Manifest validation service returned a 5xx/unreachable response after retries; treating this as an external service outage rather than a manifest failure.",
);
console.warn(lastOutput.trim());
process.exit(0);
