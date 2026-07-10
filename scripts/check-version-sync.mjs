#!/usr/bin/env node

import { readFileSync } from "node:fs";

function readJson(path) {
  return JSON.parse(readFileSync(path, "utf8"));
}

function captureVersion(path, pattern, label) {
  const source = readFileSync(path, "utf8");
  const version = source.match(pattern)?.[1];
  if (!version) {
    throw new Error(`Could not read ${label} version from ${path}`);
  }
  return version;
}

const packageJson = readJson("package.json");
const packageLock = readJson("package-lock.json");
const expected = packageJson.version;

const versions = [
  ["package-lock.json", packageLock.version],
  ["package-lock.json root package", packageLock.packages?.[""]?.version],
  [
    "src/app/metadata.ts",
    captureVersion(
      "src/app/metadata.ts",
      /export const APP_VERSION = "([^"]+)";/u,
      "app metadata",
    ),
  ],
  [
    "wps/jsplugins.xml.template",
    captureVersion(
      "wps/jsplugins.xml.template",
      /<jsplugin\s[\s\S]*?\bversion="([^"]+)"/u,
      "WPS package",
    ),
  ],
];

const mismatches = versions.filter(([, version]) => version !== expected);
if (mismatches.length > 0) {
  const details = mismatches
    .map(([location, version]) => `  - ${location}: ${String(version)}`)
    .join("\n");
  throw new Error(`Version mismatch; expected ${expected}:\n${details}`);
}

console.log(`✓ Version metadata is synchronized at ${expected}.`);
