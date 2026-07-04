#!/usr/bin/env node

/**
 * Generate an opt-in dev manifest for running the dev server behind a local
 * HTTPS reverse proxy such as portless (see docs/portless.md).
 *
 * Usage:
 *   npm run manifest:dev                          # → https://pi-excel.localhost
 *   npm run manifest:dev -- my-addin.localhost    # custom hostname
 *   DEV_HOST=pi-excel.localhost node scripts/generate-dev-manifest.mjs
 *
 * Replaces the default dev base URL (https://localhost:3000) in manifest.xml
 * with the proxy origin and writes manifest.dev.xml (fixed output path).
 * manifest.xml itself stays untouched and nothing is published to public/.
 *
 * Notes:
 * - Office add-ins require HTTPS, so the proxy origin is always https://.
 * - The add-in Id is unchanged: sideloading manifest.dev.xml replaces the
 *   default localhost:3000 sideload (they are the same add-in to Excel).
 */

import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

export const DEV_BASE_URL = "https://localhost:3000";
export const DEFAULT_DEV_PROXY_HOST = "pi-excel.localhost";

/**
 * Resolve the HTTPS origin for the dev proxy.
 *
 * Precedence: explicit CLI arg > DEV_HOST env > PORTLESS_URL env (injected by
 * portless) > default (`pi-excel.localhost`, matching `npm run dev:portless`).
 *
 * Accepts a bare hostname ("pi-excel.localhost"), host:port, or a full
 * https:// URL. Throws on anything that is not a clean https origin.
 */
export function resolveDevOrigin({ arg, env = {} } = {}) {
  const candidates = [
    { source: "argument", value: arg },
    { source: "DEV_HOST", value: env.DEV_HOST },
    { source: "PORTLESS_URL", value: env.PORTLESS_URL },
  ];

  let source = "default";
  let raw = DEFAULT_DEV_PROXY_HOST;

  for (const candidate of candidates) {
    const trimmed = typeof candidate.value === "string" ? candidate.value.trim() : "";
    if (trimmed.length > 0) {
      source = candidate.source;
      raw = trimmed;
      break;
    }
  }

  const withScheme = raw.includes("://") ? raw : `https://${raw}`;

  let url;
  try {
    url = new URL(withScheme);
  } catch {
    throw new Error(`Invalid dev proxy host (${source}): ${raw}`);
  }

  if (url.protocol !== "https:") {
    throw new Error(`Dev proxy origin must be https:// (${source}): ${raw}`);
  }

  if (url.username || url.password || url.search || url.hash || url.pathname !== "/") {
    throw new Error(`Dev proxy origin must be a bare https origin without path/query/credentials (${source}): ${raw}`);
  }

  if (!url.hostname) {
    throw new Error(`Dev proxy origin missing hostname (${source}): ${raw}`);
  }

  const origin = url.origin;

  if (origin === DEV_BASE_URL) {
    throw new Error(
      `Dev proxy origin is already the default dev URL (${DEV_BASE_URL}); nothing to generate. Use manifest.xml directly.`,
    );
  }

  return { origin, source };
}

/**
 * Replace every occurrence of the default dev base URL with the proxy origin.
 * Throws if the template does not contain the dev base URL (wrong input file).
 */
export function renderDevManifest(xml, origin) {
  if (!xml.includes(DEV_BASE_URL)) {
    throw new Error(`Input manifest does not include expected dev base URL ${DEV_BASE_URL}`);
  }

  return xml.split(DEV_BASE_URL).join(origin);
}

function fail(message) {
  console.error(`[pi-for-excel] ${message}`);
  process.exit(1);
}

function isCurrentModuleEntrypoint() {
  const entrypoint = process.argv[1];
  if (typeof entrypoint !== "string") {
    return false;
  }

  return pathToFileURL(entrypoint).href === import.meta.url;
}

if (isCurrentModuleEntrypoint()) {
  let resolved;
  try {
    resolved = resolveDevOrigin({ arg: process.argv[2], env: process.env });
  } catch (error) {
    fail(error instanceof Error ? error.message : String(error));
  }

  const repoRoot = path.resolve(process.cwd());
  const inPath = path.join(repoRoot, "manifest.xml");
  // Fixed output path by design: this script must never write anywhere else.
  const outPath = path.join(repoRoot, "manifest.dev.xml");

  if (!fs.existsSync(inPath)) {
    fail(`Missing input manifest at ${inPath}`);
  }

  const xml = fs.readFileSync(inPath, "utf-8");

  let replaced;
  try {
    replaced = renderDevManifest(xml, resolved.origin);
  } catch (error) {
    fail(error instanceof Error ? error.message : String(error));
  }

  fs.writeFileSync(outPath, replaced);
  console.log(`[pi-for-excel] Wrote ${outPath} (${resolved.origin}, from ${resolved.source})`);
  console.log(`[pi-for-excel] Validate with: node scripts/validate-manifest.mjs ${path.basename(outPath)}`);
  console.log(`[pi-for-excel] Sideload it the same way as manifest.xml — see docs/portless.md`);
}
