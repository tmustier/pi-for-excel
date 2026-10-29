import fs from "node:fs";
import path from "node:path";
import { pathToFileURL } from "node:url";

export const DEV_BASE_URL = "https://localhost:3141";
export const DEFAULT_DEV_PROXY_HOST = "pi-excel.localhost";

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
