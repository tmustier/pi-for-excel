/**
 * Proxy URL validation for Office taskpanes.
 *
 * Office add-ins are served over HTTPS. Some Office webviews (notably WKWebView on macOS)
 * will block calls to an HTTP proxy from an HTTPS taskpane (mixed content), surfacing as
 * "Load failed" / "Connection error".
 */

import type { SpreadsheetHostKind } from "../host/types.js";

export const DEFAULT_LOCAL_PROXY_URL = "https://localhost:3003";
export const WPS_DEV_HOST_GATEWAY_PROXY_URL = "http://10.0.2.2:3003";

/**
 * Resolve the build-time default proxy URL override (org/central deployments).
 * Falls back to the local proxy default unless the override is a valid
 * https:// URL — http would be blocked as mixed content in Office webviews,
 * so we refuse it here rather than baking in a broken default.
 */
export function resolveDefaultProxyUrl(raw: DynamicValue): string {
  if (typeof raw !== "string") return DEFAULT_LOCAL_PROXY_URL;
  const candidate = normalizeProxyUrl(raw);
  if (candidate.length === 0) return DEFAULT_LOCAL_PROXY_URL;

  if (!/^https:\/\//i.test(candidate)) {
    console.warn(`[pi-for-excel] Ignoring VITE_PI_DEFAULT_PROXY_URL (must be https://): ${candidate}`);
    return DEFAULT_LOCAL_PROXY_URL;
  }

  try {
    // Validate URL shape; result unused.
    new URL(candidate);
  } catch {
    console.warn(`[pi-for-excel] Ignoring VITE_PI_DEFAULT_PROXY_URL (not a valid URL): ${candidate}`);
    return DEFAULT_LOCAL_PROXY_URL;
  }

  return candidate;
}

/**
 * Effective default proxy URL for this build. Equals DEFAULT_LOCAL_PROXY_URL
 * unless the build sets VITE_PI_DEFAULT_PROXY_URL (org/central deployments —
 * see docs/central-proxy.md).
 */
export const DEFAULT_PROXY_URL = resolveDefaultProxyUrl(
  typeof import.meta.env === "undefined" ? undefined : import.meta.env.VITE_PI_DEFAULT_PROXY_URL,
);

/**
 * True when this build's default proxy is a remote (org/central) proxy.
 * UI copy uses this to swap local-helper instructions ("run npx pi-for-excel-proxy")
 * for org guidance ("contact IT / check settings").
 */
export const DEFAULT_PROXY_IS_REMOTE = !isLoopbackProxyUrl(DEFAULT_PROXY_URL);

/**
 * Target URL used for proxy reachability probes.
 *
 * Must stay inside scripts/cors-proxy-server.mjs DEFAULT_ALLOWED_TARGET_HOSTS,
 * otherwise the helper will return 403 and OAuth preflight checks will fail.
 */
export const PROXY_REACHABILITY_TARGET_URL = "https://github.com";

export const PROXY_HELPER_DOCS_URL =
  "https://github.com/tmustier/pi-for-excel/blob/main/docs/install.md#oauth-logins-and-cors-proxy";

export function normalizeProxyUrl(url: string): string {
  return url.trim().replace(/\/+$/, "");
}

/**
 * Resolve the runtime default proxy URL, including host-specific dev defaults.
 */
export function resolveRuntimeDefaultProxyUrl(options?: {
  hostKind?: SpreadsheetHostKind;
  location?: Pick<Location, "hostname" | "protocol">;
}): string {
  const loc = options?.location ?? (typeof window === "undefined" ? undefined : window.location);
  if (
    options?.hostKind === "wps" &&
    loc?.protocol === "http:" &&
    loc.hostname === "10.0.2.2"
  ) {
    return WPS_DEV_HOST_GATEWAY_PROXY_URL;
  }

  return DEFAULT_PROXY_URL;
}

/**
 * Resolve a user-configured proxy URL with sane defaults.
 */
export function resolveConfiguredProxyUrl(rawUrl: DynamicValue): string {
  const trimmed = typeof rawUrl === "string" ? rawUrl.trim() : "";
  const candidate = trimmed.length > 0 ? trimmed : DEFAULT_PROXY_URL;
  return normalizeProxyUrl(candidate);
}

/**
 * Probe whether a proxy URL is reachable and can forward to the allowlisted
 * reachability target.
 */
export async function probeProxyReachability(
  proxyUrl: string,
  timeoutMs: number = 1500,
): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const url = `${normalizeProxyUrl(proxyUrl)}/?url=${encodeURIComponent(PROXY_REACHABILITY_TARGET_URL)}`;
    const resp = await fetch(url, { signal: controller.signal });
    return resp.ok;
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

export function validateOfficeProxyUrl(url: string): string {
  const normalized = normalizeProxyUrl(url);

  if (!/^https?:\/\//i.test(normalized)) {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_PROXY_URL}`,
    );
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    throw new Error(
      `Invalid Proxy URL: "${url}". Expected a full URL like ${DEFAULT_PROXY_URL}`,
    );
  }

  // Mixed content guardrail: HTTPS taskpane -> HTTP proxy.
  // This tends to fail in Office webviews (macOS), so fail fast with guidance.
  if (typeof window !== "undefined" && window.location?.protocol === "https:" && parsed.protocol === "http:") {
    throw new Error(
      `Proxy URL is HTTP (${normalized}) but the add-in is served over HTTPS. Office webviews may block this as mixed content. ` +
        `Use ${DEFAULT_PROXY_URL} and run an HTTPS proxy. See ${PROXY_HELPER_DOCS_URL}.`,
    );
  }

  return normalized;
}

function isLoopbackHostname(hostname: string): boolean {
  const h = hostname.toLowerCase();
  if (h === "localhost") return true;
  if (h === "::1" || h === "0:0:0:0:0:0:0:1") return true;
  if (h.startsWith("127.")) return true;
  if (h.startsWith("::ffff:127.")) return true;
  return false;
}

/**
 * Returns true if the proxy URL points at a loopback/localhost address.
 * Useful for warning users when they configure a remote proxy.
 */
export function isLoopbackProxyUrl(url: string): boolean {
  const normalized = normalizeProxyUrl(url);
  try {
    const parsed = new URL(normalized);
    return isLoopbackHostname(parsed.hostname);
  } catch {
    return false;
  }
}
