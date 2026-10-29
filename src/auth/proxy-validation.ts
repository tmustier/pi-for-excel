/**
 * Proxy URL validation for Office taskpanes.
 *
 * Office add-ins are served over HTTPS. Some Office webviews (notably WKWebView on macOS)
 * will block calls to an HTTP proxy from an HTTPS taskpane (mixed content), surfacing as
 * "Load failed" / "Connection error".
 */

import type { SpreadsheetHostKind } from "../host/types.js";
import { t } from "../language/index.js";

export const DEFAULT_LOCAL_PROXY_URL = "https://localhost:3003";
export const WPS_DEV_HOST_GATEWAY_PROXY_URL = "http://10.0.2.2:3003";
export const PROXY_HELPER_COMMAND = "npx -y pi-for-excel-proxy@latest";

export function resolveDefaultProxyUrl(raw: unknown): string {
  if (typeof raw !== "string") return DEFAULT_LOCAL_PROXY_URL;
  const candidate = normalizeProxyUrl(raw);
  if (candidate.length === 0) return DEFAULT_LOCAL_PROXY_URL;

  if (!/^https:\/\//i.test(candidate)) {
    console.warn(`[pi-for-excel] Ignoring VITE_PI_DEFAULT_PROXY_URL (must be https://): ${candidate}`);
    return DEFAULT_LOCAL_PROXY_URL;
  }

  try {
    new URL(candidate);
  } catch {
    console.warn(`[pi-for-excel] Ignoring VITE_PI_DEFAULT_PROXY_URL (not a valid URL): ${candidate}`);
    return DEFAULT_LOCAL_PROXY_URL;
  }

  return candidate;
}

export const DEFAULT_PROXY_URL = resolveDefaultProxyUrl(
  typeof import.meta.env === "undefined" ? undefined : import.meta.env.VITE_PI_DEFAULT_PROXY_URL,
);

export const DEFAULT_PROXY_IS_REMOTE = !isLoopbackProxyUrl(DEFAULT_PROXY_URL);

export const PROXY_REACHABILITY_TARGET_URL = "https://github.com";
export const PROXY_HEALTH_HEADER = "x-pi-for-excel-proxy";
export const CODEX_WEBSOCKET_BRIDGE_HEADER = "x-pi-for-excel-codex-websocket-bridge";

export const PROXY_HELPER_DOCS_URL =
  "https://github.com/tmustier/pi-for-excel/blob/main/docs/install.md#oauth-logins-and-cors-proxy";

export function normalizeProxyUrl(url: string): string {
  return url.trim().replace(/\/+$/, "");
}

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

export function resolveConfiguredProxyUrl(rawUrl: unknown): string {
  const trimmed = typeof rawUrl === "string" ? rawUrl.trim() : "";
  const candidate = trimmed.length > 0 ? trimmed : DEFAULT_PROXY_URL;
  return normalizeProxyUrl(candidate);
}

export async function probeCodexWebSocketBridge(
  proxyUrl: string,
  timeoutMs: number = 1500,
): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const url = `${normalizeProxyUrl(proxyUrl)}/healthz`;
    const response = await fetch(url, { cache: "no-store", signal: controller.signal });
    return response.ok
      && response.headers.get(PROXY_HEALTH_HEADER) === "1"
      && response.headers.get(CODEX_WEBSOCKET_BRIDGE_HEADER) === "1";
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

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
    throw new Error(t("proxy.error.invalid", { url, example: DEFAULT_PROXY_URL }));
  }

  let parsed: URL;
  try {
    parsed = new URL(normalized);
  } catch {
    throw new Error(t("proxy.error.invalid", { url, example: DEFAULT_PROXY_URL }));
  }

  if (typeof window !== "undefined" && window.location?.protocol === "https:" && parsed.protocol === "http:") {
    throw new Error(t("proxy.error.mixedContent", {
      url: normalized,
      example: DEFAULT_PROXY_URL,
      guide: PROXY_HELPER_DOCS_URL,
    }));
  }

  return normalized;
}

export function isLoopbackHostname(hostname: string): boolean {
  const h = hostname.toLowerCase();
  if (h === "localhost") return true;
  if (h === "::1" || h === "0:0:0:0:0:0:0:1") return true;
  if (h.startsWith("127.")) return true;
  if (h.startsWith("::ffff:127.")) return true;
  return false;
}

export function isLoopbackProxyUrl(url: string): boolean {
  const normalized = normalizeProxyUrl(url);
  try {
    const parsed = new URL(normalized);
    return isLoopbackHostname(parsed.hostname);
  } catch {
    return false;
  }
}
