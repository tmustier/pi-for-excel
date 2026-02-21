/**
 * Helpers for outbound HTTP calls that optionally route through the configured
 * CORS proxy.
 */

import {
  DEFAULT_LOCAL_PROXY_URL,
  validateOfficeProxyUrl,
} from "../auth/proxy-validation.js";

export interface ProxyAwareSettingsStore {
  get(key: string): Promise<unknown>;
}

export interface ResolvedOutboundRequest {
  requestUrl: string;
  proxied: boolean;
  proxyBaseUrl?: string;
}

function parseEnabledFlag(value: unknown): boolean {
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return normalized === "1" || normalized === "true" || normalized === "yes";
  }
  if (typeof value === "number") return value !== 0;
  return false;
}

export async function getEnabledProxyBaseUrl(
  settings: ProxyAwareSettingsStore,
): Promise<string | undefined> {
  const enabledRaw = await settings.get("proxy.enabled");
  if (!parseEnabledFlag(enabledRaw)) return undefined;

  const proxyUrlRaw = await settings.get("proxy.url");
  const trimmed = typeof proxyUrlRaw === "string" ? proxyUrlRaw.trim() : "";
  const candidateUrl = trimmed.length > 0 ? trimmed : DEFAULT_LOCAL_PROXY_URL;

  try {
    return validateOfficeProxyUrl(candidateUrl);
  } catch {
    return undefined;
  }
}

function buildProxyRequestUrl(proxyBaseUrl: string, targetUrl: string): string {
  const normalized = proxyBaseUrl.replace(/\/+$/u, "");
  return `${normalized}/?url=${encodeURIComponent(targetUrl)}`;
}

export function resolveOutboundRequestUrl(args: {
  targetUrl: string;
  proxyBaseUrl?: string;
}): ResolvedOutboundRequest {
  const { targetUrl, proxyBaseUrl } = args;

  if (!proxyBaseUrl) {
    return {
      requestUrl: targetUrl,
      proxied: false,
    };
  }

  return {
    requestUrl: buildProxyRequestUrl(proxyBaseUrl, targetUrl),
    proxied: true,
    proxyBaseUrl,
  };
}

/* ── Proxy-down error detection ─────────────────────────────── */

const PROXY_START_COMMAND = "npx pi-for-excel-proxy";

/**
 * Common fetch error messages produced when the local proxy is unreachable.
 * WebKit: "Load failed"; Chrome: "Failed to fetch"; Node: "fetch failed" /
 * "ECONNREFUSED".
 */
const PROXY_DOWN_ERROR_PATTERNS = [
  "load failed",
  "failed to fetch",
  "fetch failed",
  "econnrefused",
  "econnreset",
  "network error",
  "networkerror",
] as const;

/**
 * Returns `true` when the error looks like a connection failure to the local
 * proxy (as opposed to an upstream API error or timeout).
 */
export function isLikelyProxyConnectionError(
  errorMessage: string,
  proxyBaseUrl: string | undefined,
): boolean {
  if (!proxyBaseUrl) return false;
  const lower = errorMessage.toLowerCase();
  return PROXY_DOWN_ERROR_PATTERNS.some((pattern) => lower.includes(pattern));
}

/**
 * Build an error message that is informative for both the agent (so it stops
 * retrying) and the user (so they know how to fix it).
 */
export function buildProxyDownErrorMessage(
  toolLabel: string,
  originalError: string,
): string {
  return (
    `${toolLabel} failed because the local CORS proxy is not running. `
    + `The Excel add-in cannot reach external APIs without it.\n\n`
    + `To fix: run \`${PROXY_START_COMMAND}\` in a terminal and keep that window open.\n\n`
    + `Do not retry — requests will keep failing until the proxy is started.\n\n`
    + `Original error: ${originalError}`
  );
}
