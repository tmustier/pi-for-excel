import {
  DEFAULT_PROXY_URL,
  validateOfficeProxyUrl,
} from "../auth/proxy-validation.js";
import type { SettingsReader } from "../storage/local/settings-store.js";

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
  settings: SettingsReader,
): Promise<string | undefined> {
  try {
    const enabledRaw = await settings.get("proxy.enabled");
    if (!parseEnabledFlag(enabledRaw)) return undefined;

    const proxyUrlRaw = await settings.get("proxy.url");
    const trimmed = typeof proxyUrlRaw === "string" ? proxyUrlRaw.trim() : "";
    const candidateUrl = trimmed.length > 0 ? trimmed : DEFAULT_PROXY_URL;
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

const PROXY_START_COMMAND = "npx pi-for-excel-proxy";

const PROXY_DOWN_TRANSPORT_PATTERNS = [
  "load failed",
  "failed to fetch",
  "networkerror when attempting to fetch resource",
  "econnrefused",
  "err_connection_refused",
  "connection refused",
] as const;

const PROXY_REACHABLE_FAILURE_PATTERNS = [
  "proxy error:",
  "request failed (",
] as const;

export function isLikelyProxyConnectionError(
  errorMessage: string,
  proxyBaseUrl: string | undefined,
): boolean {
  if (!proxyBaseUrl) return false;
  const lower = errorMessage.trim().toLowerCase();

  if (PROXY_REACHABLE_FAILURE_PATTERNS.some((pattern) => lower.includes(pattern))) {
    return false;
  }

  if (PROXY_DOWN_TRANSPORT_PATTERNS.some((pattern) => lower.includes(pattern))) {
    return true;
  }

  // Node fetch transport failures often collapse to exactly "fetch failed".
  // Keep this fallback strict so upstream/server messages that merely contain
  // those words are not misclassified as proxy-down.
  return lower === "fetch failed";
}

export function buildProxyDownErrorMessage(
  toolLabel: string,
  originalError: string,
): string {
  return (
    `Error: ${toolLabel} failed because the local CORS proxy is not running. `
    + `The Excel add-in cannot reach external APIs without it.\n\n`
    + `To fix: run \`${PROXY_START_COMMAND}\` in a terminal and keep that window open.\n\n`
    + `Do not retry — requests will keep failing until the proxy is started.\n\n`
    + `Original error: ${originalError}`
  );
}
