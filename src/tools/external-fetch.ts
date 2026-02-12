/**
 * Helpers for outbound HTTP calls that optionally route through the configured
 * CORS proxy.
 */

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";

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
  if (typeof proxyUrlRaw !== "string") return undefined;
  const trimmed = proxyUrlRaw.trim();
  if (trimmed.length === 0) return undefined;

  try {
    return validateOfficeProxyUrl(trimmed);
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
