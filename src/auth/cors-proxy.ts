/**
 * Fetch interceptor for dev + production.
 *
 * Dev:
 * - rewrites external URLs to Vite's local reverse proxies (/api-proxy/*, /oauth-proxy/*)
 *
 * Production:
 * - does NOT assume any local reverse proxy exists
 * - optionally routes *OAuth/token* endpoints through a user-configured CORS proxy
 *   (<proxy>/?url=<target>) so browser OAuth flows work in Office webviews.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  DEFAULT_LOCAL_PROXY_URL,
  validateOfficeProxyUrl,
} from "./proxy-validation.js";

const DEV_REWRITES: [string, string][] = [
  // OAuth token endpoints
  ["https://console.anthropic.com/", "/oauth-proxy/anthropic/"],
  ["https://github.com/", "/oauth-proxy/github/"],
  ["https://auth.openai.com/", "/api-proxy/openai-auth/"],
  ["https://oauth2.googleapis.com/", "/api-proxy/google-oauth/"],
  // API endpoints
  ["https://api.anthropic.com/", "/api-proxy/anthropic/"],
  ["https://api.openai.com/", "/api-proxy/openai/"],
  ["https://chatgpt.com/", "/api-proxy/chatgpt/"],
  ["https://generativelanguage.googleapis.com/", "/api-proxy/google/"],
  ["https://cloudcode-pa.googleapis.com/", "/api-proxy/google-cloudcode/"],
  ["https://daily-cloudcode-pa.sandbox.googleapis.com/", "/api-proxy/google-cloudcode-sandbox/"],
];

/** The original, un-patched fetch — use for requests that should bypass the proxy */
export let originalFetch: typeof window.fetch;

type ProxySettingsCache = {
  checkedAt: number;
  enabled: boolean;
  url?: string;
};

const proxyCache: ProxySettingsCache = {
  checkedAt: 0,
  enabled: false,
  url: undefined,
};

async function getEnabledProxyUrl(): Promise<string | undefined> {
  // OAuth flows are infrequent, but fetch() is frequent; cache for a short time.
  const now = Date.now();
  if (now - proxyCache.checkedAt < 3000) {
    return proxyCache.enabled ? proxyCache.url : undefined;
  }

  proxyCache.checkedAt = now;

  let enabled: unknown;
  let url: unknown;

  try {
    const storage = getAppStorage();
    enabled = await storage.settings.get("proxy.enabled");
    url = await storage.settings.get("proxy.url");
  } catch {
    proxyCache.enabled = false;
    proxyCache.url = undefined;
    return undefined;
  }

  proxyCache.enabled = Boolean(enabled);
  if (!proxyCache.enabled) {
    proxyCache.url = undefined;
    return undefined;
  }

  const trimmed = typeof url === "string" ? url.trim() : "";
  const candidateUrl = trimmed.length > 0 ? trimmed : DEFAULT_LOCAL_PROXY_URL;

  // Guardrails: validate proxy URL (and fail fast for mixed-content HTTP proxies).
  // This may throw and should surface to the caller.
  const validated = validateOfficeProxyUrl(candidateUrl);
  proxyCache.url = validated;
  return validated;
}

function looksLikeOAuthOrTokenEndpoint(url: string): boolean {
  // Proxy only endpoints that are known to be blocked by CORS in browsers.
  // Keep this conservative so normal fetch() calls aren't routed unexpectedly.
  try {
    const u = new URL(url);

    // Anthropic OAuth token exchange / refresh
    if (u.hostname === "console.anthropic.com" && u.pathname.startsWith("/v1/oauth/token")) return true;

    // GitHub device flow + token exchange (supports enterprise domains too)
    if (u.pathname === "/login/device/code") return true;
    if (u.pathname === "/login/oauth/access_token") return true;

    // GitHub Copilot token endpoint (github.com or GHE)
    if (u.pathname.includes("/copilot_internal/")) return true;

    // OpenAI auth endpoints (not all are browser-friendly)
    if (u.hostname === "auth.openai.com" && u.pathname.startsWith("/oauth/")) return true;

    // Google OAuth token endpoint
    if (u.hostname === "oauth2.googleapis.com") return true;

    // Google Cloud Code Assist onboarding/auth-adjacent endpoints
    if (u.hostname === "cloudcode-pa.googleapis.com" && u.pathname.startsWith("/v1internal")) return true;
    if (u.hostname === "daily-cloudcode-pa.sandbox.googleapis.com" && u.pathname.startsWith("/v1internal")) return true;

    return false;
  } catch {
    return false;
  }
}

function stripAnthropicBrowserHeader(init?: RequestInit): RequestInit | undefined {
  if (!init?.headers) return init;
  const headers = new Headers(init.headers);
  headers.delete("anthropic-dangerous-direct-browser-access");
  return { ...init, headers };
}

/**
 * Install the fetch interceptor. Call once at boot.
 */
export function installFetchInterceptor(): void {
  originalFetch = window.fetch.bind(window);

  window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = typeof input === "string"
      ? input
      : input instanceof URL
        ? input.toString()
        : input.url;

    // Relative URLs: never rewrite
    if (!/^https?:\/\//i.test(url)) {
      return originalFetch(input, init);
    }

    // Dev: Vite reverse proxies
    if (import.meta.env.DEV) {
      let rewritten: string | null = null;
      for (const [prefix, proxy] of DEV_REWRITES) {
        if (url.startsWith(prefix)) {
          rewritten = url.replace(prefix, proxy);
          break;
        }
      }

      if (!rewritten) return originalFetch(input, init);

      const newInit = stripAnthropicBrowserHeader(init);

      if (typeof input !== "string" && !(input instanceof URL) && input instanceof Request) {
        const newHeaders = new Headers(input.headers);
        newHeaders.delete("anthropic-dangerous-direct-browser-access");
        input = new Request(rewritten, { ...input, headers: newHeaders });
      } else {
        input = rewritten;
      }

      return originalFetch(input, newInit);
    }

    // Production: proxy OAuth/token endpoints through user-configured CORS proxy.
    if (looksLikeOAuthOrTokenEndpoint(url)) {
      const proxyUrl = await getEnabledProxyUrl();
      if (proxyUrl) {
        const proxied = `${proxyUrl}/?url=${encodeURIComponent(url)}`;
        return originalFetch(proxied, stripAnthropicBrowserHeader(init));
      }
    }

    return originalFetch(input, init);
  };
}
