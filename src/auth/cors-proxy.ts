import { getAppStorage } from "../storage/local/app-storage.js";

import {
  DEFAULT_PROXY_URL,
  validateOfficeProxyUrl,
} from "./proxy-validation.js";
import { rewriteDevProxyUrl } from "./dev-rewrites.js";
import type { SettingsReader } from "../storage/local/settings-store.js";

export let originalFetch: typeof window.fetch;

type ProxySettingsCache = {
  checkedAt: number;
  enabled: boolean;
  url?: string;
};

const proxyCache: ProxySettingsCache = {
  checkedAt: 0,
  enabled: false,
};

function parseProxyEnabled(value: unknown): boolean {
  return Boolean(value);
}

export async function readCorsProxySettings(
  settings: SettingsReader,
): Promise<{ enabled: boolean; url: string }> {
  try {
    const [enabled, url] = await Promise.all([
      settings.get("proxy.enabled"),
      settings.get("proxy.url"),
    ]);
    return {
      enabled: parseProxyEnabled(enabled),
      url: typeof url === "string" ? url.trim() : "",
    };
  } catch {
    return { enabled: false, url: "" };
  }
}

async function getEnabledProxyUrl(
  injectedSettings?: SettingsReader,
): Promise<string | undefined> {
  const now = Date.now();
  if (now - proxyCache.checkedAt < 3000) {
    return proxyCache.enabled ? proxyCache.url : undefined;
  }

  proxyCache.checkedAt = now;

  let settings = injectedSettings;
  if (!settings) {
    try {
      settings = getAppStorage().settings;
    } catch {
      proxyCache.enabled = false;
      delete proxyCache.url;
      return undefined;
    }
  }

  const stored = await readCorsProxySettings(settings);
  proxyCache.enabled = stored.enabled;
  if (!proxyCache.enabled) {
    delete proxyCache.url;
    return undefined;
  }

  const candidateUrl = stored.url.length > 0 ? stored.url : DEFAULT_PROXY_URL;

  const validated = validateOfficeProxyUrl(candidateUrl);
  proxyCache.url = validated;
  return validated;
}

function looksLikeOAuthOrTokenEndpoint(url: string): boolean {
  try {
    const u = new URL(url);

    if (
      (u.hostname === "console.anthropic.com" || u.hostname === "platform.claude.com") &&
      u.pathname.startsWith("/v1/oauth/token")
    ) {
      return true;
    }

    if (u.pathname === "/login/device/code") return true;
    if (u.pathname === "/login/oauth/access_token") return true;

    if (u.pathname.includes("/copilot_internal/")) return true;

    if (u.hostname === "auth.openai.com" && u.pathname.startsWith("/oauth/")) return true;

    if (u.hostname === "oauth2.googleapis.com") return true;

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

export function installFetchInterceptor(settings?: SettingsReader): void {
  originalFetch = window.fetch.bind(window);

  window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
    const url = typeof input === "string"
      ? input
      : input instanceof URL
        ? input.toString()
        : input.url;

    if (!/^https?:\/\//i.test(url)) {
      return originalFetch(input, init);
    }

    if (import.meta.env.DEV) {
      const rewritten = rewriteDevProxyUrl(url);
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

    if (looksLikeOAuthOrTokenEndpoint(url)) {
      const proxyUrl = await getEnabledProxyUrl(settings);
      if (proxyUrl) {
        const proxied = `${proxyUrl}/?url=${encodeURIComponent(url)}`;
        return originalFetch(proxied, stripAnthropicBrowserHeader(init));
      }
    }

    return originalFetch(input, init);
  };
}
