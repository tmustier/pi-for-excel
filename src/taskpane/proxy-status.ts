/**
 * Proxy (local helper) status detection and status bar indicator.
 *
 * Periodically checks whether the local HTTPS proxy is reachable and
 * dispatches a custom event so the status bar can reflect the state.
 */

import { DEFAULT_LOCAL_PROXY_URL } from "../auth/proxy-validation.js";

const DISMISSED_KEY = "pi.proxy.status.dismissed";
const CHECK_INTERVAL_MS = 30_000;
const CHECK_TIMEOUT_MS = 1_500;

export type ProxyState = "detected" | "not-detected" | "unknown";

let currentState: ProxyState = "unknown";
let intervalId: ReturnType<typeof setInterval> | undefined;

export function getProxyState(): ProxyState {
  return currentState;
}

export function isProxyDismissed(): boolean {
  try {
    return localStorage.getItem(DISMISSED_KEY) === "1";
  } catch {
    return false;
  }
}

export function dismissProxyStatus(): void {
  try {
    localStorage.setItem(DISMISSED_KEY, "1");
  } catch {
    // ignore
  }
}

async function probeProxy(proxyUrl: string): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CHECK_TIMEOUT_MS);

  try {
    // Probe without a ?url= parameter — the proxy returns 400 ("Missing target"),
    // which proves it's running. This avoids hitting the target allowlist.
    const url = `${proxyUrl.replace(/\/+$/, "")}/`;
    const resp = await fetch(url, { method: "HEAD", signal: controller.signal });
    // Any response (200, 400, 404) proves the proxy is reachable.
    // Only network errors (fetch throws) indicate it's not running.
    void resp;
    return true;
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

function dispatchProxyStateChanged(state: ProxyState): void {
  document.dispatchEvent(new CustomEvent("pi:proxy-state-changed", { detail: { state } }));
}

interface ProxySettingsReader {
  get<T>(key: string): Promise<T | null>;
}

export async function checkProxyOnce(settings: ProxySettingsReader): Promise<ProxyState> {
  let proxyUrl: string = DEFAULT_LOCAL_PROXY_URL;
  try {
    const raw = await settings.get<string>("proxy.url");
    const stored = typeof raw === "string" ? raw.trim() : "";
    if (stored.length > 0) {
      proxyUrl = stored;
    }
  } catch {
    // use default
  }

  const reachable = await probeProxy(proxyUrl);
  const newState: ProxyState = reachable ? "detected" : "not-detected";

  if (newState !== currentState) {
    currentState = newState;
    dispatchProxyStateChanged(newState);
  }

  return newState;
}

export function startProxyPolling(settings: ProxySettingsReader): () => void {
  void checkProxyOnce(settings);

  intervalId = setInterval(() => {
    void checkProxyOnce(settings);
  }, CHECK_INTERVAL_MS);

  return () => {
    if (intervalId !== undefined) {
      clearInterval(intervalId);
      intervalId = undefined;
    }
  };
}
