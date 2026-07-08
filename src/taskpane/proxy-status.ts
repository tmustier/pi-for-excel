/**
 * Proxy status detection.
 *
 * Periodically checks whether the local HTTPS proxy is reachable and
 * dispatches a custom event so UI surfaces can react.
 */

import { DEFAULT_PROXY_URL, normalizeProxyUrl } from "../auth/proxy-validation.js";

const CHECK_INTERVAL_MS = 30_000;
const CHECK_TIMEOUT_MS = 1_500;

export type ProxyState = "detected" | "not-detected" | "unknown";

let currentState: ProxyState = "unknown";
let intervalId: ReturnType<typeof setInterval> | undefined;

export function getProxyState(): ProxyState {
  return currentState;
}

async function probeProxy(proxyUrl: string): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CHECK_TIMEOUT_MS);

  try {
    // Health never proxies upstream data and includes CORS headers for allowed
    // add-in origins, so status detection does not depend on target allowlists
    // or a browser accepting an error response body from `/`.
    const url = `${normalizeProxyUrl(proxyUrl)}/healthz`;
    const resp = await fetch(url, { cache: "no-store", signal: controller.signal });
    return resp.ok;
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
  let proxyUrl: string = DEFAULT_PROXY_URL;
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
