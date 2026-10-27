/**
 * Proxy status detection.
 *
 * Periodically checks whether the local HTTPS proxy is reachable and
 * dispatches a custom event so UI surfaces can react.
 */

import {
  DEFAULT_PROXY_URL,
  isLoopbackHostname,
  isLoopbackProxyUrl,
  normalizeProxyUrl,
} from "../auth/proxy-validation.js";
import type { SettingsReader } from "../storage/local/settings-store.js";

const CHECK_INTERVAL_MS = 30_000;
const CHECK_TIMEOUT_MS = 1_500;

export type ProxyState = "detected" | "not-detected" | "unknown";

let currentState: ProxyState = "unknown";
let blockedByBrowser = false;
let intervalId: ReturnType<typeof setInterval> | undefined;

export function getProxyState(): ProxyState {
  return currentState;
}

/**
 * True when the last check could not reach a localhost proxy because the
 * browser forbids this frame from reaching the local device at all, so
 * running the proxy cannot help.
 */
export function isProxyBlockedByBrowser(): boolean {
  return blockedByBrowser;
}

/**
 * Chromium's Local Network Access gates requests from a public page to
 * localhost behind the `loopback-network` permission, which a cross-origin
 * frame can only ask for when every embedding frame delegates it. Excel on
 * the web does not, so requests fail with "Failed to fetch" and no prompt.
 * Only Chromium exposes `document.featurePolicy`; elsewhere this is false.
 */
function isLoopbackBlockedByFramePolicy(proxyUrl: string): boolean {
  if (!isLoopbackProxyUrl(proxyUrl) || isLoopbackHostname(window.location.hostname)) return false;
  // Not in TypeScript's DOM lib; the shape is Chromium's Feature Policy API.
  const policy = (document as Document & {
    featurePolicy?: { features(): string[]; allowsFeature(feature: string): boolean };
  }).featurePolicy;
  // Browsers without Local Network Access do not know the feature and allow the request.
  if (!policy || !policy.features().includes("loopback-network")) return false;
  return !policy.allowsFeature("loopback-network");
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

/**
 * Probe the proxy and dispatch the state.
 *
 * `force` dispatches even when the state is unchanged. Use it after the proxy
 * settings change, so surfaces that also depend on `proxy.enabled` (such as
 * the proxy banner) re-evaluate straight away.
 */
export async function checkProxyOnce(
  settings: SettingsReader,
  opts: { force?: boolean } = {},
): Promise<ProxyState> {
  let proxyUrl: string = DEFAULT_PROXY_URL;
  try {
    const raw = await settings.get("proxy.url");
    const stored = typeof raw === "string" ? raw.trim() : "";
    if (stored.length > 0) {
      proxyUrl = stored;
    }
  } catch {
    // use default
  }

  const reachable = await probeProxy(proxyUrl);
  const newState: ProxyState = reachable ? "detected" : "not-detected";
  const newBlocked = !reachable && isLoopbackBlockedByFramePolicy(proxyUrl);

  if (newState !== currentState || newBlocked !== blockedByBrowser || opts.force === true) {
    blockedByBrowser = newBlocked;
    currentState = newState;
    dispatchProxyStateChanged(newState);
  }

  return newState;
}

export function startProxyPolling(settings: SettingsReader): () => void {
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
