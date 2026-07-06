/**
 * Poll the local Pi for Excel proxy for OAuth redirects captured on localhost.
 *
 * Browser-safe OAuth providers still use CLI-compatible localhost redirect URIs
 * such as http://localhost:1455/auth/callback or http://localhost:8085/oauth2callback.
 * When the local proxy helper is running, it also listens on those redirect
 * ports and exposes the captured code back to the taskpane through its HTTPS/CORS endpoint.
 */

import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import {
  isLoopbackProxyUrl,
  resolveConfiguredProxyUrl,
  validateOfficeProxyUrl,
} from "./proxy-validation.js";

export interface OAuthCallbackCapture {
  providerId: string;
  code: string;
  state: string;
  url: string;
  receivedAt: number;
}

interface PollOptions {
  providerId: string;
  state: string;
  timeoutMs?: number;
  intervalMs?: number;
  signal?: AbortSignal;
}

const DEFAULT_TIMEOUT_MS = 2 * 60 * 1000;
const DEFAULT_INTERVAL_MS = 750;

function isAuthOauthCallbackCapturePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isOAuthCallbackCapture(value: DynamicValue): value is OAuthCallbackCapture {
  return (
    isAuthOauthCallbackCapturePayloadShape(value) &&
    value.status === "ready" &&
    typeof value.providerId === "string" &&
    typeof value.code === "string" &&
    typeof value.state === "string" &&
    typeof value.url === "string" &&
    typeof value.receivedAt === "number"
  );
}

function createAbortError(): Error {
  const error = new Error("OAuth callback polling aborted");
  error.name = "AbortError";
  return error;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) {
    throw createAbortError();
  }
}

function wait(ms: number, signal: AbortSignal | undefined): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(createAbortError());
      return;
    }

    let settled = false;
    let timeoutId: ReturnType<typeof setTimeout> | undefined;

    const clear = (onAbort: () => void): void => {
      if (timeoutId !== undefined) {
        clearTimeout(timeoutId);
        timeoutId = undefined;
      }
      signal?.removeEventListener("abort", onAbort);
    };

    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      clear(onAbort);
      reject(createAbortError());
    };

    const finish = (): void => {
      if (settled) return;
      settled = true;
      clear(onAbort);
      resolve();
    };

    timeoutId = setTimeout(finish, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}

async function getEnabledLocalProxyUrl(): Promise<string | null> {
  const storage = getAppStorage();
  const enabled = await storage.settings.get<boolean>("proxy.enabled");
  if (!enabled) return null;

  const rawUrl = await storage.settings.get<string>("proxy.url");
  const proxyUrl = validateOfficeProxyUrl(resolveConfiguredProxyUrl(rawUrl));
  return isLoopbackProxyUrl(proxyUrl) ? proxyUrl : null;
}

function buildCallbackPollUrl(proxyUrl: string, providerId: string, state: string): string {
  const normalizedBase = proxyUrl.endsWith("/") ? proxyUrl : `${proxyUrl}/`;
  const url = new URL(`oauth/callback/${encodeURIComponent(providerId)}`, normalizedBase);
  url.searchParams.set("state", state);
  return url.toString();
}

export async function pollOAuthCallbackCapture(
  options: PollOptions,
): Promise<OAuthCallbackCapture | null> {
  const proxyUrl = await getEnabledLocalProxyUrl().catch(() => null);
  if (!proxyUrl) return null;

  const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
  const intervalMs = options.intervalMs ?? DEFAULT_INTERVAL_MS;
  const deadline = Date.now() + timeoutMs;
  const pollUrl = buildCallbackPollUrl(proxyUrl, options.providerId, options.state);

  while (Date.now() <= deadline) {
    throwIfAborted(options.signal);

    try {
      const requestInit: RequestInit = {
        method: "GET",
        headers: { Accept: "application/json" },
      };
      if (options.signal !== undefined) {
        requestInit.signal = options.signal;
      }
      const response = await fetch(pollUrl, requestInit);

      // 4xx usually means the user is running an older proxy helper, a remote
      // proxy, or a proxy with an origin policy that cannot serve callback
      // capture. Fall back to manual paste immediately.
      if (response.status >= 400 && response.status < 500) {
        return null;
      }

      if (response.ok) {
        const payload: DynamicValue = await response.json();
        if (isOAuthCallbackCapture(payload)) {
          return payload.providerId === options.providerId && payload.state === options.state
            ? payload
            : null;
        }

        if (!isAuthOauthCallbackCapturePayloadShape(payload) || payload.status !== "pending") {
          return null;
        }
      }
    } catch {
      if (options.signal?.aborted) {
        throw createAbortError();
      }
      // Transient proxy/network errors should not break the dialog. Keep the
      // manual paste field available while we retry until the timeout expires.
    }

    await wait(intervalMs, options.signal);
  }

  return null;
}
