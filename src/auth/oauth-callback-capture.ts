import { Type, type Static } from "typebox";
import { Value } from "typebox/value";

import { getAppStorage } from "../storage/local/app-storage.js";
import {
  isLoopbackProxyUrl,
  resolveConfiguredProxyUrl,
  validateOfficeProxyUrl,
} from "./proxy-validation.js";

const callbackCaptureSchema = Type.Object({
  status: Type.Literal("ready"),
  providerId: Type.String(),
  code: Type.String(),
  state: Type.String(),
  url: Type.String(),
  receivedAt: Type.Number(),
});
const callbackResponseSchema = Type.Union([
  callbackCaptureSchema,
  Type.Object({ status: Type.Literal("pending") }),
]);

export type OAuthCallbackCapture = Omit<Static<typeof callbackCaptureSchema>, "status">;

interface PollOptions {
  providerId: string;
  state: string;
  timeoutMs?: number;
  intervalMs?: number;
  signal?: AbortSignal;
}

const DEFAULT_TIMEOUT_MS = 2 * 60 * 1000;
const DEFAULT_INTERVAL_MS = 750;

export async function pollOAuthCallbackCapture(
  options: PollOptions,
): Promise<OAuthCallbackCapture | null> {
  let proxyUrl: string;
  try {
    const { settings } = getAppStorage();
    if (!await settings.get("proxy.enabled")) return null;
    proxyUrl = validateOfficeProxyUrl(resolveConfiguredProxyUrl(await settings.get("proxy.url")));
    if (!isLoopbackProxyUrl(proxyUrl)) return null;
  } catch {
    return null;
  }

  const normalizedBase = proxyUrl.endsWith("/") ? proxyUrl : `${proxyUrl}/`;
  const pollUrl = new URL(`oauth/callback/${encodeURIComponent(options.providerId)}`, normalizedBase);
  pollUrl.searchParams.set("state", options.state);

  const timeout = new AbortController();
  const timeoutId = setTimeout(() => timeout.abort(), options.timeoutMs ?? DEFAULT_TIMEOUT_MS);
  const signal = options.signal ? AbortSignal.any([options.signal, timeout.signal]) : timeout.signal;

  try {
    while (!signal.aborted) {
      try {
        const response = await fetch(pollUrl.toString(), {
          method: "GET",
          headers: { Accept: "application/json" },
          signal,
        });
        if (signal.aborted) break;
        if (response.status >= 400 && response.status < 500) return null;

        if (response.ok) {
          const payload: unknown = await response.json();
          if (signal.aborted) break;
          if (!Value.Check(callbackResponseSchema, payload)) return null;
          if (payload.status === "ready") {
            return payload.providerId === options.providerId && payload.state === options.state
              ? payload
              : null;
          }
        }
      } catch {
      }

      if (signal.aborted) break;
      await new Promise<void>((resolve) => {
        const onAbort = (): void => {
          clearTimeout(timer);
          resolve();
        };
        const timer = setTimeout(() => {
          signal.removeEventListener("abort", onAbort);
          resolve();
        }, options.intervalMs ?? DEFAULT_INTERVAL_MS);
        signal.addEventListener("abort", onAbort, { once: true });
      });
    }

    if (options.signal?.aborted) {
      const error = new Error("OAuth callback polling aborted");
      error.name = "AbortError";
      throw error;
    }
    return null;
  } finally {
    clearTimeout(timeoutId);
  }
}
