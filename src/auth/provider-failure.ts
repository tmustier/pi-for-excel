import { createAssistantMessageEventStream, isRetryableAssistantError, type AssistantMessage, type AssistantMessageEventStream, type StreamOptions } from "@earendil-works/pi-ai";
import { Type } from "typebox";
import { Value } from "typebox/value";

const errorBodySchema = Type.Object({
  error: Type.Object({ code: Type.Optional(Type.String()), type: Type.Optional(Type.String()) }),
});

interface ProviderFailure {
  status: number;
  code: string;
  retryAfterMs: number | null;
}

// Request-local HTTP facts survive the provider's friendly-error rewriting.
const failures = new WeakMap<AssistantMessage, ProviderFailure>();

export function providerRetryDelay(message: AssistantMessage, fallbackMs: number): number | null {
  const failure = failures.get(message);
  const code = failure?.code ?? "";
  if (/usage_limit|usage_not_included|insufficient_quota|quota_exceeded|billing|budget|subscription_sharing_.*limit/iu.test(code)
    || /usage limit|insufficient_quota|quota exceeded|billing|out of budget/iu.test(message.errorMessage ?? "")) {
    // A short rate-limit code takes precedence over Codex's generic usage-limit wording.
    if (code !== "rate_limit_exceeded") return null;
  }
  const retryable = failure?.status === 429 || failure?.status === 529
    || (failure !== undefined && failure.status >= 500 && failure.status <= 599)
    || isRetryableAssistantError(message);
  if (!retryable) return null;
  return Math.min(60_000, failure?.retryAfterMs ?? fallbackMs);
}

/** Capture HTTP error facts at the transport seam, before adapters rewrite them. */
export function streamWithProviderFailure(
  produce: (options: StreamOptions) => AssistantMessageEventStream,
  options?: StreamOptions,
): AssistantMessageEventStream {
  let failure: ProviderFailure | undefined;
  const fetchFn = options?.fetch ?? globalThis.fetch;
  const stream = produce({
    ...options,
    // The action queue owns retries, their countdown and cancellation.
    maxRetries: 0,
    fetch: async (input, init) => {
      const response = await fetchFn(input, init);
      if (!response.ok) {
        const header = response.headers.get("retry-after");
        const seconds = header === null ? Number.NaN : Number(header);
        const delay = Number.isFinite(seconds) ? seconds * 1000 : Date.parse(header ?? "") - Date.now();
        let code = "";
        try {
          const raw: unknown = JSON.parse(await response.clone().text());
          if (Value.Check(errorBodySchema, raw)) code = raw.error.code ?? raw.error.type ?? "";
        } catch { /* Non-JSON errors still carry HTTP status and delay. */ }
        failure = { status: response.status, code, retryAfterMs: Number.isFinite(delay) ? Math.max(0, delay) : null };
      }
      return response;
    },
  });
  const captured = createAssistantMessageEventStream();
  void (async () => {
    for await (const event of stream) {
      if (event.type === "error" && failure) failures.set(event.error, failure);
      captured.push(event);
    }
  })();
  return captured;
}
