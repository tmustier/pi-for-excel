/**
 * Checks an API key before it is saved, by listing the provider's models.
 *
 * Listing models is free and needs only the key. A key the provider rejects
 * (401/403) is not saved. When the check cannot run (unknown provider, offline,
 * blocked by CORS, server error), the key is saved unverified: a network problem
 * now says nothing about the key.
 */

export type ApiKeyCheckResult =
  | { status: "valid" }
  | { status: "rejected"; httpStatus: number }
  | { status: "unverified" };

const CHECK_TIMEOUT_MS = 8_000;
const API_KEY_CHECK_URLS = new Map([
  ["openai", "https://api.openai.com/v1/models"],
  ["anthropic", "https://api.anthropic.com/v1/models"],
  ["google", "https://generativelanguage.googleapis.com/v1beta/models"],
  ["deepseek", "https://api.deepseek.com/models"],
  ["mistral", "https://api.mistral.ai/v1/models"],
  ["groq", "https://api.groq.com/openai/v1/models"],
  ["xai", "https://api.x.ai/v1/models"],
  ["openrouter", "https://openrouter.ai/api/v1/key"],
]);

export async function checkProviderApiKey(
  providerId: string,
  key: string,
  fetchFn: typeof globalThis.fetch = globalThis.fetch,
): Promise<ApiKeyCheckResult> {
  const url = API_KEY_CHECK_URLS.get(providerId);
  if (!url) return { status: "unverified" };

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CHECK_TIMEOUT_MS);
  try {
    const headers = new Headers({ Accept: "application/json" });
    if (providerId === "anthropic") {
      headers.set("x-api-key", key);
      headers.set("anthropic-version", "2023-06-01");
      headers.set("anthropic-dangerous-direct-browser-access", "true");
    } else if (providerId === "google") {
      headers.set("x-goog-api-key", key);
    } else {
      headers.set("Authorization", `Bearer ${key}`);
    }

    const response = await fetchFn(url, {
      headers,
      signal: controller.signal,
      cache: "no-store",
    });
    if (response.ok) return { status: "valid" };
    if (response.status === 401 || response.status === 403) {
      return { status: "rejected", httpStatus: response.status };
    }
    return { status: "unverified" };
  } catch {
    return { status: "unverified" };
  } finally {
    clearTimeout(timeout);
  }
}
