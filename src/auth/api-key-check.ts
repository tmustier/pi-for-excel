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

interface KeyCheckRequest {
  url: string;
  headers: Record<string, string>;
}

const CHECK_TIMEOUT_MS = 8_000;

function bearer(url: string, key: string): KeyCheckRequest {
  return { url, headers: { Authorization: `Bearer ${key}` } };
}

/** Model-list endpoints for providers whose API keys are entered in the taskpane. */
export function buildApiKeyCheckRequest(providerId: string, key: string): KeyCheckRequest | null {
  switch (providerId) {
    case "openai":
      return bearer("https://api.openai.com/v1/models", key);
    case "anthropic":
      return {
        url: "https://api.anthropic.com/v1/models",
        headers: {
          "x-api-key": key,
          "anthropic-version": "2023-06-01",
          "anthropic-dangerous-direct-browser-access": "true",
        },
      };
    case "google":
      return {
        url: "https://generativelanguage.googleapis.com/v1beta/models",
        headers: { "x-goog-api-key": key },
      };
    case "deepseek":
      return bearer("https://api.deepseek.com/models", key);
    case "mistral":
      return bearer("https://api.mistral.ai/v1/models", key);
    case "groq":
      return bearer("https://api.groq.com/openai/v1/models", key);
    case "xai":
      return bearer("https://api.x.ai/v1/models", key);
    case "openrouter":
      return bearer("https://openrouter.ai/api/v1/key", key);
    default:
      return null;
  }
}

export async function checkProviderApiKey(
  providerId: string,
  key: string,
  fetchFn: typeof globalThis.fetch = globalThis.fetch,
): Promise<ApiKeyCheckResult> {
  const request = buildApiKeyCheckRequest(providerId, key);
  if (!request) return { status: "unverified" };

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), CHECK_TIMEOUT_MS);
  try {
    const response = await fetchFn(request.url, {
      headers: { Accept: "application/json", ...request.headers },
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
