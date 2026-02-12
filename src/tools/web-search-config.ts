/**
 * Web-search configuration shared by tool + settings UI.
 */

export const WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY = "web.search.brave.apiKey";

export type WebSearchProvider = "brave";

export interface WebSearchConfigReader {
  get(key: string): Promise<unknown>;
}

export interface WebSearchConfigStore extends WebSearchConfigReader {
  set(key: string, value: unknown): Promise<void>;
  delete?(key: string): Promise<void>;
}

export interface WebSearchProviderConfig {
  provider: WebSearchProvider;
  apiKey?: string;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

export async function loadWebSearchProviderConfig(
  settings: WebSearchConfigReader,
): Promise<WebSearchProviderConfig> {
  const apiKeyRaw = await settings.get(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY);

  return {
    provider: "brave",
    apiKey: normalizeOptionalString(apiKeyRaw),
  };
}

export async function saveWebSearchApiKey(
  settings: WebSearchConfigStore,
  apiKey: string,
): Promise<void> {
  const normalized = apiKey.trim();
  if (normalized.length === 0) {
    throw new Error("API key cannot be empty.");
  }

  await settings.set(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY, normalized);
}

export async function clearWebSearchApiKey(settings: WebSearchConfigStore): Promise<void> {
  if (typeof settings.delete === "function") {
    await settings.delete(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY);
    return;
  }

  await settings.set(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY, "");
}

export function maskSecret(secret: string): string {
  const length = secret.length;
  if (length <= 4) {
    return "*".repeat(length);
  }

  if (length <= 8) {
    return `${secret.slice(0, 2)}${"*".repeat(length - 2)}`;
  }

  return `${secret.slice(0, 4)}${"*".repeat(length - 6)}${secret.slice(-2)}`;
}
