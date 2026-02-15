/**
 * Web-search configuration shared by tool + settings UI.
 */

export const WEB_SEARCH_PROVIDER_SETTING_KEY = "web.search.provider";
export const WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY = "web.search.brave.apiKey";
export const WEB_SEARCH_SERPER_API_KEY_SETTING_KEY = "web.search.serper.apiKey";
export const WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY = "web.search.tavily.apiKey";
export const WEB_SEARCH_JINA_API_KEY_SETTING_KEY = "web.search.jina.apiKey";

export const WEB_SEARCH_PROVIDERS = ["jina", "serper", "tavily", "brave"] as const;
export type WebSearchProvider = (typeof WEB_SEARCH_PROVIDERS)[number];

export const DEFAULT_WEB_SEARCH_PROVIDER: WebSearchProvider = "jina";

export interface WebSearchProviderInfo {
  id: WebSearchProvider;
  title: string;
  shortDescription: string;
  signupUrl: string;
  searchEndpoint: string;
  apiKeyLabel: string;
  apiKeyHelp: string;
  /** When true the provider works without an API key (key is optional for higher limits). */
  apiKeyOptional?: boolean;
}

export const WEB_SEARCH_PROVIDER_INFO: Record<WebSearchProvider, WebSearchProviderInfo> = {
  jina: {
    id: "jina",
    title: "Jina Search (default)",
    shortDescription: "Works out of the box — no signup or API key needed.",
    signupUrl: "https://jina.ai",
    searchEndpoint: "https://s.jina.ai/",
    apiKeyLabel: "Jina API key",
    apiKeyHelp: "Optional. Add a key for higher rate limits.",
    apiKeyOptional: true,
  },
  serper: {
    id: "serper",
    title: "Serper.dev",
    shortDescription: "Google SERP API, easy onboarding (free tier, no credit card).",
    signupUrl: "https://serper.dev",
    searchEndpoint: "https://google.serper.dev/search",
    apiKeyLabel: "Serper API key",
    apiKeyHelp: "Free tier available with email signup.",
  },
  tavily: {
    id: "tavily",
    title: "Tavily",
    shortDescription: "AI-native web search with relevance-ranked results.",
    signupUrl: "https://tavily.com",
    searchEndpoint: "https://api.tavily.com/search",
    apiKeyLabel: "Tavily API key",
    apiKeyHelp: "Free monthly credits, no credit card required.",
  },
  brave: {
    id: "brave",
    title: "Brave Search",
    shortDescription: "Direct Brave Search API support (existing users).",
    signupUrl: "https://api.search.brave.com",
    searchEndpoint: "https://api.search.brave.com/res/v1/web/search",
    apiKeyLabel: "Brave API key",
    apiKeyHelp: "Brave Search API subscription token.",
  },
};

export function getWebSearchEndpoint(provider: WebSearchProvider): string {
  return WEB_SEARCH_PROVIDER_INFO[provider].searchEndpoint;
}

export const WEB_SEARCH_PROVIDER_ENDPOINT_HOSTS: string[] = WEB_SEARCH_PROVIDERS.map((provider) => {
  const endpoint = getWebSearchEndpoint(provider);
  return new URL(endpoint).hostname.toLowerCase();
});

const WEB_SEARCH_API_KEY_BY_PROVIDER_SETTING_KEY: Record<WebSearchProvider, string> = {
  jina: WEB_SEARCH_JINA_API_KEY_SETTING_KEY,
  serper: WEB_SEARCH_SERPER_API_KEY_SETTING_KEY,
  tavily: WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY,
  brave: WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY,
};

export interface WebSearchConfigReader {
  get(key: string): Promise<unknown>;
}

export interface WebSearchConfigStore extends WebSearchConfigReader {
  set(key: string, value: unknown): Promise<void>;
  delete?(key: string): Promise<void>;
}

export interface WebSearchProviderConfig {
  provider: WebSearchProvider;
  apiKeys: Partial<Record<WebSearchProvider, string>>;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function parseProvider(value: unknown): WebSearchProvider | undefined {
  if (value === "jina" || value === "serper" || value === "tavily" || value === "brave") {
    return value;
  }
  return undefined;
}

export async function loadWebSearchProviderConfig(
  settings: WebSearchConfigReader,
): Promise<WebSearchProviderConfig> {
  const [providerRaw, jinaApiKeyRaw, serperApiKeyRaw, tavilyApiKeyRaw, braveApiKeyRaw] = await Promise.all([
    settings.get(WEB_SEARCH_PROVIDER_SETTING_KEY),
    settings.get(WEB_SEARCH_JINA_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_SERPER_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY),
  ]);

  const jinaApiKey = normalizeOptionalString(jinaApiKeyRaw);
  const serperApiKey = normalizeOptionalString(serperApiKeyRaw);
  const tavilyApiKey = normalizeOptionalString(tavilyApiKeyRaw);
  const braveApiKey = normalizeOptionalString(braveApiKeyRaw);

  // Prefer an explicitly saved provider. Otherwise, if any key-required provider
  // has a key, infer that provider so existing users aren't silently switched to
  // the zero-config default after an upgrade.
  const provider = parseProvider(providerRaw)
    ?? (serperApiKey ? "serper" : braveApiKey ? "brave" : tavilyApiKey ? "tavily" : DEFAULT_WEB_SEARCH_PROVIDER);

  return {
    provider,
    apiKeys: {
      jina: jinaApiKey,
      serper: serperApiKey,
      tavily: tavilyApiKey,
      brave: braveApiKey,
    },
  };
}

export async function saveWebSearchProvider(
  settings: WebSearchConfigStore,
  provider: WebSearchProvider,
): Promise<void> {
  await settings.set(WEB_SEARCH_PROVIDER_SETTING_KEY, provider);
}

export async function saveWebSearchApiKey(
  settings: WebSearchConfigStore,
  provider: WebSearchProvider,
  apiKey: string,
): Promise<void> {
  const normalized = apiKey.trim();
  if (normalized.length === 0) {
    throw new Error("API key cannot be empty.");
  }

  await settings.set(WEB_SEARCH_API_KEY_BY_PROVIDER_SETTING_KEY[provider], normalized);
}

export async function clearWebSearchApiKey(
  settings: WebSearchConfigStore,
  provider: WebSearchProvider,
): Promise<void> {
  const key = WEB_SEARCH_API_KEY_BY_PROVIDER_SETTING_KEY[provider];
  if (typeof settings.delete === "function") {
    await settings.delete(key);
    return;
  }

  await settings.set(key, "");
}

export function getApiKeyForProvider(
  config: WebSearchProviderConfig,
  provider: WebSearchProvider = config.provider,
): string | undefined {
  return normalizeOptionalString(config.apiKeys[provider]);
}

/** Returns true when the provider cannot work without an API key. */
export function isApiKeyRequired(provider: WebSearchProvider): boolean {
  return WEB_SEARCH_PROVIDER_INFO[provider].apiKeyOptional !== true;
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
