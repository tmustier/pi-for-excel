/**
 * Web-search configuration shared by tool + settings UI.
 */

const CONNECTION_STORE_KEY = "connections.store.v1";
const CONNECTION_STORE_VERSION = 1;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export const WEB_SEARCH_PROVIDER_SETTING_KEY = "web.search.provider";
export const WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY = "web.search.brave.apiKey";
export const WEB_SEARCH_SERPER_API_KEY_SETTING_KEY = "web.search.serper.apiKey";
export const WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY = "web.search.tavily.apiKey";
export const WEB_SEARCH_JINA_API_KEY_SETTING_KEY = "web.search.jina.apiKey";
export const WEB_SEARCH_FIRECRAWL_API_KEY_SETTING_KEY = "web.search.firecrawl.apiKey";

export const WEB_SEARCH_PROVIDERS = ["jina", "firecrawl", "serper", "tavily", "brave"] as const;
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
    shortDescription: "Fast web search API with free tier.",
    signupUrl: "https://jina.ai",
    searchEndpoint: "https://s.jina.ai/",
    apiKeyLabel: "Jina API key",
    apiKeyHelp: "Free tier available at jina.ai.",
  },
  firecrawl: {
    id: "firecrawl",
    title: "Firecrawl",
    shortDescription: "Web search with optional page scraping. 500 free credits.",
    signupUrl: "https://firecrawl.dev",
    searchEndpoint: "https://api.firecrawl.dev/v2/search",
    apiKeyLabel: "Firecrawl API key",
    apiKeyHelp: "Free tier with 500 credits, no credit card required.",
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
  firecrawl: WEB_SEARCH_FIRECRAWL_API_KEY_SETTING_KEY,
  serper: WEB_SEARCH_SERPER_API_KEY_SETTING_KEY,
  tavily: WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY,
  brave: WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY,
};

export const WEB_SEARCH_CONNECTION_ID = "builtin.web_search.providers";

const WEB_SEARCH_CONNECTION_SECRET_FIELD_BY_PROVIDER: Record<WebSearchProvider, string> = {
  jina: "jina_api_key",
  firecrawl: "firecrawl_api_key",
  serper: "serper_api_key",
  tavily: "tavily_api_key",
  brave: "brave_api_key",
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
  if (value === "jina" || value === "firecrawl" || value === "serper" || value === "tavily" || value === "brave") {
    return value;
  }
  return undefined;
}

function createEmptyApiKeyMap(): Partial<Record<WebSearchProvider, string>> {
  return {
    jina: undefined,
    firecrawl: undefined,
    serper: undefined,
    tavily: undefined,
    brave: undefined,
  };
}

async function loadLegacyWebSearchApiKeys(
  settings: WebSearchConfigReader,
): Promise<Partial<Record<WebSearchProvider, string>>> {
  const [jinaApiKeyRaw, firecrawlApiKeyRaw, serperApiKeyRaw, tavilyApiKeyRaw, braveApiKeyRaw] = await Promise.all([
    settings.get(WEB_SEARCH_JINA_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_FIRECRAWL_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_SERPER_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_TAVILY_API_KEY_SETTING_KEY),
    settings.get(WEB_SEARCH_BRAVE_API_KEY_SETTING_KEY),
  ]);

  return {
    jina: normalizeOptionalString(jinaApiKeyRaw),
    firecrawl: normalizeOptionalString(firecrawlApiKeyRaw),
    serper: normalizeOptionalString(serperApiKeyRaw),
    tavily: normalizeOptionalString(tavilyApiKeyRaw),
    brave: normalizeOptionalString(braveApiKeyRaw),
  };
}

async function loadConnectionStoreWebSearchApiKeys(
  settings: WebSearchConfigReader,
): Promise<Partial<Record<WebSearchProvider, string>>> {
  const rawStore = await settings.get(CONNECTION_STORE_KEY);
  if (!isRecord(rawStore)) return createEmptyApiKeyMap();

  const rawItems = rawStore.items;
  if (!isRecord(rawItems)) return createEmptyApiKeyMap();

  const rawRecord = rawItems[WEB_SEARCH_CONNECTION_ID];
  if (!isRecord(rawRecord)) return createEmptyApiKeyMap();

  const rawSecrets = rawRecord.secrets;
  if (!isRecord(rawSecrets)) return createEmptyApiKeyMap();

  const apiKeys = createEmptyApiKeyMap();

  for (const provider of WEB_SEARCH_PROVIDERS) {
    const fieldId = WEB_SEARCH_CONNECTION_SECRET_FIELD_BY_PROVIDER[provider];
    apiKeys[provider] = normalizeOptionalString(rawSecrets[fieldId]);
  }

  return apiKeys;
}

function mergeApiKeys(args: {
  primary: Partial<Record<WebSearchProvider, string>>;
  fallback: Partial<Record<WebSearchProvider, string>>;
}): Partial<Record<WebSearchProvider, string>> {
  const merged = createEmptyApiKeyMap();

  for (const provider of WEB_SEARCH_PROVIDERS) {
    merged[provider] = normalizeOptionalString(args.primary[provider])
      ?? normalizeOptionalString(args.fallback[provider]);
  }

  return merged;
}

type StoredConnectionRecord = {
  status?: "connected" | "missing" | "invalid" | "error";
  lastValidatedAt?: string;
  lastError?: string;
  secrets?: Record<string, string>;
};

async function loadConnectionStoreItems(
  settings: WebSearchConfigStore,
): Promise<Record<string, StoredConnectionRecord>> {
  const raw = await settings.get(CONNECTION_STORE_KEY);
  if (!isRecord(raw)) return {};

  const rawItems = raw.items;
  if (!isRecord(rawItems)) return {};

  const items: Record<string, StoredConnectionRecord> = {};

  for (const [connectionId, rawRecord] of Object.entries(rawItems)) {
    if (!isRecord(rawRecord)) continue;

    const rawSecrets = rawRecord.secrets;
    const secrets: Record<string, string> = {};
    if (isRecord(rawSecrets)) {
      for (const [fieldId, value] of Object.entries(rawSecrets)) {
        const normalized = normalizeOptionalString(value);
        if (!normalized) continue;
        secrets[fieldId] = normalized;
      }
    }

    const status = rawRecord.status;
    items[connectionId] = {
      status: status === "connected" || status === "missing" || status === "invalid" || status === "error"
        ? status
        : undefined,
      lastValidatedAt: normalizeOptionalString(rawRecord.lastValidatedAt),
      lastError: normalizeOptionalString(rawRecord.lastError),
      secrets,
    };
  }

  return items;
}

async function saveConnectionStoreItems(
  settings: WebSearchConfigStore,
  items: Record<string, StoredConnectionRecord>,
): Promise<void> {
  await settings.set(CONNECTION_STORE_KEY, {
    version: CONNECTION_STORE_VERSION,
    items,
  });
}

async function writeConnectionStoreWebSearchApiKeys(
  settings: WebSearchConfigStore,
  apiKeys: Partial<Record<WebSearchProvider, string>>,
): Promise<void> {
  const items = await loadConnectionStoreItems(settings);
  const previous = items[WEB_SEARCH_CONNECTION_ID];

  const secrets: Record<string, string> = {};
  for (const provider of WEB_SEARCH_PROVIDERS) {
    const normalized = normalizeOptionalString(apiKeys[provider]);
    if (!normalized) continue;
    const fieldId = WEB_SEARCH_CONNECTION_SECRET_FIELD_BY_PROVIDER[provider];
    secrets[fieldId] = normalized;
  }

  if (Object.keys(secrets).length === 0) {
    if (WEB_SEARCH_CONNECTION_ID in items) {
      delete items[WEB_SEARCH_CONNECTION_ID];
      await saveConnectionStoreItems(settings, items);
    }
    return;
  }

  items[WEB_SEARCH_CONNECTION_ID] = {
    status: "connected",
    lastValidatedAt: previous?.lastValidatedAt,
    lastError: undefined,
    secrets,
  };

  await saveConnectionStoreItems(settings, items);
}

async function clearLegacyWebSearchApiKey(
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

export async function migrateLegacyWebSearchApiKeysToConnectionStore(
  settings: WebSearchConfigStore,
): Promise<boolean> {
  const [legacyApiKeys, connectionApiKeys] = await Promise.all([
    loadLegacyWebSearchApiKeys(settings),
    loadConnectionStoreWebSearchApiKeys(settings),
  ]);

  const mergedApiKeys = mergeApiKeys({
    primary: connectionApiKeys,
    fallback: legacyApiKeys,
  });

  const shouldWriteConnectionStore = JSON.stringify(mergedApiKeys) !== JSON.stringify(connectionApiKeys);

  const providersWithLegacyKeys = WEB_SEARCH_PROVIDERS.filter((provider) =>
    normalizeOptionalString(legacyApiKeys[provider]) !== undefined,
  );

  if (!shouldWriteConnectionStore && providersWithLegacyKeys.length === 0) {
    return false;
  }

  if (shouldWriteConnectionStore) {
    await writeConnectionStoreWebSearchApiKeys(settings, mergedApiKeys);
  }

  for (const provider of providersWithLegacyKeys) {
    await clearLegacyWebSearchApiKey(settings, provider);
  }

  return shouldWriteConnectionStore || providersWithLegacyKeys.length > 0;
}

export async function loadWebSearchProviderConfig(
  settings: WebSearchConfigReader,
): Promise<WebSearchProviderConfig> {
  const [providerRaw, connectionApiKeys, legacyApiKeys] = await Promise.all([
    settings.get(WEB_SEARCH_PROVIDER_SETTING_KEY),
    loadConnectionStoreWebSearchApiKeys(settings),
    loadLegacyWebSearchApiKeys(settings),
  ]);

  const apiKeys = mergeApiKeys({
    primary: connectionApiKeys,
    fallback: legacyApiKeys,
  });

  // Prefer an explicitly saved provider. Otherwise, if any key-required provider
  // has a key, infer that provider so existing users aren't silently switched to
  // the zero-config default after an upgrade.
  const provider = parseProvider(providerRaw)
    ?? (apiKeys.firecrawl ? "firecrawl" : apiKeys.serper ? "serper" : apiKeys.brave ? "brave" : apiKeys.tavily ? "tavily" : DEFAULT_WEB_SEARCH_PROVIDER);

  return {
    provider,
    apiKeys,
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

  const current = await loadWebSearchProviderConfig(settings);
  const nextApiKeys = {
    ...current.apiKeys,
    [provider]: normalized,
  };

  await writeConnectionStoreWebSearchApiKeys(settings, nextApiKeys);
  await clearLegacyWebSearchApiKey(settings, provider);
}

export async function clearWebSearchApiKey(
  settings: WebSearchConfigStore,
  provider: WebSearchProvider,
): Promise<void> {
  const current = await loadWebSearchProviderConfig(settings);
  const nextApiKeys = {
    ...current.apiKeys,
    [provider]: undefined,
  };

  await writeConnectionStoreWebSearchApiKeys(settings, nextApiKeys);
  await clearLegacyWebSearchApiKey(settings, provider);
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

function hasRepeatedLongSegment(value: string): boolean {
  const minimumSegmentLength = 12;
  const maxSegmentLength = Math.floor(value.length / 2);
  if (maxSegmentLength < minimumSegmentLength) return false;

  for (let segmentLength = maxSegmentLength; segmentLength >= minimumSegmentLength; segmentLength -= 1) {
    for (let start = 0; start + (segmentLength * 2) <= value.length; start += 1) {
      const segment = value.slice(start, start + segmentLength);
      const nextSegment = value.slice(start + segmentLength, start + (segmentLength * 2));
      if (segment === nextSegment) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Quick format check for an API key before saving. Returns `null` when the key
 * looks plausible, or a human-readable warning string when it looks wrong.
 *
 * This is intentionally loose — it catches obvious mistakes (empty, whitespace,
 * too short, wrong prefix) without risking false rejections if a provider
 * tweaks their format.
 */
export function checkApiKeyFormat(provider: WebSearchProvider, apiKey: string): string | null {
  const key = apiKey.trim();

  if (key.length === 0) return "API key is empty.";

  if (/\s/.test(key)) {
    return "API key contains spaces or newlines — check for copy-paste errors.";
  }

  if (key.length < 10) {
    return "API key looks too short — check for truncation.";
  }

  if (hasRepeatedLongSegment(key)) {
    return "API key contains a repeated long segment — check for accidental double paste.";
  }

  if (provider === "jina" && !key.startsWith("jina_")) {
    return "Jina keys usually start with \"jina_\" — double-check the value.";
  }

  if (provider === "firecrawl" && !key.startsWith("fc-")) {
    return "Firecrawl keys usually start with \"fc-\" — double-check the value.";
  }

  if (provider === "tavily" && !key.startsWith("tvly-")) {
    return "Tavily keys usually start with \"tvly-\" — double-check the value.";
  }

  return null;
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
