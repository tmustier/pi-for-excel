/**
 * web_search — external web search (Jina/Serper/Tavily/Brave).
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { integrationsCommandHint } from "../integrations/naming.js";
import { getErrorMessage } from "../utils/errors.js";
import {
  getHttpErrorReason,
  runWithTimeoutAbort,
} from "../utils/network.js";
import { isRecord } from "../utils/type-guards.js";
import {
  buildProxyDownErrorMessage,
  getEnabledProxyBaseUrl,
  isLikelyProxyConnectionError,
  resolveOutboundRequestUrl,
  type ProxyAwareSettingsStore,
} from "./external-fetch.js";
import {
  getApiKeyForProvider,
  getWebSearchEndpoint,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  type WebSearchProvider,
  type WebSearchProviderInfo,
  WEB_SEARCH_PROVIDER_INFO,
} from "./web-search-config.js";

const WEB_SEARCH_TIMEOUT_MS = 12_000;

const RECENCY_VALUES = ["day", "week", "month", "year"] as const;
type RecencyValue = (typeof RECENCY_VALUES)[number];

const braveRecencyToFreshness: Record<RecencyValue, string> = {
  day: "pd",
  week: "pw",
  month: "pm",
  year: "py",
};

const serperRecencyToTbs: Record<RecencyValue, string> = {
  day: "qdr:d",
  week: "qdr:w",
  month: "qdr:m",
  year: "qdr:y",
};

function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(values.map((value) => Type.Literal(value)), opts);
}

const schema = Type.Object({
  query: Type.String({
    minLength: 1,
    description: "Search query.",
  }),
  recency: Type.Optional(StringEnum([...RECENCY_VALUES], {
    description: "Optional recency filter: day/week/month/year.",
  })),
  site: Type.Optional(Type.Union([
    Type.String({ description: "Optional site/domain filter (example: docs.github.com)." }),
    Type.Array(Type.String(), {
      minItems: 1,
      maxItems: 6,
      description: "Optional list of site/domain filters.",
    }),
  ])),
  max_results: Type.Optional(Type.Integer({
    minimum: 1,
    maximum: 10,
    description: "Maximum results to return (1-10). Default: 5.",
  })),
});

type Params = Static<typeof schema>;

export interface WebSearchHit {
  title: string;
  url: string;
  snippet: string;
}

export interface WebSearchFallbackInfo {
  fromProvider: WebSearchProvider;
  toProvider: "jina";
  reason: string;
}

export interface WebSearchToolDetails {
  kind: "web_search";
  ok: boolean;
  provider: WebSearchProvider;
  query: string;
  sentQuery: string;
  recency?: RecencyValue;
  siteFilters?: string[];
  maxResults: number;
  resultCount?: number;
  proxied?: boolean;
  proxyBaseUrl?: string;
  fallback?: WebSearchFallbackInfo;
  error?: string;
  /** `true` when the failure is due to the local CORS proxy being unreachable. */
  proxyDown?: boolean;
}

export interface WebSearchToolConfig {
  provider: WebSearchProvider;
  apiKey?: string;
  jinaApiKey?: string;
  proxyBaseUrl?: string;
}

export interface WebSearchExecuteConfig {
  provider: WebSearchProvider;
  apiKey: string;
  proxyBaseUrl?: string;
}

export interface WebSearchExecutionResult {
  hits: WebSearchHit[];
  sentQuery: string;
  proxied: boolean;
  proxyBaseUrl?: string;
}

export interface WebSearchToolDependencies {
  getConfig?: () => Promise<WebSearchToolConfig>;
  executeSearch?: (
    params: Params,
    config: WebSearchExecuteConfig,
    signal: AbortSignal | undefined,
  ) => Promise<WebSearchExecutionResult>;
}

export interface WebSearchApiKeyValidationResult {
  ok: boolean;
  provider: WebSearchProvider;
  message: string;
  proxied?: boolean;
  proxyBaseUrl?: string;
  resultCount?: number;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function parseSites(value: unknown): string[] {
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? [trimmed] : [];
  }

  if (!Array.isArray(value)) return [];

  const out: string[] = [];
  for (const item of value) {
    if (typeof item !== "string") continue;
    const trimmed = item.trim();
    if (trimmed.length === 0) continue;
    out.push(trimmed);
  }

  return out;
}

function isRecencyValue(value: string): value is RecencyValue {
  return value === "day" || value === "week" || value === "month" || value === "year";
}

function parseParams(raw: unknown): Params {
  if (!isRecord(raw)) {
    throw new Error("Invalid web_search params: expected an object.");
  }

  const query = normalizeOptionalString(raw.query);
  if (!query) {
    throw new Error("web_search requires a non-empty query.");
  }

  const recency = normalizeOptionalString(raw.recency);
  const recencyValue = recency && isRecencyValue(recency)
    ? recency
    : undefined;

  const sites = parseSites(raw.site);

  const maxResultsRaw = raw.max_results;
  let maxResults = 5;
  if (typeof maxResultsRaw === "number" && Number.isInteger(maxResultsRaw)) {
    if (maxResultsRaw < 1 || maxResultsRaw > 10) {
      throw new Error("max_results must be between 1 and 10.");
    }
    maxResults = maxResultsRaw;
  }

  const params: Params = {
    query,
    max_results: maxResults,
  };

  if (recencyValue) {
    params.recency = recencyValue;
  }

  if (sites.length > 0) {
    params.site = sites;
  }

  return params;
}

function buildSiteQuery(sites: string[]): string {
  if (sites.length === 0) return "";
  return sites.map((site) => `site:${site}`).join(" OR ");
}

function buildSentQuery(params: Params): string {
  const sites = parseSites(params.site);
  const siteQuery = buildSiteQuery(sites);
  if (!siteQuery) return params.query;
  return `${params.query} (${siteQuery})`;
}

function providerInfo(provider: WebSearchProvider): WebSearchProviderInfo {
  return WEB_SEARCH_PROVIDER_INFO[provider];
}

type WebSearchFailureKind = "missing_api_key" | "http" | "network" | "timeout";

class WebSearchExecutionError extends Error {
  readonly provider: WebSearchProvider;
  readonly kind: WebSearchFailureKind;
  readonly statusCode?: number;

  constructor(args: {
    provider: WebSearchProvider;
    kind: WebSearchFailureKind;
    message: string;
    statusCode?: number;
  }) {
    super(args.message);
    this.name = "WebSearchExecutionError";
    this.provider = args.provider;
    this.kind = args.kind;
    this.statusCode = args.statusCode;
  }
}

interface ProviderRequest {
  requestInit: {
    method: "GET" | "POST";
    headers: Record<string, string>;
    body?: string;
  };
  targetUrl: string;
  sentQuery: string;
}

function buildProviderRequest(
  params: Params,
  provider: WebSearchProvider,
  apiKey: string,
): ProviderRequest {
  const sentQuery = buildSentQuery(params);
  const maxResults = params.max_results ?? 5;

  if (provider === "jina") {
    const targetUrl = `${getWebSearchEndpoint(provider)}${encodeURIComponent(sentQuery)}`;
    const headers: Record<string, string> = {
      Accept: "application/json",
      "X-Retain-Images": "none",
    };

    if (apiKey) {
      headers.Authorization = `Bearer ${apiKey}`;
    }

    return {
      targetUrl,
      sentQuery,
      requestInit: {
        method: "GET",
        headers,
      },
    };
  }

  if (provider === "brave") {
    const url = new URL(getWebSearchEndpoint(provider));
    url.searchParams.set("q", sentQuery);
    url.searchParams.set("count", String(maxResults));

    if (params.recency) {
      url.searchParams.set("freshness", braveRecencyToFreshness[params.recency]);
    }

    return {
      targetUrl: url.toString(),
      sentQuery,
      requestInit: {
        method: "GET",
        headers: {
          Accept: "application/json",
          "X-Subscription-Token": apiKey,
        },
      },
    };
  }

  if (provider === "serper") {
    const body: Record<string, unknown> = {
      q: sentQuery,
      num: maxResults,
    };

    if (params.recency) {
      body.tbs = serperRecencyToTbs[params.recency];
    }

    return {
      targetUrl: getWebSearchEndpoint(provider),
      sentQuery,
      requestInit: {
        method: "POST",
        headers: {
          Accept: "application/json",
          "Content-Type": "application/json",
          "X-API-KEY": apiKey,
        },
        body: JSON.stringify(body),
      },
    };
  }

  const tavilyBody: Record<string, unknown> = {
    api_key: apiKey,
    query: sentQuery,
    max_results: maxResults,
    search_depth: "basic",
    include_answer: false,
    include_raw_content: false,
  };

  return {
    targetUrl: getWebSearchEndpoint(provider),
    sentQuery,
    requestInit: {
      method: "POST",
      headers: {
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify(tavilyBody),
    },
  };
}

interface HitParsingShape {
  titleKey: string;
  urlKey: string;
  snippetKeys: readonly string[];
}

function readArrayPath(payload: unknown, path: readonly string[]): unknown[] {
  let cursor: unknown = payload;

  for (const key of path) {
    if (!isRecord(cursor)) {
      return [];
    }

    cursor = cursor[key];
  }

  return Array.isArray(cursor) ? cursor : [];
}

function parseHitsFromEntries(entries: readonly unknown[], shape: HitParsingShape): WebSearchHit[] {
  const hits: WebSearchHit[] = [];

  for (const entry of entries) {
    if (!isRecord(entry)) continue;

    const title = normalizeOptionalString(entry[shape.titleKey]);
    const url = normalizeOptionalString(entry[shape.urlKey]);

    if (!title || !url) continue;

    let snippet = "";
    for (const snippetKey of shape.snippetKeys) {
      const candidate = normalizeOptionalString(entry[snippetKey]);
      if (!candidate) continue;
      snippet = candidate;
      break;
    }

    hits.push({
      title,
      url,
      snippet,
    });
  }

  return hits;
}

function parseBraveHits(payload: unknown): WebSearchHit[] {
  return parseHitsFromEntries(
    readArrayPath(payload, ["web", "results"]),
    {
      titleKey: "title",
      urlKey: "url",
      snippetKeys: ["description", "snippet"],
    },
  );
}

function parseSerperHits(payload: unknown): WebSearchHit[] {
  return parseHitsFromEntries(
    readArrayPath(payload, ["organic"]),
    {
      titleKey: "title",
      urlKey: "link",
      snippetKeys: ["snippet"],
    },
  );
}

function parseTavilyHits(payload: unknown): WebSearchHit[] {
  return parseHitsFromEntries(
    readArrayPath(payload, ["results"]),
    {
      titleKey: "title",
      urlKey: "url",
      snippetKeys: ["content"],
    },
  );
}

function parseJinaHits(payload: unknown): WebSearchHit[] {
  return parseHitsFromEntries(
    readArrayPath(payload, ["data"]),
    {
      titleKey: "title",
      urlKey: "url",
      snippetKeys: ["description", "content"],
    },
  );
}

const SEARCH_HIT_PARSERS: Record<WebSearchProvider, (payload: unknown) => WebSearchHit[]> = {
  jina: parseJinaHits,
  brave: parseBraveHits,
  serper: parseSerperHits,
  tavily: parseTavilyHits,
};

function parseSearchHits(provider: WebSearchProvider, payload: unknown, maxResults?: number): WebSearchHit[] {
  const hits = SEARCH_HIT_PARSERS[provider](payload);
  if (typeof maxResults === "number" && hits.length > maxResults) {
    return hits.slice(0, maxResults);
  }
  return hits;
}

function shouldFallbackToJina(provider: WebSearchProvider, error: unknown): boolean {
  if (provider === "jina") return false;

  if (error instanceof WebSearchExecutionError) {
    if (error.kind === "missing_api_key") return true;
    if (error.kind === "network" || error.kind === "timeout") return true;

    if (error.kind === "http") {
      const status = error.statusCode;
      if (typeof status !== "number") return false;
      return status === 401 || status === 403 || status === 429 || status >= 500;
    }

    return false;
  }

  const message = getErrorMessage(error).toLowerCase();

  if (/\b(401|403|429)\b/.test(message)) {
    return true;
  }

  if (/\b5\d{2}\b/.test(message)) {
    return true;
  }

  if (
    message.includes("timed out")
    || message.includes("timeout")
    || message.includes("networkerror")
    || message.includes("fetch failed")
    || message.includes("econnrefused")
    || message.includes("econnreset")
    || message.includes("enotfound")
  ) {
    return true;
  }

  return false;
}

function summarizeFallbackReason(message: string): string {
  const compact = message.replaceAll(/\s+/g, " ").trim();
  if (compact.length <= 180) return compact;
  return `${compact.slice(0, 177)}...`;
}

function buildFallbackWarning(fallback: WebSearchFallbackInfo): string {
  const fromInfo = providerInfo(fallback.fromProvider);
  const toInfo = providerInfo(fallback.toProvider);

  return `⚠️ ${fromInfo.title} search failed (${fallback.reason}) — used ${toInfo.title} for this request. Check your ${fromInfo.apiKeyLabel} in ${integrationsCommandHint()}.`;
}

function buildResultMarkdown(args: {
  provider: WebSearchProvider;
  params: Params;
  sentQuery: string;
  hits: WebSearchHit[];
  proxied: boolean;
  proxyBaseUrl?: string;
  fallback?: WebSearchFallbackInfo;
}): string {
  const { provider, params, sentQuery, hits, proxied, proxyBaseUrl, fallback } = args;

  const providerTitle = providerInfo(provider).title;

  const lines: string[] = [];

  if (fallback) {
    lines.push(buildFallbackWarning(fallback));
    lines.push("");
  }

  lines.push(`Web search via ${providerTitle}`);
  lines.push("");
  lines.push("Sent:");
  lines.push(`- query: \`${sentQuery}\``);

  if (params.recency) {
    lines.push(`- recency: ${params.recency}`);
  }

  const sites = parseSites(params.site);
  if (sites.length > 0) {
    lines.push(`- sites: ${sites.join(", ")}`);
  }

  lines.push(`- max results requested: ${params.max_results ?? 5}`);
  lines.push(`- transport: ${proxied ? `proxy (${proxyBaseUrl ?? "configured proxy"})` : "direct"}`);
  lines.push("");

  if (hits.length === 0) {
    lines.push("No results found.");
    return lines.join("\n");
  }

  lines.push("Results:");
  for (let i = 0; i < hits.length; i += 1) {
    const hit = hits[i];
    const index = i + 1;
    lines.push(`[${index}] [${hit.title}](${hit.url})`);
    if (hit.snippet.trim().length > 0) {
      lines.push(`    ${hit.snippet}`);
    }
  }

  return lines.join("\n");
}

async function defaultGetConfig(): Promise<WebSearchToolConfig> {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  const settings: ProxyAwareSettingsStore = storageModule.getAppStorage().settings;

  const [providerConfig, proxyBaseUrl] = await Promise.all([
    loadWebSearchProviderConfig(settings),
    getEnabledProxyBaseUrl(settings),
  ]);

  return {
    provider: providerConfig.provider,
    apiKey: getApiKeyForProvider(providerConfig),
    jinaApiKey: getApiKeyForProvider(providerConfig, "jina"),
    proxyBaseUrl,
  };
}

async function defaultExecuteSearch(
  params: Params,
  config: WebSearchExecuteConfig,
  signal: AbortSignal | undefined,
): Promise<WebSearchExecutionResult> {
  const request = buildProviderRequest(params, config.provider, config.apiKey);
  const resolved = resolveOutboundRequestUrl({
    targetUrl: request.targetUrl,
    proxyBaseUrl: config.proxyBaseUrl,
  });

  try {
    return await runWithTimeoutAbort({
      signal,
      timeoutMs: WEB_SEARCH_TIMEOUT_MS,
      timeoutErrorMessage: `web_search timed out after ${WEB_SEARCH_TIMEOUT_MS}ms.`,
      run: async (requestSignal) => {
        const response = await fetch(resolved.requestUrl, {
          method: request.requestInit.method,
          headers: request.requestInit.headers,
          body: request.requestInit.body,
          signal: requestSignal,
        });

        const text = await response.text();

        if (!response.ok) {
          const reason = getHttpErrorReason(response.status, text);
          throw new WebSearchExecutionError({
            provider: config.provider,
            kind: "http",
            statusCode: response.status,
            message: `${providerInfo(config.provider).title} search request failed (${response.status}): ${reason}`,
          });
        }

        let payload: unknown = null;
        try {
          payload = JSON.parse(text);
        } catch {
          payload = null;
        }

        const hits = parseSearchHits(config.provider, payload, params.max_results);

        return {
          hits,
          sentQuery: request.sentQuery,
          proxied: resolved.proxied,
          proxyBaseUrl: resolved.proxyBaseUrl,
        };
      },
    });
  } catch (error: unknown) {
    if (error instanceof WebSearchExecutionError) {
      throw error;
    }

    if (error instanceof Error && error.message === "Aborted") {
      throw error;
    }

    const message = getErrorMessage(error);
    const kind: WebSearchFailureKind = message.includes("timed out after") ? "timeout" : "network";

    throw new WebSearchExecutionError({
      provider: config.provider,
      kind,
      message,
    });
  }
}

export async function validateWebSearchApiKey(args: {
  provider: WebSearchProvider;
  apiKey: string;
  proxyBaseUrl?: string;
  signal?: AbortSignal;
}): Promise<WebSearchApiKeyValidationResult> {
  const provider = args.provider;
  const normalizedApiKey = args.apiKey.trim();
  if (normalizedApiKey.length === 0) {
    return {
      ok: false,
      provider,
      message: "API key cannot be empty.",
    };
  }

  try {
    const result = await defaultExecuteSearch(
      {
        query: "Excel formulas",
        max_results: 1,
      },
      {
        provider,
        apiKey: normalizedApiKey,
        proxyBaseUrl: args.proxyBaseUrl,
      },
      args.signal,
    );

    const transport = result.proxied ? `proxy (${result.proxyBaseUrl ?? "configured"})` : "direct";

    return {
      ok: true,
      provider,
      message: `${providerInfo(provider).title} key is valid (${result.hits.length} result${result.hits.length === 1 ? "" : "s"}, ${transport}).`,
      proxied: result.proxied,
      proxyBaseUrl: result.proxyBaseUrl,
      resultCount: result.hits.length,
    };
  } catch (error: unknown) {
    return {
      ok: false,
      provider,
      message: getErrorMessage(error),
    };
  }
}

export function createWebSearchTool(
  dependencies: WebSearchToolDependencies = {},
): AgentTool<TSchema, WebSearchToolDetails> {
  const getConfig = dependencies.getConfig ?? defaultGetConfig;
  const executeSearch = dependencies.executeSearch ?? defaultExecuteSearch;

  return {
    name: "web_search",
    label: "Web Search",
    description:
      "Search the public web. Returns compact, cited links with snippets. Works out of the box with Jina (default); optionally Serper, Tavily, or Brave.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<WebSearchToolDetails>> => {
      let params: Params | null = null;
      let configuredProvider: WebSearchProvider = "jina";
      let usedProxyBaseUrl: string | undefined;

      try {
        const parsedParams = parseParams(rawParams);
        params = parsedParams;

        const config = await getConfig();
        configuredProvider = config.provider;
        usedProxyBaseUrl = config.proxyBaseUrl;

        const configuredApiKey = normalizeOptionalString(config.apiKey) ?? "";
        const fallbackJinaApiKey = normalizeOptionalString(config.jinaApiKey) ?? "";

        const runSearch = (
          provider: WebSearchProvider,
          apiKey: string,
        ): Promise<WebSearchExecutionResult> => executeSearch(
          parsedParams,
          {
            provider,
            apiKey,
            proxyBaseUrl: config.proxyBaseUrl,
          },
          signal,
        );

        let effectiveProvider = configuredProvider;
        let fallback: WebSearchFallbackInfo | undefined;

        let result: WebSearchExecutionResult;
        try {
          if (!configuredApiKey && isApiKeyRequired(configuredProvider)) {
            throw new WebSearchExecutionError({
              provider: configuredProvider,
              kind: "missing_api_key",
              message: `Web search API key is missing. Open ${integrationsCommandHint()} and set the ${providerInfo(configuredProvider).apiKeyLabel}.`,
            });
          }

          result = await runSearch(configuredProvider, configuredApiKey);
        } catch (error: unknown) {
          if (!shouldFallbackToJina(configuredProvider, error)) {
            throw error;
          }

          fallback = {
            fromProvider: configuredProvider,
            toProvider: "jina",
            reason: summarizeFallbackReason(getErrorMessage(error)),
          };

          try {
            result = await runSearch("jina", fallbackJinaApiKey);
            effectiveProvider = "jina";
          } catch (fallbackError: unknown) {
            const primaryMessage = getErrorMessage(error);
            const fallbackMessage = getErrorMessage(fallbackError);
            throw new Error(`${primaryMessage}; fallback to Jina Search also failed: ${fallbackMessage}`);
          }
        }

        return {
          content: [{ type: "text", text: buildResultMarkdown({
            provider: effectiveProvider,
            params: parsedParams,
            sentQuery: result.sentQuery,
            hits: result.hits,
            proxied: result.proxied,
            proxyBaseUrl: result.proxyBaseUrl,
            fallback,
          }) }],
          details: {
            kind: "web_search",
            ok: true,
            provider: effectiveProvider,
            query: parsedParams.query,
            sentQuery: result.sentQuery,
            recency: parsedParams.recency,
            siteFilters: parseSites(parsedParams.site),
            maxResults: parsedParams.max_results ?? 5,
            resultCount: result.hits.length,
            proxied: result.proxied,
            proxyBaseUrl: result.proxyBaseUrl,
            fallback,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        const proxyDown = isLikelyProxyConnectionError(message, usedProxyBaseUrl);
        const displayMessage = proxyDown
          ? buildProxyDownErrorMessage("Web search", message)
          : `Error: ${message}`;
        const fallbackQuery = params?.query
          ?? (isRecord(rawParams) && typeof rawParams.query === "string" ? rawParams.query : "");

        return {
          content: [{ type: "text", text: displayMessage }],
          details: {
            kind: "web_search",
            ok: false,
            provider: configuredProvider,
            query: fallbackQuery,
            sentQuery: fallbackQuery,
            maxResults: params?.max_results ?? 5,
            error: message,
            proxyDown,
          },
        };
      }
    },
  };
}
