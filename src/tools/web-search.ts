/**
 * web_search — external web search (Brave Search API).
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { getErrorMessage } from "../utils/errors.js";
import { isRecord } from "../utils/type-guards.js";
import {
  getEnabledProxyBaseUrl,
  resolveOutboundRequestUrl,
  type ProxyAwareSettingsStore,
} from "./external-fetch.js";
import { loadWebSearchProviderConfig } from "./web-search-config.js";

const BRAVE_WEB_SEARCH_ENDPOINT = "https://api.search.brave.com/res/v1/web/search";
const WEB_SEARCH_TIMEOUT_MS = 12_000;

const RECENCY_VALUES = ["day", "week", "month", "year"] as const;
type RecencyValue = (typeof RECENCY_VALUES)[number];

const recencyToFreshness: Record<RecencyValue, string> = {
  day: "pd",
  week: "pw",
  month: "pm",
  year: "py",
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

export interface WebSearchToolDetails {
  kind: "web_search";
  ok: boolean;
  provider: "brave";
  query: string;
  sentQuery: string;
  recency?: RecencyValue;
  siteFilters?: string[];
  maxResults: number;
  resultCount?: number;
  proxied?: boolean;
  proxyBaseUrl?: string;
  error?: string;
}

export interface WebSearchToolConfig {
  apiKey?: string;
  proxyBaseUrl?: string;
}

export interface WebSearchToolDependencies {
  getConfig?: () => Promise<WebSearchToolConfig>;
  executeSearch?: (
    params: Params,
    config: Required<Pick<WebSearchToolConfig, "apiKey">> & {
      proxyBaseUrl?: string;
    },
    signal: AbortSignal | undefined,
  ) => Promise<{
    hits: WebSearchHit[];
    sentQuery: string;
    proxied: boolean;
    proxyBaseUrl?: string;
  }>;
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

function buildBraveSearchUrl(params: Params): { targetUrl: string; sentQuery: string } {
  const sentQuery = buildSentQuery(params);
  const url = new URL(BRAVE_WEB_SEARCH_ENDPOINT);
  url.searchParams.set("q", sentQuery);
  url.searchParams.set("count", String(params.max_results ?? 5));

  if (params.recency) {
    url.searchParams.set("freshness", recencyToFreshness[params.recency]);
  }

  return {
    targetUrl: url.toString(),
    sentQuery,
  };
}

function parseSearchHits(payload: unknown): WebSearchHit[] {
  if (!isRecord(payload)) return [];
  if (!isRecord(payload.web)) return [];
  const results = payload.web.results;
  if (!Array.isArray(results)) return [];

  const hits: WebSearchHit[] = [];

  for (const entry of results) {
    if (!isRecord(entry)) continue;

    const title = normalizeOptionalString(entry.title);
    const url = normalizeOptionalString(entry.url);
    const description = normalizeOptionalString(entry.description) ?? normalizeOptionalString(entry.snippet) ?? "";

    if (!title || !url) continue;

    hits.push({
      title,
      url,
      snippet: description,
    });
  }

  return hits;
}

function buildResultMarkdown(args: {
  params: Params;
  sentQuery: string;
  hits: WebSearchHit[];
  proxied: boolean;
  proxyBaseUrl?: string;
}): string {
  const { params, sentQuery, hits, proxied, proxyBaseUrl } = args;

  const lines: string[] = [];
  lines.push("Web search via Brave Search");
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

function isAbortError(error: unknown): boolean {
  if (error instanceof DOMException) {
    return error.name === "AbortError";
  }

  if (error instanceof Error) {
    return error.name === "AbortError";
  }

  return false;
}

async function defaultGetConfig(): Promise<WebSearchToolConfig> {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  const settings: ProxyAwareSettingsStore = storageModule.getAppStorage().settings;

  const [providerConfig, proxyBaseUrl] = await Promise.all([
    loadWebSearchProviderConfig(settings),
    getEnabledProxyBaseUrl(settings),
  ]);

  return {
    apiKey: providerConfig.apiKey,
    proxyBaseUrl,
  };
}

async function defaultExecuteSearch(
  params: Params,
  config: Required<Pick<WebSearchToolConfig, "apiKey">> & {
    proxyBaseUrl?: string;
  },
  signal: AbortSignal | undefined,
): Promise<{
  hits: WebSearchHit[];
  sentQuery: string;
  proxied: boolean;
  proxyBaseUrl?: string;
}> {
  const { targetUrl, sentQuery } = buildBraveSearchUrl(params);
  const resolved = resolveOutboundRequestUrl({
    targetUrl,
    proxyBaseUrl: config.proxyBaseUrl,
  });

  const controller = new AbortController();
  const timeoutId = setTimeout(() => {
    controller.abort();
  }, WEB_SEARCH_TIMEOUT_MS);

  const abortFromCaller = () => {
    controller.abort();
  };

  if (signal) {
    if (signal.aborted) {
      controller.abort();
    } else {
      signal.addEventListener("abort", abortFromCaller, { once: true });
    }
  }

  try {
    const response = await fetch(resolved.requestUrl, {
      method: "GET",
      headers: {
        Accept: "application/json",
        "X-Subscription-Token": config.apiKey,
      },
      signal: controller.signal,
    });

    const text = await response.text();

    if (!response.ok) {
      const reason = text.trim().length > 0 ? text.trim() : `HTTP ${response.status}`;
      throw new Error(`Brave search request failed (${response.status}): ${reason}`);
    }

    let payload: unknown = null;
    try {
      payload = JSON.parse(text);
    } catch {
      payload = null;
    }

    const hits = parseSearchHits(payload);

    return {
      hits,
      sentQuery,
      proxied: resolved.proxied,
      proxyBaseUrl: resolved.proxyBaseUrl,
    };
  } catch (error: unknown) {
    if (isAbortError(error)) {
      if (signal?.aborted) {
        throw new Error("Aborted");
      }
      throw new Error(`web_search timed out after ${WEB_SEARCH_TIMEOUT_MS}ms.`);
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
    if (signal) {
      signal.removeEventListener("abort", abortFromCaller);
    }
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
      "Search the public web via Brave Search. Returns compact, cited links with snippets.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<WebSearchToolDetails>> => {
      let params: Params | null = null;

      try {
        params = parseParams(rawParams);

        const config = await getConfig();
        const apiKey = normalizeOptionalString(config.apiKey);
        if (!apiKey) {
          throw new Error(
            "Web search API key is missing. Open /skills and set the Brave API key.",
          );
        }

        const result = await executeSearch(
          params,
          {
            apiKey,
            proxyBaseUrl: config.proxyBaseUrl,
          },
          signal,
        );

        return {
          content: [{ type: "text", text: buildResultMarkdown({
            params,
            sentQuery: result.sentQuery,
            hits: result.hits,
            proxied: result.proxied,
            proxyBaseUrl: result.proxyBaseUrl,
          }) }],
          details: {
            kind: "web_search",
            ok: true,
            provider: "brave",
            query: params.query,
            sentQuery: result.sentQuery,
            recency: params.recency,
            siteFilters: parseSites(params.site),
            maxResults: params.max_results ?? 5,
            resultCount: result.hits.length,
            proxied: result.proxied,
            proxyBaseUrl: result.proxyBaseUrl,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        const fallbackQuery = params?.query
          ?? (isRecord(rawParams) && typeof rawParams.query === "string" ? rawParams.query : "");

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "web_search",
            ok: false,
            provider: "brave",
            query: fallbackQuery,
            sentQuery: fallbackQuery,
            maxResults: params?.max_results ?? 5,
            error: message,
          },
        };
      }
    },
  };
}
