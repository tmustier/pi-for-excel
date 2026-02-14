/**
 * fetch_page — fetch a webpage URL and extract readable markdown content.
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { getErrorMessage } from "../utils/errors.js";
import { runWithTimeoutAbort } from "../utils/network.js";
import {
  getEnabledProxyBaseUrl,
  resolveOutboundRequestUrl,
  type ProxyAwareSettingsStore,
} from "./external-fetch.js";

const FETCH_PAGE_TIMEOUT_MS = 15_000;
const DEFAULT_MAX_CHARS = 12_000;
const MAX_ALLOWED_CHARS = 50_000;
const MIN_ALLOWED_CHARS = 1_000;
const MAX_RAW_BODY_CHARS = 500_000;
const DOMAIN_MIN_INTERVAL_MS = 1_000;

const domainLastRequestAt = new Map<string, number>();

const schema = Type.Object({
  url: Type.String({
    minLength: 1,
    description: "Page URL to fetch. Must use http or https.",
  }),
  max_chars: Type.Optional(Type.Integer({
    minimum: MIN_ALLOWED_CHARS,
    maximum: MAX_ALLOWED_CHARS,
    description: `Maximum extracted characters to return (${MIN_ALLOWED_CHARS}-${MAX_ALLOWED_CHARS}). Default: ${DEFAULT_MAX_CHARS}.`,
  })),
});

type Params = Static<typeof schema>;

export interface FetchPageToolDetails {
  kind: "fetch_page";
  ok: boolean;
  url: string;
  title?: string;
  chars?: number;
  truncated?: boolean;
  proxied?: boolean;
  proxyBaseUrl?: string;
  contentType?: string;
  error?: string;
}

export interface FetchPageToolConfig {
  proxyBaseUrl?: string;
}

interface FetchPageResponse {
  status: number;
  ok: boolean;
  contentType: string;
  body: string;
}

export interface FetchPageToolDependencies {
  getConfig?: () => Promise<FetchPageToolConfig>;
  executeFetch?: (
    requestUrl: string,
    signal: AbortSignal | undefined,
  ) => Promise<FetchPageResponse>;
  now?: () => number;
}

function normalizeOptionalString(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function parseParams(raw: unknown): Params {
  if (!raw || typeof raw !== "object") {
    throw new Error("Invalid fetch_page params: expected an object.");
  }

  const params = raw as Record<string, unknown>;
  const url = normalizeOptionalString(params.url);
  if (!url) {
    throw new Error("fetch_page requires a non-empty url.");
  }

  let maxChars = DEFAULT_MAX_CHARS;
  const maxCharsRaw = params.max_chars;
  if (typeof maxCharsRaw === "number" && Number.isInteger(maxCharsRaw)) {
    if (maxCharsRaw < MIN_ALLOWED_CHARS || maxCharsRaw > MAX_ALLOWED_CHARS) {
      throw new Error(`max_chars must be between ${MIN_ALLOWED_CHARS} and ${MAX_ALLOWED_CHARS}.`);
    }
    maxChars = maxCharsRaw;
  }

  return {
    url,
    max_chars: maxChars,
  };
}

function ensureHttpUrl(rawUrl: string): URL {
  let parsed: URL;
  try {
    parsed = new URL(rawUrl);
  } catch {
    throw new Error("Invalid URL.");
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error("Only http(s) URLs are supported.");
  }

  return parsed;
}

function normalizeText(text: string): string {
  return text
    .replace(/\s+/gu, " ")
    .trim();
}

function stripHtmlTags(html: string): string {
  return normalizeText(
    html
      .replace(/<script[\s\S]*?<\/script(?:\s+[^>]*)?>/giu, " ")
      .replace(/<style[\s\S]*?<\/style(?:\s+[^>]*)?>/giu, " ")
      .replace(/<[^>]+>/gu, " "),
  );
}

function truncateText(text: string, maxChars: number): { text: string; truncated: boolean } {
  if (text.length <= maxChars) {
    return { text, truncated: false };
  }

  return {
    text: `${text.slice(0, maxChars)}\n\n[...truncated]`,
    truncated: true,
  };
}

function htmlToMarkdown(html: string): { title?: string; markdown: string } {
  const parserCtor = typeof DOMParser !== "undefined" ? DOMParser : undefined;
  if (!parserCtor) {
    return {
      markdown: stripHtmlTags(html),
    };
  }

  const doc = new parserCtor().parseFromString(html, "text/html");

  const title = normalizeOptionalString(doc.title);

  const removeSelectors = [
    "script",
    "style",
    "noscript",
    "nav",
    "header",
    "footer",
    "aside",
    "form",
  ];

  for (const selector of removeSelectors) {
    for (const node of doc.querySelectorAll(selector)) {
      node.remove();
    }
  }

  const root = doc.querySelector("main, article, [role='main']") ?? doc.body;
  if (!root) {
    return {
      title,
      markdown: stripHtmlTags(html),
    };
  }

  const parts: string[] = [];
  const elements = root.querySelectorAll("h1, h2, h3, p, li, blockquote, pre");

  for (const element of elements) {
    const text = normalizeText(element.textContent ?? "");
    if (text.length === 0) continue;

    const tag = element.tagName.toLowerCase();
    if (tag === "h1") {
      parts.push(`# ${text}`);
      continue;
    }
    if (tag === "h2") {
      parts.push(`## ${text}`);
      continue;
    }
    if (tag === "h3") {
      parts.push(`### ${text}`);
      continue;
    }
    if (tag === "li") {
      parts.push(`- ${text}`);
      continue;
    }
    if (tag === "blockquote") {
      parts.push(`> ${text}`);
      continue;
    }
    if (tag === "pre") {
      parts.push(`\`${text}\``);
      continue;
    }

    parts.push(text);
  }

  if (parts.length === 0) {
    return {
      title,
      markdown: stripHtmlTags(html),
    };
  }

  const deduped: string[] = [];
  for (const part of parts) {
    if (deduped[deduped.length - 1] === part) continue;
    deduped.push(part);
  }

  return {
    title,
    markdown: deduped.join("\n\n"),
  };
}

function extractReadableMarkdown(args: {
  body: string;
  contentType: string;
  maxChars: number;
}): { title?: string; markdown: string; truncated: boolean } {
  const { body, contentType, maxChars } = args;

  const limitedRaw = body.length > MAX_RAW_BODY_CHARS
    ? body.slice(0, MAX_RAW_BODY_CHARS)
    : body;

  const isHtml = /text\/html/i.test(contentType) || (!/json|xml/i.test(contentType) && /<html|<body|<p|<div/i.test(limitedRaw));

  const extracted = isHtml
    ? htmlToMarkdown(limitedRaw)
    : { markdown: normalizeText(limitedRaw) };

  const truncated = truncateText(extracted.markdown, maxChars);
  return {
    title: extracted.title,
    markdown: truncated.text,
    truncated: truncated.truncated || limitedRaw.length < body.length,
  };
}

function enforceDomainRateLimit(hostname: string, now: number): void {
  const last = domainLastRequestAt.get(hostname);
  if (typeof last === "number" && now - last < DOMAIN_MIN_INTERVAL_MS) {
    throw new Error(`Rate limited for ${hostname}. Please wait ${DOMAIN_MIN_INTERVAL_MS}ms between requests.`);
  }

  domainLastRequestAt.set(hostname, now);
}

async function defaultGetConfig(): Promise<FetchPageToolConfig> {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  const settings: ProxyAwareSettingsStore = storageModule.getAppStorage().settings;
  const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
  return { proxyBaseUrl };
}

async function defaultExecuteFetch(
  requestUrl: string,
  signal: AbortSignal | undefined,
): Promise<FetchPageResponse> {
  const response = await fetch(requestUrl, {
    method: "GET",
    headers: {
      Accept: "text/html, text/plain;q=0.9, application/xhtml+xml;q=0.8",
    },
    signal,
  });

  const body = await response.text();

  return {
    status: response.status,
    ok: response.ok,
    contentType: response.headers.get("content-type") ?? "",
    body,
  };
}

function buildResultMarkdown(args: {
  url: string;
  title?: string;
  markdown: string;
  truncated: boolean;
  proxied: boolean;
  proxyBaseUrl?: string;
}): string {
  const lines: string[] = [];
  lines.push("Fetched page content");
  lines.push("");
  lines.push(`- url: ${args.url}`);
  lines.push(`- transport: ${args.proxied ? `proxy (${args.proxyBaseUrl ?? "configured proxy"})` : "direct"}`);
  if (args.title) {
    lines.push(`- title: ${args.title}`);
  }
  if (args.truncated) {
    lines.push("- note: extracted content was truncated to fit context limits");
  }
  lines.push("");
  lines.push(args.markdown.length > 0 ? args.markdown : "(No readable content found.)");
  return lines.join("\n");
}

export function createFetchPageTool(
  dependencies: FetchPageToolDependencies = {},
): AgentTool<TSchema, FetchPageToolDetails> {
  const getConfig = dependencies.getConfig ?? defaultGetConfig;
  const executeFetch = dependencies.executeFetch ?? defaultExecuteFetch;
  const now = dependencies.now ?? (() => Date.now());

  return {
    name: "fetch_page",
    label: "Fetch Page",
    description: "Fetch a URL and extract readable page content as markdown.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<FetchPageToolDetails>> => {
      let targetUrl = "";

      try {
        const params = parseParams(rawParams);
        const parsedUrl = ensureHttpUrl(params.url);
        targetUrl = parsedUrl.toString();

        enforceDomainRateLimit(parsedUrl.hostname, now());

        const config = await getConfig();
        const resolved = resolveOutboundRequestUrl({
          targetUrl,
          proxyBaseUrl: config.proxyBaseUrl,
        });

        const response = await runWithTimeoutAbort({
          signal,
          timeoutMs: FETCH_PAGE_TIMEOUT_MS,
          timeoutErrorMessage: `fetch_page timed out after ${FETCH_PAGE_TIMEOUT_MS}ms.`,
          run: (requestSignal) => executeFetch(resolved.requestUrl, requestSignal),
        });

        if (!response.ok) {
          const reason = normalizeOptionalString(response.body) ?? `HTTP ${response.status}`;
          throw new Error(`fetch_page request failed (${response.status}): ${reason}`);
        }

        const extracted = extractReadableMarkdown({
          body: response.body,
          contentType: response.contentType,
          maxChars: params.max_chars ?? DEFAULT_MAX_CHARS,
        });

        return {
          content: [{
            type: "text",
            text: buildResultMarkdown({
              url: targetUrl,
              title: extracted.title,
              markdown: extracted.markdown,
              truncated: extracted.truncated,
              proxied: resolved.proxied,
              proxyBaseUrl: resolved.proxyBaseUrl,
            }),
          }],
          details: {
            kind: "fetch_page",
            ok: true,
            url: targetUrl,
            title: extracted.title,
            chars: extracted.markdown.length,
            truncated: extracted.truncated,
            proxied: resolved.proxied,
            proxyBaseUrl: resolved.proxyBaseUrl,
            contentType: response.contentType,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "fetch_page",
            ok: false,
            url: targetUrl,
            error: message,
          },
        };
      }
    },
  };
}
