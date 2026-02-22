import type {
  Api,
  AssistantMessage,
  Message,
  Model,
  Usage,
} from "@mariozechner/pi-ai";
import { getModels, getProviders } from "@mariozechner/pi-ai";

import type { HttpRequestOptions } from "../commands/extension-api.js";
import { isRecord } from "../utils/type-guards.js";
import type { StoredExtensionSource } from "./store.js";

const DEFAULT_EXTENSION_HTTP_TIMEOUT_MS = 15_000;
const MAX_EXTENSION_HTTP_TIMEOUT_MS = 30_000;
const MAX_EXTENSION_HTTP_BODY_BYTES = 1_000_000;

export function getRuntimeManagerErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}

export function isApiModel(model: unknown): model is Model<Api> {
  if (!isRecord(model)) {
    return false;
  }

  return (
    typeof model.id === "string"
    && typeof model.provider === "string"
    && typeof model.api === "string"
    && typeof model.name === "string"
  );
}

function createZeroUsage(): Usage {
  return {
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheWrite: 0,
    totalTokens: 0,
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      total: 0,
    },
  };
}

export function resolveModelForCompletion(args: {
  fallbackModel: Model<Api>;
  requestedModel?: string;
}): Model<Api> {
  const { fallbackModel, requestedModel } = args;

  if (!requestedModel) {
    return fallbackModel;
  }

  const trimmed = requestedModel.trim();
  if (trimmed.length === 0) {
    return fallbackModel;
  }

  const findModelByProviderAndId = (providerName: string, modelId: string): Model<Api> | null => {
    for (const provider of getProviders()) {
      if (provider !== providerName) {
        continue;
      }

      const match = getModels(provider).find((model) => model.id === modelId);
      if (match) {
        return match;
      }
    }

    return null;
  };

  const slashIndex = trimmed.indexOf("/");
  if (slashIndex > 0 && slashIndex < trimmed.length - 1) {
    const requestedProvider = trimmed.slice(0, slashIndex);
    const requestedId = trimmed.slice(slashIndex + 1);
    const match = findModelByProviderAndId(requestedProvider, requestedId);

    if (!match) {
      throw new Error(`Unknown model: ${trimmed}`);
    }

    return match;
  }

  const providerMatch = findModelByProviderAndId(fallbackModel.provider, trimmed);
  if (providerMatch) {
    return providerMatch;
  }

  for (const provider of getProviders()) {
    const match = getModels(provider).find((model) => model.id === trimmed);
    if (match) {
      return match;
    }
  }

  throw new Error(`Unknown model: ${trimmed}`);
}

export function parseLlmMessages(
  messages: readonly { role: "user" | "assistant"; content: string }[],
  model: Model<Api>,
): Message[] {
  const parsed: Message[] = [];
  let timestamp = Date.now();

  for (const message of messages) {
    const content = typeof message.content === "string" ? message.content : String(message.content);

    if (message.role === "user") {
      parsed.push({
        role: "user",
        content: [{ type: "text", text: content }],
        timestamp,
      });
      timestamp += 1;
      continue;
    }

    const assistantMessage: AssistantMessage = {
      role: "assistant",
      content: [{ type: "text", text: content }],
      api: model.api,
      provider: model.provider,
      model: model.id,
      usage: createZeroUsage(),
      stopReason: "stop",
      timestamp,
    };

    parsed.push(assistantMessage);
    timestamp += 1;
  }

  return parsed;
}

export function extractAssistantText(message: AssistantMessage): string {
  return message.content
    .flatMap((item) => {
      if (item.type !== "text") {
        return [];
      }

      return [item.text];
    })
    .join("");
}

function normalizeCompletionMessageContent(content: string, label: string): string {
  const normalized = content.trim();
  if (normalized.length === 0) {
    throw new Error(`${label} cannot be empty.`);
  }

  return normalized;
}

export function normalizeDownloadFilename(filename: string): string {
  const trimmed = filename.trim();
  if (trimmed.length === 0) {
    throw new Error("Download filename cannot be empty.");
  }

  return trimmed;
}

export function createExtensionAgentMessage(extensionName: string, label: string, content: string): Message {
  const normalizedContent = normalizeCompletionMessageContent(content, label);

  return {
    role: "user",
    content: [{
      type: "text",
      text: `[Extension ${extensionName}]\n${normalizedContent}`,
    }],
    timestamp: Date.now(),
  };
}

const EXTENSION_LLM_SESSION_KEY_SEGMENT = "ext-llm";

function normalizeExtensionSessionKeyPart(value: string | undefined): string {
  if (typeof value !== "string") {
    return "";
  }

  return value.trim();
}

/**
 * Side `llm.complete` requests are intentionally independent from the main
 * runtime loop. Use a deterministic extension-scoped session key so prefix
 * churn/debug signals don't get mixed into the primary runtime session key.
 */
export function createExtensionLlmCompletionSessionId(args: {
  agentSessionId: string | undefined;
  extensionId: string;
}): string {
  const normalizedAgentSessionId = normalizeExtensionSessionKeyPart(args.agentSessionId);
  const normalizedExtensionId = normalizeExtensionSessionKeyPart(args.extensionId);
  const extensionSegment = normalizedExtensionId.length > 0 ? normalizedExtensionId : "unknown-extension";

  if (normalizedAgentSessionId.length === 0) {
    return `${EXTENSION_LLM_SESSION_KEY_SEGMENT}:${extensionSegment}`;
  }

  return `${normalizedAgentSessionId}::${EXTENSION_LLM_SESSION_KEY_SEGMENT}:${extensionSegment}`;
}

export function isBlockedExtensionHostname(hostname: string): boolean {
  const normalized = hostname.trim().toLowerCase();
  if (normalized.length === 0) {
    return true;
  }

  const unwrapped = normalized.startsWith("[") && normalized.endsWith("]")
    ? normalized.slice(1, -1)
    : normalized;

  if (unwrapped === "localhost" || unwrapped.endsWith(".localhost") || unwrapped.endsWith(".local")) {
    return true;
  }

  if (unwrapped === "0.0.0.0" || unwrapped === "127.0.0.1") {
    return true;
  }

  if (unwrapped === "::1") {
    return true;
  }

  const ipv4Match = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/u.exec(unwrapped);

  if (ipv4Match) {
    const octets = ipv4Match.slice(1).map((part) => Number(part));
    if (octets.some((octet) => Number.isNaN(octet) || octet < 0 || octet > 255)) {
      return true;
    }

    const [a, b] = octets;
    if (a === 10 || a === 127 || a === 0) {
      return true;
    }

    if (a === 169 && b === 254) {
      return true;
    }

    if (a === 192 && b === 168) {
      return true;
    }

    if (a === 172 && b >= 16 && b <= 31) {
      return true;
    }

    return false;
  }

  if (unwrapped.includes(":")) {
    if (unwrapped.startsWith("fd") || unwrapped.startsWith("fc") || unwrapped.startsWith("fe80:")) {
      return true;
    }
  }

  return false;
}

export function normalizeHttpOptions(options: HttpRequestOptions | undefined): {
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE" | "HEAD";
  headers: Record<string, string>;
  body: string | undefined;
  timeoutMs: number;
} {
  const candidateMethod = options?.method ?? "GET";
  const method = candidateMethod === "GET"
    || candidateMethod === "POST"
    || candidateMethod === "PUT"
    || candidateMethod === "PATCH"
    || candidateMethod === "DELETE"
    || candidateMethod === "HEAD"
    ? candidateMethod
    : "GET";
  const headers = options?.headers ?? {};
  const body = options?.body;
  const timeout = options?.timeoutMs ?? DEFAULT_EXTENSION_HTTP_TIMEOUT_MS;
  const boundedTimeout = Math.max(1, Math.min(MAX_EXTENSION_HTTP_TIMEOUT_MS, timeout));

  return {
    method,
    headers,
    body,
    timeoutMs: boundedTimeout,
  };
}

export async function readLimitedResponseBody(response: Response): Promise<string> {
  const bodyText = await response.text();
  const byteLength = new TextEncoder().encode(bodyText).length;

  if (byteLength > MAX_EXTENSION_HTTP_BODY_BYTES) {
    throw new Error(
      `HTTP response body too large (${byteLength} bytes). Limit is ${MAX_EXTENSION_HTTP_BODY_BYTES} bytes.`,
    );
  }

  return bodyText;
}

export function normalizeExtensionName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Extension name cannot be empty");
  }

  return trimmed;
}

export function normalizeInlineCode(code: string): string {
  if (code.trim().length === 0) {
    throw new Error("Extension code cannot be empty");
  }

  return code;
}

export function normalizeRemoteUrl(url: string): string {
  let parsed: URL;
  try {
    parsed = new URL(url.trim());
  } catch {
    throw new Error("Invalid URL");
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error("Extension URL must use http:// or https://");
  }

  return parsed.toString();
}

export function describeExtensionSource(source: StoredExtensionSource): string {
  if (source.kind === "module") {
    return source.specifier;
  }

  const lines = source.code.split("\n").length;
  return `inline code (${source.code.length} chars, ${lines} lines)`;
}
