/**
 * Office taskpanes run in a browser webview.
 *
 * Many LLM providers (and especially subscription / OAuth-based flows) either:
 * - block browser requests via CORS, or
 * - require the "anthropic-dangerous-direct-browser-access" header, which some orgs disable.
 *
 * We route requests through a user-configured local CORS proxy when enabled.
 */

import { streamSimple, type Api, type Context, type Model, type StreamOptions } from "@mariozechner/pi-ai";

import { isDebugEnabled } from "../debug/debug.js";
import { selectToolBundle, type ToolBundleId } from "../context/tool-disclosure.js";
import { normalizeProxyUrl, validateOfficeProxyUrl } from "./proxy-validation.js";

export type GetProxyUrl = () => Promise<string | undefined>;

function shouldProxyProvider(provider: string, apiKey?: string): boolean {
  const p = provider.toLowerCase();

  switch (p) {
    // Known to require proxy in browser webviews (CORS blocked)
    case "openai-codex":
      return true;

    // Anthropic OAuth tokens are blocked by CORS; some orgs also block direct browser access.
    // We proxy OAuth tokens (sk-ant-oat-*) unconditionally.
    case "anthropic":
      return typeof apiKey === "string" && apiKey.startsWith("sk-ant-oat");

    // Google Cloud Code Assist providers frequently require proxy in Office webviews.
    case "google-gemini-cli":
    case "google-antigravity":
      return true;

    // Z-AI always requires proxy (matches pi-web-ui default)
    case "zai":
      return true;

    default:
      return false;
  }
}

function applyProxy(model: Model<Api>, proxyUrl: string): Model<Api> {
  if (!model.baseUrl) return model;
  if (!/^https?:\/\//i.test(model.baseUrl)) return model;

  const normalizedProxy = normalizeProxyUrl(proxyUrl);

  // Avoid double-proxying
  if (model.baseUrl.startsWith(`${normalizedProxy}/?url=`)) return model;

  return {
    ...model,
    baseUrl: `${normalizedProxy}/?url=${encodeURIComponent(model.baseUrl)}`,
  };
}

/**
 * Is this call a tool-result continuation?
 *
 * Used for debug/telemetry (payload snapshots and status pills), not for
 * capability gating. Continuations still receive a deterministic tool bundle
 * so the agent can complete multi-step tool loops in a single turn.
 */
function isToolContinuation(messages: Context["messages"]): boolean {
  if (messages.length === 0) return false;
  const last = messages[messages.length - 1];
  return last.role === "toolResult";
}

// ---------------------------------------------------------------------------
// Payload stats — lightweight counters for debug pill in status bar.
// ---------------------------------------------------------------------------

export interface PayloadStats {
  /** LLM calls since app start. */
  calls: number;
  /** Chars of system prompt on last call. */
  systemChars: number;
  /** Chars of tool schemas (compact JSON) on last call, 0 if no tools. */
  toolSchemaChars: number;
  /** Number of tool definitions on last call, 0 if no tools. */
  toolCount: number;
  /** Number of messages on last call. */
  messageCount: number;
  /** Total chars of all messages (JSON-serialized) on last call. */
  messageChars: number;
}

export interface PayloadShapeSummary {
  rootType: "object" | "array" | "primitive" | "null";
  topLevelKeys: string[];
  rootArrayLength?: number;
  arrayFields: Array<{ key: string; length: number }>;
}

export interface PayloadSnapshot {
  call: number;
  timestamp: number;
  sessionId?: string;
  provider: string;
  modelId: string;
  isToolContinuation: boolean;
  toolBundle: ToolBundleId;
  toolsIncluded: boolean;
  systemChars: number;
  toolSchemaChars: number;
  messageChars: number;
  totalChars: number;
  toolCount: number;
  messageCount: number;
  payloadShape?: PayloadShapeSummary;
}

const stats: PayloadStats = {
  calls: 0,
  systemChars: 0,
  toolSchemaChars: 0,
  toolCount: 0,
  messageCount: 0,
  messageChars: 0,
};

/**
 * Debug retention limits.
 *
 * Why 24?
 * - Large enough to inspect recent multi-step tool loops in one session
 *   (typically several user turns with follow-up calls).
 * - Small enough to keep memory bounded in long-lived taskpane sessions.
 *
 * If we need deeper forensic traces, we can raise this temporarily,
 * but should keep it bounded to avoid silent memory growth.
 */
const MAX_PAYLOAD_SNAPSHOTS = 24;

/**
 * Number of session-scoped contexts kept for debug inspection.
 *
 * This is separate from the per-call ring buffer above and is capped for
 * the same reason: preserve useful visibility without unbounded retention.
 */
const MAX_SESSION_CONTEXTS = 24;

const payloadSnapshots: PayloadSnapshot[] = [];

/** Snapshot of the last LLM context (only kept when debug is on). */
let lastContext: Context | undefined;
const lastContextBySession = new Map<string, Context>();

export function getPayloadStats(): Readonly<PayloadStats> {
  return stats;
}

export function getPayloadSnapshots(): readonly PayloadSnapshot[] {
  if (!isDebugEnabled()) return [];
  return payloadSnapshots;
}

export function getLastContext(sessionId?: string): Context | undefined {
  if (!isDebugEnabled()) return undefined;
  if (sessionId) return lastContextBySession.get(sessionId);
  return lastContext;
}

function getSessionId(options: StreamOptions | undefined): string | undefined {
  const sessionId = options?.sessionId;
  if (typeof sessionId !== "string") return undefined;
  if (sessionId.trim().length === 0) return undefined;
  return sessionId;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function summarizePayloadShape(payload: unknown): PayloadShapeSummary {
  if (payload === null) {
    return {
      rootType: "null",
      topLevelKeys: [],
      arrayFields: [],
    };
  }

  if (Array.isArray(payload)) {
    return {
      rootType: "array",
      topLevelKeys: [],
      rootArrayLength: payload.length,
      arrayFields: [],
    };
  }

  if (!isRecord(payload)) {
    return {
      rootType: "primitive",
      topLevelKeys: [],
      arrayFields: [],
    };
  }

  const entries = Object.entries(payload);
  const arrayFields: Array<{ key: string; length: number }> = [];
  for (const [key, value] of entries) {
    if (Array.isArray(value)) {
      arrayFields.push({ key, length: value.length });
    }
  }

  return {
    rootType: "object",
    topLevelKeys: entries.map(([key]) => key).sort(),
    arrayFields,
  };
}

function pushSnapshot(snapshot: PayloadSnapshot): void {
  payloadSnapshots.push(snapshot);
  if (payloadSnapshots.length > MAX_PAYLOAD_SNAPSHOTS) {
    const overflow = payloadSnapshots.length - MAX_PAYLOAD_SNAPSHOTS;
    payloadSnapshots.splice(0, overflow);
  }
}

function setSessionContext(sessionId: string, context: Context): void {
  // Refresh insertion order to behave like an LRU cache.
  if (lastContextBySession.has(sessionId)) {
    lastContextBySession.delete(sessionId);
  }
  lastContextBySession.set(sessionId, context);

  if (lastContextBySession.size <= MAX_SESSION_CONTEXTS) return;

  const oldest = lastContextBySession.keys().next().value;
  if (typeof oldest === "string") {
    lastContextBySession.delete(oldest);
  }
}

function upsertPayloadShape(call: number, payload: unknown): void {
  let index = -1;
  for (let i = payloadSnapshots.length - 1; i >= 0; i -= 1) {
    if (payloadSnapshots[i].call === call) {
      index = i;
      break;
    }
  }
  if (index < 0) return;

  payloadSnapshots[index] = {
    ...payloadSnapshots[index],
    payloadShape: summarizePayloadShape(payload),
  };

  document.dispatchEvent(new Event("pi:status-update"));
}

function recordCall(
  model: Model<Api>,
  context: Context,
  options: StreamOptions | undefined,
  continuation: boolean,
  toolBundle: ToolBundleId,
): { call: number; captureSnapshot: boolean } {
  stats.calls += 1;
  stats.systemChars = context.systemPrompt?.length ?? 0;
  stats.messageCount = context.messages.length;

  let msgChars = 0;
  for (const m of context.messages) msgChars += JSON.stringify(m).length;
  stats.messageChars = msgChars;

  if (context.tools) {
    stats.toolCount = context.tools.length;
    let chars = 0;
    for (const t of context.tools) chars += JSON.stringify(t).length;
    stats.toolSchemaChars = chars;
  } else {
    stats.toolCount = 0;
    stats.toolSchemaChars = 0;
  }

  const call = stats.calls;
  const captureSnapshot = isDebugEnabled();

  if (captureSnapshot) {
    lastContext = context;

    const sessionId = getSessionId(options);
    if (sessionId) {
      setSessionContext(sessionId, context);
    }

    const totalChars = stats.systemChars + stats.toolSchemaChars + stats.messageChars;
    pushSnapshot({
      call,
      timestamp: Date.now(),
      sessionId,
      provider: model.provider,
      modelId: model.id,
      isToolContinuation: continuation,
      toolBundle,
      toolsIncluded: stats.toolCount > 0,
      systemChars: stats.systemChars,
      toolSchemaChars: stats.toolSchemaChars,
      messageChars: stats.messageChars,
      totalChars,
      toolCount: stats.toolCount,
      messageCount: stats.messageCount,
    });
  }

  document.dispatchEvent(new Event("pi:status-update"));
  return { call, captureSnapshot };
}

function withPayloadHook(
  options: StreamOptions | undefined,
  call: number,
  captureSnapshot: boolean,
): StreamOptions | undefined {
  const originalOnPayload = options?.onPayload;
  if (!captureSnapshot && !originalOnPayload) return options;

  const onPayload = (payload: unknown) => {
    if (captureSnapshot) {
      upsertPayloadShape(call, payload);
    }
    originalOnPayload?.(payload);
  };

  if (options) return { ...options, onPayload };
  return { onPayload };
}

/**
 * Create a StreamFn compatible with Agent that proxies provider base URLs when needed.
 */
export function createOfficeStreamFn(getProxyUrl: GetProxyUrl) {
  return async (model: Model<Api>, context: Context, options?: StreamOptions) => {
    const continuation = isToolContinuation(context.messages);

    // Always expose tools (via deterministic bundle selection), including
    // continuation calls after tool results. This preserves full agent loops.
    const toolSelection = selectToolBundle(context);

    const effectiveContext = toolSelection.tools === context.tools
      ? context
      : { ...context, tools: toolSelection.tools };

    const callRecord = recordCall(
      model,
      effectiveContext,
      options,
      continuation,
      toolSelection.bundleId,
    );
    const effectiveOptions = withPayloadHook(options, callRecord.call, callRecord.captureSnapshot);

    const proxyUrl = await getProxyUrl();
    if (!proxyUrl) {
      return streamSimple(model, effectiveContext, effectiveOptions);
    }

    if (!shouldProxyProvider(model.provider, options?.apiKey)) {
      return streamSimple(model, effectiveContext, effectiveOptions);
    }

    // Guardrails: fail fast for known-bad proxy configs (e.g., HTTP proxy from HTTPS taskpane).
    const validated = validateOfficeProxyUrl(proxyUrl);

    return streamSimple(applyProxy(model, validated), effectiveContext, effectiveOptions);
  };
}
