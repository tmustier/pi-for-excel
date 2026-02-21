/**
 * tmux — Experimental local tmux bridge adapter.
 *
 * This tool stays registered for a stable tool list/prompt cache,
 * but execution is gated by:
 * - bridge URL override from /experimental tmux-bridge-url (or default https://localhost:3341)
 * - reachable bridge health endpoint
 *
 * The local bridge contract (v1) is a POST JSON request to /v1/tmux.
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";
import { getErrorMessage } from "../utils/errors.js";
import { isRecord } from "../utils/type-guards.js";
import {
  extractBridgeErrorMessage,
  isAbortError,
  joinBridgeUrl,
  tryParseBridgeJson,
} from "./bridge-http-utils.js";
import {
  DEFAULT_TMUX_BRIDGE_URL,
  TMUX_BRIDGE_URL_SETTING_KEY,
} from "./experimental-tool-gates.js";

const TMUX_BRIDGE_API_PATH = "/v1/tmux";
const DEFAULT_TMUX_BRIDGE_TIMEOUT_MS = 15_000;
const TMUX_BRIDGE_TIMEOUT_BUFFER_MS = 5_000;
const MAX_TMUX_BRIDGE_TIMEOUT_MS = 245_000;
const DEFAULT_SEND_AND_CAPTURE_TIMEOUT_MS = 5_000;

export const TMUX_BRIDGE_TOKEN_SETTING_KEY = "tmux.bridge.token";

const TMUX_ACTIONS = [
  "list_sessions",
  "create_session",
  "send_keys",
  "capture_pane",
  "send_and_capture",
  "kill_session",
] as const;

type TmuxAction = (typeof TMUX_ACTIONS)[number];

const TMUX_ACTION_SET = new Set<string>(TMUX_ACTIONS);

function isTmuxAction(value: unknown): value is TmuxAction {
  return typeof value === "string" && TMUX_ACTION_SET.has(value);
}

function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((value) => Type.Literal(value)),
    opts,
  );
}

const schema = Type.Object({
  action: StringEnum([...TMUX_ACTIONS], {
    description:
      "Tmux operation to run on the local bridge. " +
      "Use list_sessions first, then create_session/send_keys/capture_pane. " +
      "For long-running commands, prefer send_and_capture with wait_for or capture_pane with wait_ms.",
  }),
  session: Type.Optional(Type.String({
    description:
      "Target tmux session name. Required for send_keys, capture_pane, send_and_capture, and kill_session. " +
      "Optional for create_session (bridge may auto-generate a name).",
  })),
  cwd: Type.Optional(Type.String({
    description:
      "Working directory used when creating a session. Only applies to create_session.",
  })),
  text: Type.Optional(Type.String({
    description:
      "Literal text to send to the tmux pane (for example shell commands like `pi`). " +
      "Applies to send_keys and send_and_capture.",
  })),
  keys: Type.Optional(Type.Array(Type.String({
    description:
      "Additional tmux key tokens (for example: Enter, C-c, Up). " +
      "Applies to send_keys and send_and_capture.",
  }), {
    description: "Optional key token sequence sent in order.",
  })),
  enter: Type.Optional(Type.Boolean({
    description: "If true, append Enter after sending text/keys.",
  })),
  lines: Type.Optional(Type.Integer({
    minimum: 1,
    maximum: 5000,
    description:
      "How many trailing pane lines to capture. Applies to capture_pane and send_and_capture.",
  })),
  wait_for: Type.Optional(Type.String({
    description:
      "Regex string to wait for before capturing output (send_and_capture only).",
  })),
  timeout_ms: Type.Optional(Type.Integer({
    minimum: 100,
    maximum: 120_000,
    description: "Max wait time in ms for send_and_capture.",
  })),
  wait_ms: Type.Optional(Type.Integer({
    minimum: 0,
    maximum: 120_000,
    description:
      "Optional pause before capture (capture_pane/send_and_capture). " +
      "Use this to avoid rapid polling for long-running commands.",
  })),
  join_wrapped: Type.Optional(Type.Boolean({
    description:
      "If true, request wrapped terminal lines to be joined by the bridge before return.",
  })),
});

type Params = Static<typeof schema>;

export interface TmuxBridgeConfig {
  url: string;
  token?: string;
}

export interface TmuxBridgeRequest {
  action: TmuxAction;
  session?: string;
  cwd?: string;
  text?: string;
  keys?: string[];
  enter?: boolean;
  lines?: number;
  wait_for?: string;
  timeout_ms?: number;
  wait_ms?: number;
  join_wrapped?: boolean;
}

export interface TmuxBridgeResponse {
  ok: boolean;
  action: TmuxAction;
  session?: string;
  sessions?: string[];
  output?: string;
  error?: string;
  metadata?: Record<string, unknown>;
}

export interface TmuxToolDetails {
  kind: "tmux_bridge";
  ok: boolean;
  action: TmuxAction;
  bridgeUrl?: string;
  session?: string;
  sessionsCount?: number;
  outputPreview?: string;
  error?: string;
}

export interface TmuxToolDependencies {
  getBridgeConfig?: () => Promise<TmuxBridgeConfig | null>;
  callBridge?: (
    request: TmuxBridgeRequest,
    config: TmuxBridgeConfig,
    signal: AbortSignal | undefined,
  ) => Promise<TmuxBridgeResponse>;
}

function cleanOptionalString(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function cleanOptionalStringArray(values: string[] | undefined): string[] | undefined {
  if (!Array.isArray(values)) return undefined;

  const cleaned = values
    .map((value) => value.trim())
    .filter((value) => value.length > 0);

  return cleaned.length > 0 ? cleaned : undefined;
}

function toOptionalBoolean(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

function toOptionalInteger(value: unknown): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  return value;
}

function parseParams(raw: unknown): Params {
  if (!isRecord(raw)) {
    throw new Error("Invalid tmux params: expected an object.");
  }

  if (!isTmuxAction(raw.action)) {
    throw new Error("Invalid tmux action.");
  }

  const keys = parseStringArray(raw.keys);
  const params: Params = {
    action: raw.action,
  };

  if (typeof raw.session === "string") params.session = raw.session;
  if (typeof raw.cwd === "string") params.cwd = raw.cwd;
  if (typeof raw.text === "string") params.text = raw.text;
  if (keys !== undefined) params.keys = keys;

  const enter = toOptionalBoolean(raw.enter);
  if (enter !== undefined) params.enter = enter;

  const lines = toOptionalInteger(raw.lines);
  if (lines !== undefined) params.lines = lines;

  const waitFor = typeof raw.wait_for === "string" ? raw.wait_for : undefined;
  if (waitFor !== undefined) params.wait_for = waitFor;

  const timeoutMs = toOptionalInteger(raw.timeout_ms);
  if (timeoutMs !== undefined) params.timeout_ms = timeoutMs;

  const waitMs = toOptionalInteger(raw.wait_ms);
  if (waitMs !== undefined) params.wait_ms = waitMs;

  const joinWrapped = toOptionalBoolean(raw.join_wrapped);
  if (joinWrapped !== undefined) params.join_wrapped = joinWrapped;

  return params;
}

function hasSendPayload(params: Params): boolean {
  const hasText = cleanOptionalString(params.text) !== undefined;
  const hasKeys = cleanOptionalStringArray(params.keys) !== undefined;
  return hasText || hasKeys || params.enter === true;
}

function requireSession(params: Params): string {
  const session = cleanOptionalString(params.session);
  if (!session) {
    throw new Error(`session is required for ${params.action}`);
  }
  return session;
}

function validateActionParams(params: Params): void {
  if (params.lines !== undefined && (params.lines < 1 || params.lines > 5000)) {
    throw new Error("lines must be between 1 and 5000");
  }

  if (params.timeout_ms !== undefined && (params.timeout_ms < 100 || params.timeout_ms > 120_000)) {
    throw new Error("timeout_ms must be between 100 and 120000");
  }

  if (params.wait_ms !== undefined && (params.wait_ms < 0 || params.wait_ms > 120_000)) {
    throw new Error("wait_ms must be between 0 and 120000");
  }

  switch (params.action) {
    case "list_sessions":
      return;

    case "create_session":
      return;

    case "capture_pane":
    case "kill_session":
      requireSession(params);
      return;

    case "send_keys":
    case "send_and_capture": {
      requireSession(params);
      if (!hasSendPayload(params)) {
        throw new Error(
          `${params.action} requires at least one of: text, keys, or enter=true`,
        );
      }
      return;
    }
  }
}

function toBridgeRequest(params: Params): TmuxBridgeRequest {
  return {
    action: params.action,
    session: cleanOptionalString(params.session),
    cwd: cleanOptionalString(params.cwd),
    text: cleanOptionalString(params.text),
    keys: cleanOptionalStringArray(params.keys),
    enter: params.enter,
    lines: params.lines,
    wait_for: cleanOptionalString(params.wait_for),
    timeout_ms: params.timeout_ms,
    wait_ms: params.wait_ms,
    join_wrapped: params.join_wrapped,
  };
}

function parseStringArray(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined;

  const out: string[] = [];
  for (const item of value) {
    if (typeof item === "string") {
      out.push(item);
    }
  }

  return out.length > 0 ? out : [];
}

function parseBridgeResponse(value: unknown, fallbackAction: TmuxAction): TmuxBridgeResponse {
  if (!isRecord(value)) {
    return {
      ok: true,
      action: fallbackAction,
    };
  }

  const action = isTmuxAction(value.action) ? value.action : fallbackAction;
  const ok = typeof value.ok === "boolean" ? value.ok : true;
  const session = typeof value.session === "string" ? value.session : undefined;
  const sessions = parseStringArray(value.sessions);
  const output = typeof value.output === "string" ? value.output : undefined;
  const error = typeof value.error === "string" ? value.error : undefined;
  const metadata = isRecord(value.metadata) ? value.metadata : undefined;

  return {
    ok,
    action,
    session,
    sessions,
    output,
    error,
    metadata,
  };
}

async function defaultGetBridgeConfig(): Promise<TmuxBridgeConfig | null> {
  let rawUrl = DEFAULT_TMUX_BRIDGE_URL;
  let token: string | undefined;

  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const settings = storageModule.getAppStorage().settings;

    const urlValue = await settings.get<string>(TMUX_BRIDGE_URL_SETTING_KEY);
    const configuredUrl = typeof urlValue === "string" ? urlValue.trim() : "";
    if (configuredUrl.length > 0) {
      rawUrl = configuredUrl;
    }

    const tokenValue = await settings.get<string>(TMUX_BRIDGE_TOKEN_SETTING_KEY);
    token = typeof tokenValue === "string" && tokenValue.trim().length > 0
      ? tokenValue.trim()
      : undefined;
  } catch {
    // Fall back to default localhost URL when settings are unavailable.
  }

  try {
    const normalizedUrl = validateOfficeProxyUrl(rawUrl);
    return {
      url: normalizedUrl,
      token,
    };
  } catch {
    return null;
  }
}

export function computeTmuxFetchTimeoutMs(request: TmuxBridgeRequest): number {
  const waitMs = typeof request.wait_ms === "number" ? request.wait_ms : 0;

  const captureTimeoutMs =
    request.action === "send_and_capture"
      ? (typeof request.timeout_ms === "number"
        ? request.timeout_ms
        : DEFAULT_SEND_AND_CAPTURE_TIMEOUT_MS)
      : 0;

  const computedTimeoutMs = waitMs + captureTimeoutMs + TMUX_BRIDGE_TIMEOUT_BUFFER_MS;
  if (computedTimeoutMs < DEFAULT_TMUX_BRIDGE_TIMEOUT_MS) {
    return DEFAULT_TMUX_BRIDGE_TIMEOUT_MS;
  }

  return Math.min(computedTimeoutMs, MAX_TMUX_BRIDGE_TIMEOUT_MS);
}

async function defaultCallBridge(
  request: TmuxBridgeRequest,
  config: TmuxBridgeConfig,
  signal: AbortSignal | undefined,
): Promise<TmuxBridgeResponse> {
  const endpoint = joinBridgeUrl(config.url, TMUX_BRIDGE_API_PATH);
  const controller = new AbortController();
  const timeoutMs = computeTmuxFetchTimeoutMs(request);

  const timeoutId = setTimeout(() => {
    controller.abort();
  }, timeoutMs);

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

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };

  if (config.token) {
    headers.Authorization = `Bearer ${config.token}`;
  }

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(request),
      signal: controller.signal,
    });

    const rawBody = await response.text();
    const parsedBody = tryParseBridgeJson(rawBody);

    if (!response.ok) {
      const payloadError = extractBridgeErrorMessage(parsedBody);
      const textError = rawBody.trim().length > 0 ? rawBody.trim() : null;
      const reason = payloadError ?? textError ?? `HTTP ${response.status}`;
      throw new Error(`Tmux bridge request failed (${response.status}): ${reason}`);
    }

    if (parsedBody === null) {
      return {
        ok: true,
        action: request.action,
        session: request.session,
        output: rawBody.trim().length > 0 ? rawBody : undefined,
      };
    }

    const parsed = parseBridgeResponse(parsedBody, request.action);
    if (!parsed.ok) {
      throw new Error(parsed.error ?? "Tmux bridge rejected the request.");
    }

    return parsed;
  } catch (error: unknown) {
    if (isAbortError(error)) {
      if (signal?.aborted) {
        throw new Error("Aborted");
      }
      throw new Error(`Tmux bridge request timed out after ${timeoutMs}ms.`);
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
    if (signal) {
      signal.removeEventListener("abort", abortFromCaller);
    }
  }
}

function renderOutputBlock(output: string | undefined): string {
  if (!output) return "";

  const trimmed = output.trim();
  if (trimmed.length === 0) return "";

  return `\n\n\`\`\`\n${trimmed}\n\`\`\``;
}

function formatBridgeSuccessText(
  request: TmuxBridgeRequest,
  response: TmuxBridgeResponse,
): string {
  switch (request.action) {
    case "list_sessions": {
      const sessions = response.sessions ?? [];
      if (sessions.length === 0) {
        return "No tmux sessions found.";
      }

      return `Tmux sessions:\n${sessions.map((session) => `- ${session}`).join("\n")}`;
    }

    case "create_session": {
      const session = response.session ?? request.session ?? "(bridge-assigned)";
      return `Created tmux session \"${session}\".${renderOutputBlock(response.output)}`;
    }

    case "send_keys": {
      const session = response.session ?? request.session ?? "(unknown session)";
      const outputBlock = renderOutputBlock(response.output);
      if (outputBlock.length > 0) {
        return `Sent keys to tmux session \"${session}\".${outputBlock}`;
      }

      return (
        `Sent keys to tmux session \"${session}\". ` +
        "Use send_and_capture (with wait_for when possible) or capture_pane (optionally with wait_ms) to fetch terminal output."
      );
    }

    case "capture_pane": {
      const session = response.session ?? request.session ?? "(unknown session)";
      const output = response.output?.trim();
      if (!output) {
        return `Captured tmux pane for \"${session}\" (no output).`;
      }
      return `Captured tmux pane for \"${session}\":${renderOutputBlock(response.output)}`;
    }

    case "send_and_capture": {
      const session = response.session ?? request.session ?? "(unknown session)";
      const output = response.output?.trim();
      if (!output) {
        return `Sent keys and captured tmux pane for \"${session}\" (no output).`;
      }
      return `Sent keys and captured tmux pane for \"${session}\":${renderOutputBlock(response.output)}`;
    }

    case "kill_session": {
      const session = response.session ?? request.session ?? "(unknown session)";
      return `Killed tmux session \"${session}\".`;
    }
  }
}

function buildOutputPreview(output: string | undefined): string | undefined {
  if (!output) return undefined;

  const trimmed = output.trim();
  if (trimmed.length === 0) return undefined;

  const firstLine = trimmed.split("\n")[0] ?? trimmed;
  return firstLine.length > 180 ? `${firstLine.slice(0, 177)}…` : firstLine;
}

function buildMissingBridgeConfigurationMessage(): string {
  return (
    "Tmux bridge URL is unavailable. " +
    "By default Pi uses https://localhost:3341; set /experimental tmux-bridge-url <url> to override it."
  );
}

export function createTmuxTool(
  dependencies: TmuxToolDependencies = {},
): AgentTool<TSchema, TmuxToolDetails> {
  const getBridgeConfig = dependencies.getBridgeConfig ?? defaultGetBridgeConfig;
  const callBridge = dependencies.callBridge ?? defaultCallBridge;

  return {
    name: "tmux",
    label: "Tmux",
    description:
      "Interact with a local tmux bridge. " +
      "Actions: list/create/send/capture/kill sessions for local shell workflows, including launching installed CLIs like `pi`.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<TmuxToolDetails>> => {
      let params: Params | null = null;

      try {
        params = parseParams(rawParams);
        validateActionParams(params);

        const bridgeConfig = await getBridgeConfig();
        if (!bridgeConfig) {
          return {
            content: [{ type: "text", text: buildMissingBridgeConfigurationMessage() }],
            details: {
              kind: "tmux_bridge",
              ok: false,
              action: params.action,
              error: "missing_bridge_url",
            },
          };
        }

        const request = toBridgeRequest(params);
        const response = await callBridge(request, bridgeConfig, signal);

        if (!response.ok) {
          throw new Error(response.error ?? "Tmux bridge rejected the request.");
        }

        const session = response.session ?? request.session;

        return {
          content: [{ type: "text", text: formatBridgeSuccessText(request, response) }],
          details: {
            kind: "tmux_bridge",
            ok: true,
            action: request.action,
            bridgeUrl: bridgeConfig.url,
            session,
            sessionsCount: response.sessions?.length,
            outputPreview: buildOutputPreview(response.output),
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);
        const fallbackAction =
          params?.action ??
          (isRecord(rawParams) && isTmuxAction(rawParams.action)
            ? rawParams.action
            : "list_sessions");

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "tmux_bridge",
            ok: false,
            action: fallbackAction,
            error: message,
          },
        };
      }
    },
  };
}
