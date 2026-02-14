/**
 * python_run — Experimental local Python bridge adapter.
 *
 * This tool stays registered for a stable tool list/prompt cache.
 * Execution requires a configured and reachable bridge URL
 * (/experimental python-bridge-url https://localhost:<port>).
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";
import { getErrorMessage } from "../utils/errors.js";
import { isRecord } from "../utils/type-guards.js";
import { PYTHON_BRIDGE_URL_SETTING_KEY } from "./experimental-tool-gates.js";

const PYTHON_BRIDGE_API_PATH = "/v1/python-run";
const DEFAULT_BRIDGE_TIMEOUT_MS = 20_000;
const MIN_TIMEOUT_MS = 100;
const MAX_TIMEOUT_MS = 120_000;

export const PYTHON_BRIDGE_TOKEN_SETTING_KEY = "python.bridge.token";

const schema = Type.Object({
  code: Type.String({
    minLength: 1,
    maxLength: 40_000,
    description:
      "Python code to execute. Optionally assign JSON-serializable data to `result` to return it explicitly.",
  }),
  input_json: Type.Optional(Type.String({
    maxLength: 200_000,
    description:
      "Optional JSON string exposed inside Python as `input_data`.",
  })),
  timeout_ms: Type.Optional(Type.Integer({
    minimum: MIN_TIMEOUT_MS,
    maximum: MAX_TIMEOUT_MS,
    description: "Optional execution timeout in ms.",
  })),
});

type Params = Static<typeof schema>;

export interface PythonBridgeConfig {
  url: string;
  token?: string;
}

export interface PythonBridgeRequest {
  code: string;
  input_json?: string;
  timeout_ms?: number;
}

export interface PythonBridgeResponse {
  ok: boolean;
  action: "run_python";
  exit_code?: number;
  stdout?: string;
  stderr?: string;
  result_json?: string;
  truncated?: boolean;
  error?: string;
  metadata?: Record<string, unknown>;
}

export interface PythonRunToolDetails {
  kind: "python_bridge";
  ok: boolean;
  action: "run_python";
  bridgeUrl?: string;
  exitCode?: number;
  stdoutPreview?: string;
  stderrPreview?: string;
  resultPreview?: string;
  truncated?: boolean;
  error?: string;
}

export interface PythonRunToolDependencies {
  getBridgeConfig?: () => Promise<PythonBridgeConfig | null>;
  callBridge?: (
    request: PythonBridgeRequest,
    config: PythonBridgeConfig,
    signal: AbortSignal | undefined,
  ) => Promise<PythonBridgeResponse>;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return isRecord(value) && !Array.isArray(value);
}

function cleanOptionalString(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? value : undefined;
}

function toOptionalInteger(value: unknown): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  return value;
}

function parseParams(raw: unknown): Params {
  if (!isPlainObject(raw)) {
    throw new Error("Invalid python_run params: expected an object.");
  }

  if (typeof raw.code !== "string" || raw.code.trim().length === 0) {
    throw new Error("code is required.");
  }

  const params: Params = {
    code: raw.code,
  };

  if (typeof raw.input_json === "string") {
    params.input_json = raw.input_json;
  }

  const timeoutMs = toOptionalInteger(raw.timeout_ms);
  if (timeoutMs !== undefined) {
    params.timeout_ms = timeoutMs;
  }

  return params;
}

function validateParams(params: Params): void {
  if (params.timeout_ms !== undefined && (params.timeout_ms < MIN_TIMEOUT_MS || params.timeout_ms > MAX_TIMEOUT_MS)) {
    throw new Error(`timeout_ms must be between ${MIN_TIMEOUT_MS} and ${MAX_TIMEOUT_MS}`);
  }

  if (params.input_json !== undefined) {
    try {
      void JSON.parse(params.input_json);
    } catch {
      throw new Error("input_json must be valid JSON.");
    }
  }
}

function toBridgeRequest(params: Params): PythonBridgeRequest {
  return {
    code: params.code,
    input_json: cleanOptionalString(params.input_json),
    timeout_ms: params.timeout_ms,
  };
}

function parseBridgeResponse(value: unknown): PythonBridgeResponse {
  if (!isPlainObject(value)) {
    return {
      ok: true,
      action: "run_python",
    };
  }

  const ok = typeof value.ok === "boolean" ? value.ok : true;
  const exitCode = typeof value.exit_code === "number" ? value.exit_code : undefined;
  const stdout = typeof value.stdout === "string" ? value.stdout : undefined;
  const stderr = typeof value.stderr === "string" ? value.stderr : undefined;
  const resultJson = typeof value.result_json === "string" ? value.result_json : undefined;
  const truncated = typeof value.truncated === "boolean" ? value.truncated : undefined;
  const error = typeof value.error === "string" ? value.error : undefined;
  const metadata = isPlainObject(value.metadata) ? value.metadata : undefined;

  return {
    ok,
    action: "run_python",
    exit_code: exitCode,
    stdout,
    stderr,
    result_json: resultJson,
    truncated,
    error,
    metadata,
  };
}

function tryParseJson(text: string): unknown {
  const trimmed = text.trim();
  if (trimmed.length === 0) return null;

  try {
    return JSON.parse(trimmed);
  } catch {
    return null;
  }
}

function extractBridgeErrorMessage(value: unknown): string | null {
  if (isPlainObject(value) && typeof value.error === "string") {
    return value.error;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : null;
  }

  return null;
}

function joinBridgeUrl(baseUrl: string, path: string): string {
  return `${baseUrl.replace(/\/+$/, "")}${path}`;
}

async function getSettingsStore() {
  const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
  return storageModule.getAppStorage().settings;
}

export async function getDefaultPythonBridgeConfig(): Promise<PythonBridgeConfig | null> {
  try {
    const settings = await getSettingsStore();

    const urlValue = await settings.get<string>(PYTHON_BRIDGE_URL_SETTING_KEY);
    const rawUrl = typeof urlValue === "string" ? urlValue.trim() : "";
    if (rawUrl.length === 0) return null;

    const normalizedUrl = validateOfficeProxyUrl(rawUrl);

    const tokenValue = await settings.get<string>(PYTHON_BRIDGE_TOKEN_SETTING_KEY);
    const token = typeof tokenValue === "string" && tokenValue.trim().length > 0
      ? tokenValue.trim()
      : undefined;

    return {
      url: normalizedUrl,
      token,
    };
  } catch {
    return null;
  }
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

function computeFetchTimeoutMs(request: PythonBridgeRequest): number {
  const requested = request.timeout_ms;
  if (typeof requested !== "number") return DEFAULT_BRIDGE_TIMEOUT_MS;
  return Math.min(requested + 2_000, MAX_TIMEOUT_MS + 2_000);
}

export async function callDefaultPythonBridge(
  request: PythonBridgeRequest,
  config: PythonBridgeConfig,
  signal: AbortSignal | undefined,
): Promise<PythonBridgeResponse> {
  const endpoint = joinBridgeUrl(config.url, PYTHON_BRIDGE_API_PATH);
  const controller = new AbortController();
  const timeoutMs = computeFetchTimeoutMs(request);

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
    const parsedBody = tryParseJson(rawBody);

    if (!response.ok) {
      const payloadError = extractBridgeErrorMessage(parsedBody);
      const textError = rawBody.trim().length > 0 ? rawBody.trim() : null;
      const reason = payloadError ?? textError ?? `HTTP ${response.status}`;
      throw new Error(`Python bridge request failed (${response.status}): ${reason}`);
    }

    if (parsedBody === null) {
      return {
        ok: true,
        action: "run_python",
        stdout: rawBody.trim().length > 0 ? rawBody : undefined,
      };
    }

    const parsed = parseBridgeResponse(parsedBody);
    if (!parsed.ok) {
      throw new Error(parsed.error ?? "Python bridge rejected the request.");
    }

    return parsed;
  } catch (error: unknown) {
    if (isAbortError(error)) {
      if (signal?.aborted) {
        throw new Error("Aborted");
      }
      throw new Error(`Python bridge request timed out after ${timeoutMs}ms.`);
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
    if (signal) {
      signal.removeEventListener("abort", abortFromCaller);
    }
  }
}

function renderCodeBlock(text: string, language?: string): string {
  const trimmed = text.trim();
  if (trimmed.length === 0) return "";

  const header = language ? `\`\`\`${language}` : "\`\`\`";
  return `${header}\n${trimmed}\n\`\`\``;
}

function formatBridgeSuccessText(response: PythonBridgeResponse): string {
  const exitCode = response.exit_code ?? 0;
  const sections: string[] = [`Ran Python snippet (exit code ${exitCode}).`];

  if (response.result_json && response.result_json.trim().length > 0) {
    sections.push(`Result JSON:\n\n${renderCodeBlock(response.result_json, "json")}`);
  }

  if (response.stdout && response.stdout.trim().length > 0) {
    sections.push(`Stdout:\n\n${renderCodeBlock(response.stdout)}`);
  }

  if (response.stderr && response.stderr.trim().length > 0) {
    sections.push(`Stderr:\n\n${renderCodeBlock(response.stderr)}`);
  }

  if (sections.length === 1) {
    sections.push("No output was produced.");
  }

  if (response.truncated) {
    sections.push("⚠️ Output was truncated by bridge limits.");
  }

  sections.push(
    "If you want these results in Excel, call write_cells with the target range and values.",
  );

  return sections.join("\n\n");
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
    "Python bridge URL is not configured. " +
    "Run /experimental python-bridge-url https://localhost:3340 to set it up."
  );
}

export function createPythonRunTool(
  dependencies: PythonRunToolDependencies = {},
): AgentTool<TSchema, PythonRunToolDetails> {
  const getBridgeConfig = dependencies.getBridgeConfig ?? getDefaultPythonBridgeConfig;
  const callBridge = dependencies.callBridge ?? callDefaultPythonBridge;

  return {
    name: "python_run",
    label: "Python Run",
    description:
      "Run Python code through a local bridge. " +
      "Pass optional input_json, inspect stdout/stderr, and chain results into write_cells.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<PythonRunToolDetails>> => {
      let params: Params | null = null;

      try {
        params = parseParams(rawParams);
        validateParams(params);

        const bridgeConfig = await getBridgeConfig();
        if (!bridgeConfig) {
          return {
            content: [{ type: "text", text: buildMissingBridgeConfigurationMessage() }],
            details: {
              kind: "python_bridge",
              ok: false,
              action: "run_python",
              error: "missing_bridge_url",
            },
          };
        }

        const request = toBridgeRequest(params);
        const response = await callBridge(request, bridgeConfig, signal);

        if (!response.ok) {
          throw new Error(response.error ?? "Python bridge rejected the request.");
        }

        return {
          content: [{ type: "text", text: formatBridgeSuccessText(response) }],
          details: {
            kind: "python_bridge",
            ok: true,
            action: "run_python",
            bridgeUrl: bridgeConfig.url,
            exitCode: response.exit_code,
            stdoutPreview: buildOutputPreview(response.stdout),
            stderrPreview: buildOutputPreview(response.stderr),
            resultPreview: buildOutputPreview(response.result_json),
            truncated: response.truncated,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "python_bridge",
            ok: false,
            action: "run_python",
            error: message,
          },
        };
      }
    },
  };
}
