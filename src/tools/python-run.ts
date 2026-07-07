function isToolsPythonRunPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * python_run — Execute Python via native bridge (preferred) or Pyodide fallback.
 *
 * This tool stays registered for a stable tool list/prompt cache.
 * If no native bridge is reachable (configured override or default localhost URL),
 * it can run in-browser via Pyodide.
 */

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { SettingsStore } from "../storage/local/settings-store.js";
import { Type, type Static, type TSchema } from "@sinclair/typebox";

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";
import { getErrorMessage } from "../utils/errors.js";
import {
  extractBridgeErrorMessage,
  isAbortError,
  joinBridgeUrl,
  tryParseBridgeJson,
} from "./bridge-http-utils.js";
import {
  DEFAULT_PYTHON_BRIDGE_URL,
  PYTHON_BRIDGE_URL_SETTING_KEY,
} from "./experimental-tool-gates.js";

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
  source?: "configured" | "default";
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
  metadata?: DynamicObject;
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
  gateReason?: "missing_bridge_url" | "invalid_bridge_url" | "bridge_unreachable";
  skillHint?: string;
}

export interface PythonRunToolDependencies {
  getBridgeConfig?: () => Promise<PythonBridgeConfig | null>;
  callBridge?: (
    request: PythonBridgeRequest,
    config: PythonBridgeConfig,
    signal: AbortSignal | undefined,
  ) => Promise<PythonBridgeResponse>;
  /** Override Pyodide availability check (for testing). */
  isPyodideAvailable?: () => boolean;
  /** Override Pyodide runtime call (for testing). */
  callPyodide?: (
    request: PythonBridgeRequest,
    signal: AbortSignal | undefined,
  ) => Promise<PythonBridgeResponse>;
}

function isPythonBridgePayloadShape(value: DynamicValue): value is DynamicObject {
  return isToolsPythonRunPayloadShape(value) && !Array.isArray(value);
}

function cleanOptionalString(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? value : undefined;
}

function toOptionalInteger(value: DynamicValue): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  return value;
}

function parseParams(raw: DynamicValue): Params {
  if (!isPythonBridgePayloadShape(raw)) {
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
  const inputJson = cleanOptionalString(params.input_json);
  return {
    code: params.code,
    ...(inputJson !== undefined ? { input_json: inputJson } : {}),
    ...(params.timeout_ms !== undefined ? { timeout_ms: params.timeout_ms } : {}),
  };
}

function parseBridgeResponse(value: DynamicValue): PythonBridgeResponse {
  if (!isPythonBridgePayloadShape(value)) {
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
  const metadata = isPythonBridgePayloadShape(value.metadata) ? value.metadata : undefined;

  return {
    ok,
    action: "run_python",
    ...(exitCode !== undefined ? { exit_code: exitCode } : {}),
    ...(stdout !== undefined ? { stdout } : {}),
    ...(stderr !== undefined ? { stderr } : {}),
    ...(resultJson !== undefined ? { result_json: resultJson } : {}),
    ...(truncated !== undefined ? { truncated } : {}),
    ...(error !== undefined ? { error } : {}),
    ...(metadata !== undefined ? { metadata } : {}),
  };
}

async function getSettingsStore(): Promise<SettingsStore> {
  const storageModule = await import("../storage/local/app-storage.js");
  return storageModule.getAppStorage().settings;
}

export async function getDefaultPythonBridgeConfig(): Promise<PythonBridgeConfig | null> {
  let rawUrl = DEFAULT_PYTHON_BRIDGE_URL;
  let source: "configured" | "default" = "default";
  let token: string | undefined;

  try {
    const settings = await getSettingsStore();

    const urlValue = await settings.get<string>(PYTHON_BRIDGE_URL_SETTING_KEY);
    const configuredUrl = typeof urlValue === "string" ? urlValue.trim() : "";
    if (configuredUrl.length > 0) {
      rawUrl = configuredUrl;
      source = "configured";
    }

    const tokenValue = await settings.get<string>(PYTHON_BRIDGE_TOKEN_SETTING_KEY);
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
      ...(token !== undefined ? { token } : {}),
      source,
    };
  } catch {
    return null;
  }
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
    const parsedBody = tryParseBridgeJson(rawBody);

    if (!response.ok) {
      const payloadError = extractBridgeErrorMessage(parsedBody);
      const textError = rawBody.trim().length > 0 ? rawBody.trim() : null;
      const reason = payloadError ?? textError ?? `HTTP ${response.status}`;
      throw new Error(`Python bridge request failed (${response.status}): ${reason}`);
    }

    if (parsedBody === null) {
      const stdout = rawBody.trim().length > 0 ? rawBody : undefined;
      return {
        ok: true,
        action: "run_python",
        ...(stdout !== undefined ? { stdout } : {}),
      };
    }

    const parsed = parseBridgeResponse(parsedBody);
    if (!parsed.ok) {
      throw new Error(parsed.error ?? "Python bridge rejected the request.");
    }

    return parsed;
  } catch (error) {
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

function buildNoPythonAvailableMessage(): string {
  return (
    "Python is unavailable in this environment. " +
    "The current browser/WebView does not support WebAssembly Workers (needed for in-browser Pyodide). " +
    "Power users can configure a native Python bridge in Settings → Experimental."
  );
}

function withSkillHintLine(message: string, skillName: string): string {
  return `${message}\nSkill: ${skillName}`;
}

function shouldAttachPythonBridgeSkillHint(message: string): boolean {
  const normalized = message.toLowerCase();

  return normalized.includes("python bridge")
    || normalized.includes("python-bridge-url")
    || normalized.includes("bridge url")
    || normalized.includes("no_python_runtime")
    || normalized.includes("webassembly workers")
    || normalized.includes("bridge unavailable")
    || normalized.includes("bridge request")
    || normalized.includes("failed to fetch")
    || normalized.includes("fetch failed")
    || normalized.includes("network request failed")
    || normalized.includes("econnrefused");
}

export function shouldFallbackToPyodideAfterBridgeError(
  error: DynamicValue,
  bridgeConfig: PythonBridgeConfig,
): boolean {
  if (bridgeConfig.source !== "default") {
    return false;
  }

  const message = getErrorMessage(error).toLowerCase();

  return (
    message.includes("timed out")
    || message.includes("failed to fetch")
    || message.includes("fetch failed")
    || message.includes("networkerror")
    || message.includes("network request failed")
    || message.includes("ecconnrefused")
    || message.includes("econnrefused")
    || message.includes("econnreset")
    || message.includes("enotfound")
  );
}

async function getDefaultPyodideAvailable(): Promise<boolean> {
  try {
    const { isPyodideAvailable: check } = await import("../python/pyodide-runtime.js");
    return check();
  } catch {
    return false;
  }
}

async function getDefaultCallPyodide(
  request: PythonBridgeRequest,
  signal: AbortSignal | undefined,
): Promise<PythonBridgeResponse> {
  const { callPyodideRuntime } = await import("../python/pyodide-runtime.js");
  return callPyodideRuntime(request, signal ?? undefined);
}

export function createPythonRunTool(
  dependencies: PythonRunToolDependencies = {},
): AgentTool<TSchema, PythonRunToolDetails> {
  const getBridgeConfig = dependencies.getBridgeConfig ?? getDefaultPythonBridgeConfig;
  const callBridge = dependencies.callBridge ?? callDefaultPythonBridge;
  const checkPyodide = dependencies.isPyodideAvailable;
  const pyodideCall = dependencies.callPyodide;

  return {
    name: "python_run",
    label: "Python Run",
    description:
      "Run Python code in-browser via Pyodide (no setup needed). " +
      "Standard library and pure-Python packages (numpy, pandas, etc.) work automatically. " +
      "If a local Python bridge is configured (or running on default localhost URL), uses local Python instead. " +
      "Pass optional input_json, inspect stdout/stderr, and chain results into write_cells.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: DynamicValue,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<PythonRunToolDetails>> => {
      let params: Params | null = null;

      try {
        params = parseParams(rawParams);
        validateParams(params);

        const bridgeConfig = await getBridgeConfig();
        const request = toBridgeRequest(params);

        // Prefer native bridge when configured or available at default localhost URL.
        if (bridgeConfig) {
          try {
            const response = await callBridge(request, bridgeConfig, signal);

            if (!response.ok) {
              throw new Error(response.error ?? "Python bridge rejected the request.");
            }

            const stdoutPreview = buildOutputPreview(response.stdout);
            const stderrPreview = buildOutputPreview(response.stderr);
            const resultPreview = buildOutputPreview(response.result_json);
            return {
              content: [{ type: "text", text: formatBridgeSuccessText(response) }],
              details: {
                kind: "python_bridge",
                ok: true,
                action: "run_python",
                bridgeUrl: bridgeConfig.url,
                ...(response.exit_code !== undefined ? { exitCode: response.exit_code } : {}),
                ...(stdoutPreview !== undefined ? { stdoutPreview } : {}),
                ...(stderrPreview !== undefined ? { stderrPreview } : {}),
                ...(resultPreview !== undefined ? { resultPreview } : {}),
                ...(response.truncated !== undefined ? { truncated: response.truncated } : {}),
              },
            };
          } catch (error) {
            if (!shouldFallbackToPyodideAfterBridgeError(error, bridgeConfig)) {
              throw error;
            }
          }
        }

        // Fall back to in-browser Pyodide
        const pyodideAvailable = checkPyodide
          ? checkPyodide()
          : await getDefaultPyodideAvailable();

        if (!pyodideAvailable) {
          return {
            content: [{
              type: "text",
              text: withSkillHintLine(buildNoPythonAvailableMessage(), "python-bridge"),
            }],
            details: {
              kind: "python_bridge",
              ok: false,
              action: "run_python",
              error: "no_python_runtime",
              skillHint: "python-bridge",
            },
          };
        }

        const callPyodideFn = pyodideCall ?? getDefaultCallPyodide;
        const response = await callPyodideFn(request, signal);

        if (!response.ok) {
          throw new Error(response.error ?? "Pyodide execution failed.");
        }

        const stdoutPreview = buildOutputPreview(response.stdout);
        const stderrPreview = buildOutputPreview(response.stderr);
        const resultPreview = buildOutputPreview(response.result_json);
        return {
          content: [{ type: "text", text: formatBridgeSuccessText(response) }],
          details: {
            kind: "python_bridge",
            ok: true,
            action: "run_python",
            ...(response.exit_code !== undefined ? { exitCode: response.exit_code } : {}),
            ...(stdoutPreview !== undefined ? { stdoutPreview } : {}),
            ...(stderrPreview !== undefined ? { stderrPreview } : {}),
            ...(resultPreview !== undefined ? { resultPreview } : {}),
            ...(response.truncated !== undefined ? { truncated: response.truncated } : {}),
          },
        };
      } catch (error) {
        const message = getErrorMessage(error);
        const skillHint = shouldAttachPythonBridgeSkillHint(message)
          ? "python-bridge"
          : undefined;

        return {
          content: [{
            type: "text",
            text: skillHint
              ? `Error: ${withSkillHintLine(message, skillHint)}`
              : `Error: ${message}`,
          }],
          details: {
            kind: "python_bridge",
            ok: false,
            action: "run_python",
            error: message,
            ...(skillHint !== undefined ? { skillHint } : {}),
          },
        };
      }
    },
  };
}
