/**
 * libreoffice_convert — Experimental local LibreOffice bridge adapter.
 *
 * This tool stays registered for a stable tool list/prompt cache,
 * but execution is gated by:
 * - /experimental on python-bridge
 * - /experimental python-bridge-url https://localhost:<port>
 * - reachable bridge health endpoint
 */

import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { Type, type TSchema } from "@sinclair/typebox";

import { getErrorMessage } from "../utils/errors.js";
import { isRecord } from "../utils/type-guards.js";
import { getDefaultPythonBridgeConfig } from "./python-run.js";

const LIBREOFFICE_BRIDGE_API_PATH = "/v1/libreoffice-convert";
const DEFAULT_BRIDGE_TIMEOUT_MS = 45_000;
const MIN_TIMEOUT_MS = 1_000;
const MAX_TIMEOUT_MS = 300_000;

const TARGET_FORMATS = ["csv", "pdf", "xlsx"] as const;
type LibreOfficeTargetFormat = (typeof TARGET_FORMATS)[number];

const TARGET_FORMAT_SET = new Set<string>(TARGET_FORMATS);

function isTargetFormat(value: unknown): value is LibreOfficeTargetFormat {
  return typeof value === "string" && TARGET_FORMAT_SET.has(value);
}

function stringEnum<T extends readonly string[]>(values: T, description: string) {
  return Type.Union(
    values.map((value) => Type.Literal(value)),
    { description },
  );
}

const schema = Type.Object({
  input_path: Type.String({
    minLength: 1,
    description: "Absolute path of the source file to convert.",
  }),
  target_format: stringEnum(TARGET_FORMATS, "Target format: csv, pdf, or xlsx."),
  output_path: Type.Optional(Type.String({
    minLength: 1,
    description: "Optional absolute output path. If omitted, bridge chooses a sibling filename.",
  })),
  overwrite: Type.Optional(Type.Boolean({
    description: "If true, overwrite output_path when it already exists.",
  })),
  timeout_ms: Type.Optional(Type.Integer({
    minimum: MIN_TIMEOUT_MS,
    maximum: MAX_TIMEOUT_MS,
    description: "Optional conversion timeout in ms.",
  })),
});

interface Params {
  input_path: string;
  target_format: LibreOfficeTargetFormat;
  output_path?: string;
  overwrite?: boolean;
  timeout_ms?: number;
}

export interface LibreOfficeBridgeConfig {
  url: string;
  token?: string;
}

export interface LibreOfficeConvertRequest {
  input_path: string;
  target_format: LibreOfficeTargetFormat;
  output_path?: string;
  overwrite?: boolean;
  timeout_ms?: number;
}

export interface LibreOfficeConvertResponse {
  ok: boolean;
  action: "convert";
  input_path?: string;
  target_format?: LibreOfficeTargetFormat;
  output_path?: string;
  bytes?: number;
  converter?: string;
  error?: string;
  metadata?: Record<string, unknown>;
}

export interface LibreOfficeConvertToolDetails {
  kind: "libreoffice_bridge";
  ok: boolean;
  action: "convert";
  bridgeUrl?: string;
  inputPath?: string;
  targetFormat?: LibreOfficeTargetFormat;
  outputPath?: string;
  bytes?: number;
  converter?: string;
  error?: string;
}

export interface LibreOfficeConvertToolDependencies {
  getBridgeConfig?: () => Promise<LibreOfficeBridgeConfig | null>;
  callBridge?: (
    request: LibreOfficeConvertRequest,
    config: LibreOfficeBridgeConfig,
    signal: AbortSignal | undefined,
  ) => Promise<LibreOfficeConvertResponse>;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return isRecord(value) && !Array.isArray(value);
}

function cleanOptionalString(value: string | undefined): string | undefined {
  if (typeof value !== "string") return undefined;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function toOptionalInteger(value: unknown): number | undefined {
  if (typeof value !== "number") return undefined;
  if (!Number.isFinite(value)) return undefined;
  if (!Number.isInteger(value)) return undefined;
  return value;
}

function isAbsolutePath(value: string): boolean {
  if (value.startsWith("/")) return true;
  if (/^[A-Za-z]:[\\/]/.test(value)) return true;
  if (value.startsWith("\\\\")) return true;
  return false;
}

function parseParams(raw: unknown): Params {
  if (!isPlainObject(raw)) {
    throw new Error("Invalid libreoffice_convert params: expected an object.");
  }

  if (typeof raw.input_path !== "string" || raw.input_path.trim().length === 0) {
    throw new Error("input_path is required.");
  }

  if (!isTargetFormat(raw.target_format)) {
    throw new Error("target_format must be one of: csv, pdf, xlsx.");
  }

  const params: Params = {
    input_path: raw.input_path,
    target_format: raw.target_format,
  };

  if (typeof raw.output_path === "string") {
    params.output_path = raw.output_path;
  }

  if (typeof raw.overwrite === "boolean") {
    params.overwrite = raw.overwrite;
  }

  const timeoutMs = toOptionalInteger(raw.timeout_ms);
  if (timeoutMs !== undefined) {
    params.timeout_ms = timeoutMs;
  }

  return params;
}

function validateParams(params: Params): void {
  const inputPath = params.input_path.trim();
  if (!isAbsolutePath(inputPath)) {
    throw new Error("input_path must be an absolute path.");
  }

  if (params.output_path !== undefined) {
    const outputPath = params.output_path.trim();
    if (!isAbsolutePath(outputPath)) {
      throw new Error("output_path must be an absolute path.");
    }
  }

  if (params.timeout_ms !== undefined && (params.timeout_ms < MIN_TIMEOUT_MS || params.timeout_ms > MAX_TIMEOUT_MS)) {
    throw new Error(`timeout_ms must be between ${MIN_TIMEOUT_MS} and ${MAX_TIMEOUT_MS}`);
  }
}

function toBridgeRequest(params: Params): LibreOfficeConvertRequest {
  return {
    input_path: params.input_path.trim(),
    target_format: params.target_format,
    output_path: cleanOptionalString(params.output_path),
    overwrite: params.overwrite,
    timeout_ms: params.timeout_ms,
  };
}

function parseBridgeResponse(value: unknown): LibreOfficeConvertResponse {
  if (!isPlainObject(value)) {
    return {
      ok: true,
      action: "convert",
    };
  }

  const ok = typeof value.ok === "boolean" ? value.ok : true;
  const inputPath = typeof value.input_path === "string" ? value.input_path : undefined;
  const targetFormat = isTargetFormat(value.target_format) ? value.target_format : undefined;
  const outputPath = typeof value.output_path === "string" ? value.output_path : undefined;
  const bytes = typeof value.bytes === "number" ? value.bytes : undefined;
  const converter = typeof value.converter === "string" ? value.converter : undefined;
  const error = typeof value.error === "string" ? value.error : undefined;
  const metadata = isPlainObject(value.metadata) ? value.metadata : undefined;

  return {
    ok,
    action: "convert",
    input_path: inputPath,
    target_format: targetFormat,
    output_path: outputPath,
    bytes,
    converter,
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

async function defaultGetBridgeConfig(): Promise<LibreOfficeBridgeConfig | null> {
  return getDefaultPythonBridgeConfig();
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

function computeFetchTimeoutMs(request: LibreOfficeConvertRequest): number {
  const requested = request.timeout_ms;
  if (typeof requested !== "number") return DEFAULT_BRIDGE_TIMEOUT_MS;
  return Math.min(requested + 2_000, MAX_TIMEOUT_MS + 2_000);
}

async function defaultCallBridge(
  request: LibreOfficeConvertRequest,
  config: LibreOfficeBridgeConfig,
  signal: AbortSignal | undefined,
): Promise<LibreOfficeConvertResponse> {
  const endpoint = joinBridgeUrl(config.url, LIBREOFFICE_BRIDGE_API_PATH);
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
      throw new Error(`LibreOffice bridge request failed (${response.status}): ${reason}`);
    }

    if (parsedBody === null) {
      return {
        ok: true,
        action: "convert",
      };
    }

    const parsed = parseBridgeResponse(parsedBody);
    if (!parsed.ok) {
      throw new Error(parsed.error ?? "LibreOffice bridge rejected the request.");
    }

    return parsed;
  } catch (error: unknown) {
    if (isAbortError(error)) {
      if (signal?.aborted) {
        throw new Error("Aborted");
      }
      throw new Error(`LibreOffice bridge request timed out after ${timeoutMs}ms.`);
    }

    throw error;
  } finally {
    clearTimeout(timeoutId);
    if (signal) {
      signal.removeEventListener("abort", abortFromCaller);
    }
  }
}

function formatBridgeSuccessText(response: LibreOfficeConvertResponse): string {
  const inputPath = response.input_path ?? "(input)";
  const outputPath = response.output_path ?? "(output)";
  const target = response.target_format ?? "file";

  const sections: string[] = [
    `Converted \`${inputPath}\` to **${target.toUpperCase()}** at \`${outputPath}\`.`,
  ];

  if (typeof response.bytes === "number") {
    sections.push(`Output size: ${response.bytes} bytes.`);
  }

  if (response.converter) {
    sections.push(`Converter: ${response.converter}.`);
  }

  sections.push("You can now read or import the converted file using your local workflow.");

  return sections.join("\n\n");
}

function buildMissingBridgeConfigurationMessage(): string {
  return (
    "Python/LibreOffice bridge URL is not configured. " +
    "Run /experimental python-bridge-url https://localhost:3340 and enable /experimental on python-bridge."
  );
}

export function createLibreOfficeConvertTool(
  dependencies: LibreOfficeConvertToolDependencies = {},
): AgentTool<TSchema, LibreOfficeConvertToolDetails> {
  const getBridgeConfig = dependencies.getBridgeConfig ?? defaultGetBridgeConfig;
  const callBridge = dependencies.callBridge ?? defaultCallBridge;

  return {
    name: "libreoffice_convert",
    label: "LibreOffice Convert",
    description:
      "Convert spreadsheet files via a local LibreOffice bridge (csv/pdf/xlsx). " +
      "Useful for offline exports and format transformations.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      rawParams: unknown,
      signal: AbortSignal | undefined,
    ): Promise<AgentToolResult<LibreOfficeConvertToolDetails>> => {
      try {
        const params = parseParams(rawParams);
        validateParams(params);

        const bridgeConfig = await getBridgeConfig();
        if (!bridgeConfig) {
          return {
            content: [{ type: "text", text: buildMissingBridgeConfigurationMessage() }],
            details: {
              kind: "libreoffice_bridge",
              ok: false,
              action: "convert",
              error: "missing_bridge_url",
            },
          };
        }

        const request = toBridgeRequest(params);
        const response = await callBridge(request, bridgeConfig, signal);

        if (!response.ok) {
          throw new Error(response.error ?? "LibreOffice bridge rejected the request.");
        }

        return {
          content: [{ type: "text", text: formatBridgeSuccessText(response) }],
          details: {
            kind: "libreoffice_bridge",
            ok: true,
            action: "convert",
            bridgeUrl: bridgeConfig.url,
            inputPath: response.input_path ?? request.input_path,
            targetFormat: response.target_format ?? request.target_format,
            outputPath: response.output_path,
            bytes: response.bytes,
            converter: response.converter,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "libreoffice_bridge",
            ok: false,
            action: "convert",
            error: message,
          },
        };
      }
    },
  };
}
