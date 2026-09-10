import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import { IsSchema, Type, type Static, type TSchema } from "typebox";
import { Value } from "typebox/value";

import type {
  HttpRequestOptions,
  LlmCompletionRequest,
  LlmCompletionResult,
} from "../../commands/extension-api.js";

function isExtensionsSandboxRuntimeHelpersPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type WidgetPlacement = "above-input" | "below-input";

const sandboxLlmCompletionRequestSchema = Type.Object({
  model: Type.Optional(Type.String()),
  systemPrompt: Type.Optional(Type.String()),
  messages: Type.Array(Type.Object({
    role: Type.Union([Type.Literal("user"), Type.Literal("assistant")]),
    content: Type.String(),
  }, { additionalProperties: true })),
  maxTokens: Type.Optional(Type.Integer({ minimum: 1 })),
}, { additionalProperties: true });

type SandboxLlmCompletionRequestDto = Static<typeof sandboxLlmCompletionRequestSchema>;

export function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}

export function sanitizeText(value: unknown): string {
  if (typeof value !== "string") {
    return "";
  }

  return value;
}

export function asNonEmptyString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${field} must be a non-empty string.`);
  }

  return value.trim();
}

export function asSandboxPayload(value: unknown, field: string): Record<string, unknown> {
  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(value)) {
    throw new Error(`${field} must be an object.`);
  }

  return value;
}

export function asFiniteNumberOrNull(value: unknown): number | null {
  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return null;
  }

  return value;
}

export function asFiniteNumberOrNullOrUndefined(value: unknown): number | null | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (value === null) {
    return null;
  }

  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return undefined;
  }

  return value;
}

export function asWidgetPlacementOrUndefined(value: unknown): WidgetPlacement | undefined {
  if (value === "above-input" || value === "below-input") {
    return value;
  }

  return undefined;
}

export function asBooleanOrUndefined(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

export function parseSandboxLlmCompletionRequest(requestRaw: unknown): LlmCompletionRequest {
  if (!Value.Check(sandboxLlmCompletionRequestSchema, requestRaw)) {
    const error = Value.Errors(sandboxLlmCompletionRequestSchema, requestRaw)[0];
    const field = error?.instancePath ? error.instancePath.replaceAll("/", ".") : "";
    const reason = error?.message ?? "request does not match the expected schema";
    throw new Error(`llm_complete request${field} is invalid: ${reason}.`);
  }

  const request: SandboxLlmCompletionRequestDto = requestRaw;
  return {
    ...(request.model !== undefined ? { model: request.model } : {}),
    ...(request.systemPrompt !== undefined ? { systemPrompt: request.systemPrompt } : {}),
    messages: request.messages.map(({ role, content }) => ({ role, content })),
    ...(request.maxTokens !== undefined ? { maxTokens: request.maxTokens } : {}),
  };
}

export async function dispatchSandboxLlmCompletion(
  paramsRaw: unknown,
  complete: (request: LlmCompletionRequest) => Promise<LlmCompletionResult>,
): Promise<LlmCompletionResult> {
  const payload = asSandboxPayload(paramsRaw, "llm_complete params");
  return complete(parseSandboxLlmCompletionRequest(payload.request));
}

function asHttpMethodOrUndefined(value: unknown): HttpRequestOptions["method"] | undefined {
  return value === "GET"
    || value === "POST"
    || value === "PUT"
    || value === "PATCH"
    || value === "DELETE"
    || value === "HEAD"
    ? value
    : undefined;
}

export function parseSandboxHttpRequestOptions(optionsRaw: unknown): HttpRequestOptions | undefined {
  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(optionsRaw)) {
    return undefined;
  }

  const headersRaw = optionsRaw.headers;
  let headers: Record<string, string> | undefined;
  if (isExtensionsSandboxRuntimeHelpersPayloadShape(headersRaw)) {
    headers = {};
    for (const [key, value] of Object.entries(headersRaw)) {
      if (typeof value === "string") {
        headers[key] = value;
      }
    }
  }

  const normalizedConnection = typeof optionsRaw.connection === "string"
    ? optionsRaw.connection.trim()
    : "";

  const method = asHttpMethodOrUndefined(optionsRaw.method);
  const body = typeof optionsRaw.body === "string" ? optionsRaw.body : undefined;
  const timeoutMs = typeof optionsRaw.timeoutMs === "number" ? optionsRaw.timeoutMs : undefined;

  return {
    ...(method !== undefined ? { method } : {}),
    ...(headers !== undefined ? { headers } : {}),
    ...(body !== undefined ? { body } : {}),
    ...(timeoutMs !== undefined ? { timeoutMs } : {}),
    ...(normalizedConnection.length > 0 ? { connection: normalizedConnection } : {}),
  };
}

function isTypeBoxSchema(value: unknown): value is TSchema {
  return IsSchema(value);
}

export function normalizeSandboxToolParameters(raw: unknown): TSchema {
  if (isTypeBoxSchema(raw)) {
    return raw;
  }

  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(raw)) {
    throw new Error("register_tool parameters must be an object schema.");
  }

  return Type.Unsafe<unknown>(raw);
}

export function normalizeSandboxToolResult(raw: unknown): AgentToolResult<unknown> {
  const content: Array<{ type: "text"; text: string }> = [];

  if (isExtensionsSandboxRuntimeHelpersPayloadShape(raw) && Array.isArray(raw.content)) {
    for (const item of raw.content) {
      if (!isExtensionsSandboxRuntimeHelpersPayloadShape(item)) {
        continue;
      }

      if (item.type !== "text") {
        continue;
      }

      if (typeof item.text !== "string") {
        continue;
      }

      content.push({
        type: "text",
        text: item.text,
      });
    }
  }

  if (content.length === 0) {
    const fallbackText = isExtensionsSandboxRuntimeHelpersPayloadShape(raw) && Array.isArray(raw.content)
      ? "Sandbox tool returned non-text content; showing serialized payload instead."
      : "Sandbox tool returned an invalid payload; showing serialized payload instead.";

    content.push({
      type: "text",
      text: `${fallbackText}\n\n\`\`\`json\n${JSON.stringify(raw, null, 2)}\n\`\`\``,
    });
  }

  const details = isExtensionsSandboxRuntimeHelpersPayloadShape(raw) && Object.prototype.hasOwnProperty.call(raw, "details")
    ? raw.details
    : undefined;

  return {
    content,
    details,
  };
}
