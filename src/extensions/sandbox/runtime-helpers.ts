import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import { Kind, Type, type Static, type TSchema } from "@sinclair/typebox";
import { Value } from "@sinclair/typebox/value";

import type {
  HttpRequestOptions,
  LlmCompletionRequest,
  LlmCompletionResult,
} from "../../commands/extension-api.js";

function isExtensionsSandboxRuntimeHelpersPayloadShape(value: DynamicValue): value is DynamicObject {
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

export function getErrorMessage(error: DynamicValue): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }

  return String(error);
}

export function sanitizeText(value: DynamicValue): string {
  if (typeof value !== "string") {
    return "";
  }

  return value;
}

export function asNonEmptyString(value: DynamicValue, field: string): string {
  if (typeof value !== "string" || value.trim().length === 0) {
    throw new Error(`${field} must be a non-empty string.`);
  }

  return value.trim();
}

export function asSandboxPayload(value: DynamicValue, field: string): DynamicObject {
  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(value)) {
    throw new Error(`${field} must be an object.`);
  }

  return value;
}

export function asFiniteNumberOrNull(value: DynamicValue): number | null {
  if (typeof value !== "number" || Number.isNaN(value) || !Number.isFinite(value)) {
    return null;
  }

  return value;
}

export function asFiniteNumberOrNullOrUndefined(value: DynamicValue): number | null | undefined {
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

export function asWidgetPlacementOrUndefined(value: DynamicValue): WidgetPlacement | undefined {
  if (value === "above-input" || value === "below-input") {
    return value;
  }

  return undefined;
}

export function asBooleanOrUndefined(value: DynamicValue): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

export function parseSandboxLlmCompletionRequest(requestRaw: DynamicValue): LlmCompletionRequest {
  if (!Value.Check(sandboxLlmCompletionRequestSchema, requestRaw)) {
    const error = Value.Errors(sandboxLlmCompletionRequestSchema, requestRaw).First();
    const field = error?.path ? error.path.replaceAll("/", ".") : "";
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
  paramsRaw: DynamicValue,
  complete: (request: LlmCompletionRequest) => Promise<LlmCompletionResult>,
): Promise<LlmCompletionResult> {
  const payload = asSandboxPayload(paramsRaw, "llm_complete params");
  return complete(parseSandboxLlmCompletionRequest(payload.request));
}

function asHttpMethodOrUndefined(value: DynamicValue): HttpRequestOptions["method"] | undefined {
  return value === "GET"
    || value === "POST"
    || value === "PUT"
    || value === "PATCH"
    || value === "DELETE"
    || value === "HEAD"
    ? value
    : undefined;
}

export function parseSandboxHttpRequestOptions(optionsRaw: DynamicValue): HttpRequestOptions | undefined {
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

function isTypeBoxSchema(value: DynamicValue): value is TSchema {
  return isExtensionsSandboxRuntimeHelpersPayloadShape(value) && Kind in value;
}

export function normalizeSandboxToolParameters(raw: DynamicValue): TSchema {
  if (isTypeBoxSchema(raw)) {
    return raw;
  }

  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(raw)) {
    throw new Error("register_tool parameters must be an object schema.");
  }

  return Type.Unsafe<DynamicValue>(raw);
}

export function normalizeSandboxToolResult(raw: DynamicValue): AgentToolResult<DynamicValue> {
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
