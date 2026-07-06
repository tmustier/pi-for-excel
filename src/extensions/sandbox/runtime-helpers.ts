import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import { Kind, Type, type TSchema } from "@sinclair/typebox";

import type { HttpRequestOptions, LlmCompletionRequest } from "../../commands/extension-api.js";

function isExtensionsSandboxRuntimeHelpersPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

type WidgetPlacement = "above-input" | "below-input";

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
  if (!isExtensionsSandboxRuntimeHelpersPayloadShape(requestRaw)) {
    throw new Error("llm_complete request must be an object.");
  }

  const messagesRaw = requestRaw.messages;
  if (!Array.isArray(messagesRaw)) {
    throw new Error("llm_complete request.messages must be an array.");
  }

  const messages: LlmCompletionRequest["messages"] = [];
  for (const value of messagesRaw) {
    if (!isExtensionsSandboxRuntimeHelpersPayloadShape(value)) {
      throw new Error("llm_complete messages entries must be objects.");
    }

    const role = value.role;
    const content = value.content;
    if ((role !== "user" && role !== "assistant") || typeof content !== "string") {
      throw new Error("llm_complete messages entries must contain role + string content.");
    }

    messages.push({ role, content });
  }

  return {
    model: typeof requestRaw.model === "string" ? requestRaw.model : undefined,
    systemPrompt: typeof requestRaw.systemPrompt === "string" ? requestRaw.systemPrompt : undefined,
    messages,
    maxTokens: typeof requestRaw.maxTokens === "number" ? requestRaw.maxTokens : undefined,
  };
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

  return {
    method: asHttpMethodOrUndefined(optionsRaw.method),
    headers,
    body: typeof optionsRaw.body === "string" ? optionsRaw.body : undefined,
    timeoutMs: typeof optionsRaw.timeoutMs === "number" ? optionsRaw.timeoutMs : undefined,
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
