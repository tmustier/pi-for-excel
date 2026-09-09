import assert from "node:assert/strict";
import { test } from "node:test";

import { Kind, Type } from "@sinclair/typebox";

import { normalizeSandboxToolParameters } from "../src/extensions/sandbox-runtime.ts";
import {
  dispatchSandboxLlmCompletion,
  parseSandboxHttpRequestOptions,
  parseSandboxLlmCompletionRequest,
} from "../src/extensions/sandbox/runtime-helpers.ts";
function isSandboxRuntimeSchemaTestPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


void test("normalizeSandboxToolParameters keeps TypeBox schema unchanged", () => {
  const schema = Type.Object({
    text: Type.String(),
  });

  const normalized = normalizeSandboxToolParameters(schema);
  assert.equal(normalized, schema);
});

void test("normalizeSandboxToolParameters accepts plain JSON schema objects", () => {
  const rawSchema = {
    type: "object",
    properties: {
      text: {
        type: "string",
      },
    },
    required: ["text"],
    additionalProperties: false,
  };

  const normalized = normalizeSandboxToolParameters(rawSchema);

  assert.ok(isSandboxRuntimeSchemaTestPayloadShape(normalized));
  assert.equal(normalized.type, "object");
  assert.ok(Array.isArray(normalized.required));
  assert.equal(normalized.additionalProperties, false);
  assert.ok(Kind in normalized);
});

void test("normalizeSandboxToolParameters rejects non-object schema values", () => {
  assert.throws(
    () => {
      normalizeSandboxToolParameters("type:string");
    },
    /object schema/i,
  );
});

void test("parseSandboxLlmCompletionRequest validates and normalizes request payload", () => {
  const request = parseSandboxLlmCompletionRequest({
    model: "openai/gpt-5-mini",
    systemPrompt: "system",
    maxTokens: 123,
    messages: [
      { role: "user", content: "hello", extensionMetadata: true },
      { role: "assistant", content: "hi" },
    ],
    extensionMetadata: "preserved as accepted input, omitted from the service DTO",
  });

  assert.deepEqual(request, {
    model: "openai/gpt-5-mini",
    systemPrompt: "system",
    maxTokens: 123,
    messages: [
      { role: "user", content: "hello" },
      { role: "assistant", content: "hi" },
    ],
  });

  assert.throws(
    () => {
      parseSandboxLlmCompletionRequest({ messages: "bad" });
    },
    /request\.messages is invalid/i,
  );
});

void test("sandbox LLM dispatch rejects malformed declared fields before service execution", async () => {
  const invalidRequests: DynamicValue[] = [
    { messages: [], maxTokens: Number.NaN },
    { messages: [], maxTokens: -1 },
    { messages: [], maxTokens: 1.5 },
    { messages: [], maxTokens: Number.POSITIVE_INFINITY },
    { messages: [], model: 42 },
    { messages: [], systemPrompt: false },
  ];
  let serviceCalls = 0;

  for (const request of invalidRequests) {
    await assert.rejects(
      dispatchSandboxLlmCompletion(
        { request },
        () => {
          serviceCalls += 1;
          return Promise.resolve({ content: "unexpected", model: "test/model" });
        },
      ),
      /llm_complete request\.(?:maxTokens|model|systemPrompt) is invalid/i,
    );
  }

  assert.equal(serviceCalls, 0);
});

void test("sandbox LLM dispatch forwards a valid normalized request to the service", async () => {
  let received: DynamicValue = null;
  const result = await dispatchSandboxLlmCompletion(
    {
      request: {
        messages: [{ role: "user", content: "hello", ignored: true }],
        maxTokens: 1,
        ignored: true,
      },
    },
    (request) => {
      received = request;
      return Promise.resolve({ content: "hi", model: "test/model" });
    },
  );

  assert.deepEqual(received, {
    messages: [{ role: "user", content: "hello" }],
    maxTokens: 1,
  });
  assert.deepEqual(result, { content: "hi", model: "test/model" });
});

void test("parseSandboxHttpRequestOptions normalizes method/headers/body", () => {
  const options = parseSandboxHttpRequestOptions({
    method: "POST",
    headers: {
      Authorization: "Bearer token",
      Invalid: 42,
    },
    body: "{\"ok\":true}",
    timeoutMs: 2500,
    connection: "  acme  ",
  });

  assert.deepEqual(options, {
    method: "POST",
    headers: {
      Authorization: "Bearer token",
    },
    body: "{\"ok\":true}",
    timeoutMs: 2500,
    connection: "acme",
  });

  assert.equal(parseSandboxHttpRequestOptions("not-an-object"), undefined);
});
