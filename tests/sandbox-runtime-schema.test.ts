import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "typebox";

import { normalizeSandboxToolParameters } from "../src/extensions/sandbox-runtime.ts";
import {
  dispatchSandboxLlmCompletion,
  parseSandboxHttpRequestOptions,
} from "../src/extensions/sandbox/runtime-helpers.ts";
void test("normalizeSandboxToolParameters keeps TypeBox schema unchanged", () => {
  const schema = Type.Object({
    text: Type.String(),
  });

  const normalized = normalizeSandboxToolParameters(schema);
  assert.equal(normalized, schema);
});

void test("sandbox LLM dispatch rejects malformed declared fields before service execution", async () => {
  const invalidRequests: unknown[] = [
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
  let received: unknown = null;
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
