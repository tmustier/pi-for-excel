import assert from "node:assert/strict";
import { test } from "node:test";

import { Kind, Type } from "@sinclair/typebox";

import { normalizeSandboxToolParameters } from "../src/extensions/sandbox-runtime.ts";
import { isRecord } from "../src/utils/type-guards.ts";

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

  assert.ok(isRecord(normalized));
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
