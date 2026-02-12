import assert from "node:assert/strict";
import { test } from "node:test";

import {
  describeExtensionRuntimeMode,
  isSandboxCandidateTrust,
  resolveExtensionRuntimeMode,
} from "../src/extensions/runtime-mode.ts";
import {
  collectSandboxUiActionIds,
  normalizeSandboxUiNode,
} from "../src/extensions/sandbox-ui.ts";

void test("isSandboxCandidateTrust only flags inline/remote trust", () => {
  assert.equal(isSandboxCandidateTrust("builtin"), false);
  assert.equal(isSandboxCandidateTrust("local-module"), false);
  assert.equal(isSandboxCandidateTrust("inline-code"), true);
  assert.equal(isSandboxCandidateTrust("remote-url"), true);
});

void test("resolveExtensionRuntimeMode uses host runtime when sandbox flag is disabled", () => {
  assert.equal(resolveExtensionRuntimeMode("builtin", false), "host");
  assert.equal(resolveExtensionRuntimeMode("local-module", false), "host");
  assert.equal(resolveExtensionRuntimeMode("inline-code", false), "host");
  assert.equal(resolveExtensionRuntimeMode("remote-url", false), "host");
});

void test("resolveExtensionRuntimeMode routes inline/remote trust to sandbox when enabled", () => {
  assert.equal(resolveExtensionRuntimeMode("builtin", true), "host");
  assert.equal(resolveExtensionRuntimeMode("local-module", true), "host");
  assert.equal(resolveExtensionRuntimeMode("inline-code", true), "sandbox-iframe");
  assert.equal(resolveExtensionRuntimeMode("remote-url", true), "sandbox-iframe");
});

void test("describeExtensionRuntimeMode returns user-facing labels", () => {
  assert.equal(describeExtensionRuntimeMode("host"), "host runtime");
  assert.equal(describeExtensionRuntimeMode("sandbox-iframe"), "sandbox iframe");
});

void test("normalizeSandboxUiNode downgrades unsafe tags and strips invalid action ids", () => {
  const normalized = normalizeSandboxUiNode({
    kind: "element",
    tag: "script",
    className: "safe invalid! also_safe",
    actionId: "bad action id",
    children: [
      {
        kind: "text",
        text: "hello",
      },
    ],
  });

  assert.equal(normalized.kind, "element");
  if (normalized.kind !== "element") {
    return;
  }

  assert.equal(normalized.tag, "div");
  assert.equal(normalized.actionId, undefined);
  assert.equal(normalized.className, "safe also_safe");
  assert.equal(normalized.children.length, 1);
});

void test("normalizeSandboxUiNode preserves valid interactive action ids", () => {
  const normalized = normalizeSandboxUiNode({
    kind: "element",
    tag: "button",
    actionId: "widget:save",
    children: [
      {
        kind: "text",
        text: "Save",
      },
    ],
  });

  assert.equal(normalized.kind, "element");
  if (normalized.kind !== "element") {
    return;
  }

  assert.equal(normalized.tag, "button");
  assert.equal(normalized.actionId, "widget:save");

  const actionIds = collectSandboxUiActionIds(normalized);
  assert.deepEqual(actionIds, ["widget:save"]);
});

void test("normalizeSandboxUiNode falls back to empty text node for invalid payload", () => {
  const normalized = normalizeSandboxUiNode({
    kind: "unknown",
  });

  assert.deepEqual(normalized, {
    kind: "text",
    text: "",
  });
});
