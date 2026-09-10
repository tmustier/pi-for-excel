import assert from "node:assert/strict";
import { test } from "node:test";

import { describeExtensionRuntimeMode } from "../src/extensions/runtime-mode.ts";
import {
  collectSandboxUiActionIds,
  normalizeSandboxUiNode,
} from "../src/extensions/sandbox-ui.ts";

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
