import assert from "node:assert/strict";
import { test } from "node:test";

import {
  getSendText,
  resolveInputAutoGrowHeight,
  shouldSendOnEnter,
} from "../src/ui/pi-input-behavior.ts";

void test("getSendText trims non-empty prompts and rejects whitespace", () => {
  assert.equal(getSendText("  build a TSLA P&L  "), "build a TSLA P&L");
  assert.equal(getSendText(" \n\t "), null);
});

void test("shouldSendOnEnter only submits real non-slash prompts", () => {
  assert.equal(shouldSendOnEnter({ key: "Enter", shiftKey: false, isStreaming: false, value: "write A1" }), true);
  assert.equal(shouldSendOnEnter({ key: "Enter", shiftKey: true, isStreaming: false, value: "write A1" }), false);
  assert.equal(shouldSendOnEnter({ key: "a", shiftKey: false, isStreaming: false, value: "write A1" }), false);
  assert.equal(shouldSendOnEnter({ key: "Enter", shiftKey: false, isStreaming: true, value: "write A1" }), false);
  assert.equal(shouldSendOnEnter({ key: "Enter", shiftKey: false, isStreaming: false, value: "   " }), false);
  assert.equal(shouldSendOnEnter({ key: "Enter", shiftKey: false, isStreaming: false, value: "/help" }), false);
});

void test("resolveInputAutoGrowHeight respects CSS cap before viewport fallback", () => {
  assert.equal(resolveInputAutoGrowHeight({ scrollHeight: 220, viewportHeight: 800, cssMaxHeight: 120 }), 120);
  assert.equal(resolveInputAutoGrowHeight({ scrollHeight: 80, viewportHeight: 800, cssMaxHeight: 120 }), 80);
});

void test("resolveInputAutoGrowHeight uses a smaller short-pane viewport fallback", () => {
  assert.equal(resolveInputAutoGrowHeight({ scrollHeight: 220, viewportHeight: 500 }), 140);
  assert.equal(resolveInputAutoGrowHeight({ scrollHeight: 220, viewportHeight: 800 }), 220);
  assert.equal(resolveInputAutoGrowHeight({ scrollHeight: 500, viewportHeight: 800 }), 320);
});
