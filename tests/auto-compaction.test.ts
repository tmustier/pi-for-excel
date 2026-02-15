import assert from "node:assert/strict";
import { test } from "node:test";

import type { AgentMessage } from "@mariozechner/pi-agent-core";

import { shouldAutoCompactForProjectedTokens } from "../src/compaction/auto-compaction.ts";
import {
  buildCompactionMemoryFocusInstruction,
  collectCompactionMemoryCues,
  mergeCompactionAdditionalFocus,
} from "../src/compaction/memory-nudge.ts";

function createUserMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "user",
    content: text,
    timestamp,
  };
}

function createAssistantMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "text", text }],
    timestamp,
  };
}

void test("does not trigger auto-compaction before hard threshold", () => {
  const shouldCompact = shouldAutoCompactForProjectedTokens({
    projectedTokens: 169_999,
    contextWindow: 200_000,
  });

  assert.equal(shouldCompact, false);
});

void test("triggers auto-compaction after hard threshold", () => {
  const shouldCompact = shouldAutoCompactForProjectedTokens({
    projectedTokens: 170_001,
    contextWindow: 200_000,
  });

  assert.equal(shouldCompact, true);
});

void test("small context windows still use reserve-based hard threshold", () => {
  const below = shouldAutoCompactForProjectedTokens({
    projectedTokens: 16_384,
    contextWindow: 32_768,
  });
  const above = shouldAutoCompactForProjectedTokens({
    projectedTokens: 16_385,
    contextWindow: 32_768,
  });

  assert.equal(below, false);
  assert.equal(above, true);
});

void test("collects memory cues from user messages and ignores auto-context", () => {
  const messages: AgentMessage[] = [
    createUserMessage("[Auto-context] Please remember this summary.", 1),
    createUserMessage("Please remember this: this workbook uses calendar year.", 2),
    createAssistantMessage("Got it.", 3),
    createUserMessage("Don't forget to keep EUR as the default currency.", 4),
  ];

  const summary = collectCompactionMemoryCues(messages);

  assert.equal(summary.cueCount, 2);
  assert.equal(summary.snippets.length, 2);
  assert.ok(summary.snippets.every((snippet) => !snippet.startsWith("[Auto-context]")));
  assert.match(summary.snippets[0] ?? "", /remember this/i);
  assert.match(summary.snippets[1] ?? "", /don['’]t forget/i);
});

void test("deduplicates snippets and respects snippet limits", () => {
  const messages: AgentMessage[] = [
    createUserMessage("Remember this: freeze panes on Summary.", 1),
    createUserMessage("Remember this: freeze panes on Summary.", 2),
    createUserMessage("Please save this for future reference: Revenue is net of refunds.", 3),
    createUserMessage("Please save this for future reference: Revenue is net of refunds.", 4),
  ];

  const summary = collectCompactionMemoryCues(messages, 1);

  assert.equal(summary.cueCount, 4);
  assert.equal(summary.snippets.length, 1);
});

void test("builds memory focus instructions only when cues are present", () => {
  const noCueInstruction = buildCompactionMemoryFocusInstruction({
    cueCount: 0,
    snippets: [],
  });
  assert.equal(noCueInstruction, null);

  const instruction = buildCompactionMemoryFocusInstruction({
    cueCount: 2,
    snippets: [
      "Remember this: workbook uses calendar year.",
      "Don't forget EUR defaults.",
    ],
  });

  assert.notEqual(instruction, null);
  assert.match(instruction ?? "", /instructions tool/i);
  assert.match(instruction ?? "", /notes\//i);
  assert.match(instruction ?? "", /Memory to persist/i);
  assert.match(instruction ?? "", /Potential user cues/i);
});

void test("merges compaction focus parts", () => {
  const merged = mergeCompactionAdditionalFocus(
    "focus on formulas",
    null,
    "capture durable memory",
  );

  assert.equal(merged, "focus on formulas\n\ncapture durable memory");
  assert.equal(mergeCompactionAdditionalFocus(" ", null, undefined), undefined);
});
