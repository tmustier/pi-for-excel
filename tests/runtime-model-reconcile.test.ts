import assert from "node:assert/strict";
import { test } from "node:test";

import { getModel } from "@earendil-works/pi-ai";

import { resolveRuntimeModelSwap } from "../src/taskpane/runtime-model-reconcile.ts";

const openaiApiModel = getModel("openai", "gpt-5.5");
const codexModel = getModel("openai-codex", "gpt-5.5");

void test("swaps a runtime stuck on an unconfigured provider to the default model (#553)", () => {
  // Fresh-install flow: runtime created with the absolute fallback
  // (openai/gpt-5.5) before login, then the user connects ChatGPT.
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: ["openai-codex"],
    defaultModel: codexModel,
    isStreaming: false,
  });

  assert.ok(swap, "expected a swap for an unusable provider");
  assert.equal(swap.model.provider, "openai-codex");
  assert.equal(swap.model.id, "gpt-5.5");
  assert.equal(swap.thinkingLevel, codexModel.reasoning ? "high" : "off");
});

void test("leaves runtimes alone when their provider is still configured", () => {
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: ["openai", "openai-codex"],
    defaultModel: codexModel,
    isStreaming: false,
  });

  assert.equal(swap, null);
});

void test("does not swap while the runtime is streaming", () => {
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: ["openai-codex"],
    defaultModel: codexModel,
    isStreaming: true,
  });

  assert.equal(swap, null);
});

void test("does not swap when no providers are configured", () => {
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: [],
    defaultModel: codexModel,
    isStreaming: false,
  });

  assert.equal(swap, null);
});

void test("does not swap onto a default model whose provider is also unusable", () => {
  // e.g. copilot-only setups where the default-model rules used to fall back
  // to openai/gpt-5.5 — trading one wrong API-key prompt for another.
  const swap = resolveRuntimeModelSwap({
    currentModel: getModel("anthropic", "claude-opus-4-8"),
    availableProviders: ["github-copilot"],
    defaultModel: openaiApiModel,
    isStreaming: false,
  });

  assert.equal(swap, null);
});

void test("sets thinkingLevel to off when swapping onto a non-reasoning model", () => {
  const nonReasoning = {
    ...codexModel,
    reasoning: false,
  };

  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: ["openai-codex"],
    defaultModel: nonReasoning,
    isStreaming: false,
  });

  assert.ok(swap);
  assert.equal(swap.thinkingLevel, "off");
});
