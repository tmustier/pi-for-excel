import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
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
    isBusy: false,
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
    isBusy: false,
  });

  assert.equal(swap, null);
});

void test("does not swap while the runtime is working (streaming or queue-busy)", () => {
  // isBusy covers both `agent.state.isStreaming` and `actionQueue.isBusy()`
  // — the latter is true for non-streaming work too (/compact, pre-prompt
  // auto-compaction, queued slash commands). Working sessions must never
  // have their model swapped underneath them.
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: ["openai-codex"],
    defaultModel: codexModel,
    isBusy: true,
  });

  assert.equal(swap, null);
});

void test("does not swap when no providers are configured", () => {
  const swap = resolveRuntimeModelSwap({
    currentModel: openaiApiModel,
    availableProviders: [],
    defaultModel: codexModel,
    isBusy: false,
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
    isBusy: false,
  });

  assert.equal(swap, null);
});

void test("init.ts reconcile loop guards on streaming AND queue-busy runtimes", () => {
  // Wiring pin for the working-session invariant: the reconcile loop in
  // init.ts must skip runtimes doing non-streaming work (actionQueue busy)
  // as well as streaming ones, before any model mutation.
  const initSource = readFileSync(
    path.resolve(process.cwd(), "src/taskpane/init.ts"),
    "utf8",
  );

  assert.notEqual(
    initSource.indexOf(
      "if (runtime.agent.state.isStreaming || runtime.actionQueue.isBusy()) {",
    ),
    -1,
    "expected reconcile loop to skip streaming or queue-busy runtimes",
  );
  assert.notEqual(
    initSource.indexOf(
      "isBusy: runtime.agent.state.isStreaming || runtime.actionQueue.isBusy(),",
    ),
    -1,
    "expected resolveRuntimeModelSwap to receive the combined busy state",
  );
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
    isBusy: false,
  });

  assert.ok(swap);
  assert.equal(swap.thinkingLevel, "off");
});
