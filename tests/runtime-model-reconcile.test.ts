import assert from "node:assert/strict";
import { test } from "node:test";

import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import { resolveRuntimeModelSwap } from "../src/taskpane/runtime-model-reconcile.ts";

const openaiApiModel = getBuiltinModel("openai", "gpt-5.6-sol");

void test("does not swap onto a default model whose provider is also unusable", () => {
  // e.g. copilot-only setups where the default-model rules used to fall back
  // to openai/gpt-5.6-sol — trading one wrong API-key prompt for another.
  const swap = resolveRuntimeModelSwap({
    currentModel: getBuiltinModel("anthropic", "claude-opus-4-8"),
    availableProviders: ["github-copilot"],
    defaultModel: openaiApiModel,
    isBusy: false,
  });

  assert.equal(swap, null);
});

