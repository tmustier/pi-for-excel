import assert from "node:assert/strict";
import { test } from "node:test";

import { Type } from "@sinclair/typebox";
import type { Api, Context, Model } from "@mariozechner/pi-ai";

import {
  createPrefixFingerprint,
  getPrefixChangeReasons,
  PrefixChangeTracker,
} from "../src/auth/prefix-churn.ts";

function createModel(provider: string, id: string): Model<Api> {
  return {
    id,
    name: id,
    provider,
    api: "openai-responses",
    baseUrl: "https://example.com/v1",
    reasoning: false,
    input: ["text"],
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    },
    contextWindow: 200000,
    maxTokens: 8192,
  };
}

function createContext(args?: {
  systemPrompt?: string;
  includeTool?: boolean;
}): Context {
  const tools = args?.includeTool
    ? [
      {
        name: "search_docs",
        description: "Search docs",
        parameters: Type.Object({ query: Type.String() }),
      },
    ]
    : undefined;

  return {
    systemPrompt: args?.systemPrompt,
    messages: [],
    tools,
  };
}

void test("createPrefixFingerprint is deterministic for identical model/context", () => {
  const model = createModel("openai", "gpt-5.2");
  const context = createContext({ systemPrompt: "You are helpful", includeTool: true });

  const first = createPrefixFingerprint(model, context);
  const second = createPrefixFingerprint(model, context);

  assert.deepEqual(first, second);
});

void test("getPrefixChangeReasons reports model/system/tool deltas", () => {
  const previous = createPrefixFingerprint(
    createModel("openai", "gpt-5.2"),
    createContext({ systemPrompt: "base", includeTool: true }),
  );

  const next = createPrefixFingerprint(
    createModel("anthropic", "claude-sonnet-4-5"),
    createContext({ systemPrompt: "base v2", includeTool: false }),
  );

  const reasons = getPrefixChangeReasons(previous, next);
  assert.deepEqual(reasons, ["model", "systemPrompt", "tools"]);
});

void test("PrefixChangeTracker tracks prefixes per session id", () => {
  const tracker = new PrefixChangeTracker();
  const model = createModel("openai", "gpt-5.2");
  const initial = createPrefixFingerprint(model, createContext({ systemPrompt: "base" }));
  const updated = createPrefixFingerprint(model, createContext({ systemPrompt: "base v2" }));

  assert.deepEqual(tracker.observe("session-a", initial), []);
  assert.deepEqual(tracker.observe("session-a", updated), ["systemPrompt"]);

  // Different session should not inherit churn from session-a.
  assert.deepEqual(tracker.observe("session-b", updated), []);
});

void test("PrefixChangeTracker evicts oldest session keys when over budget", () => {
  const tracker = new PrefixChangeTracker({ maxTrackedSessions: 2 });

  const model = createModel("openai", "gpt-5.2");
  const base = createPrefixFingerprint(model, createContext({ systemPrompt: "base" }));
  const changed = createPrefixFingerprint(model, createContext({ systemPrompt: "changed" }));

  assert.deepEqual(tracker.observe("session-a", base), []);
  assert.deepEqual(tracker.observe("session-b", base), []);
  assert.deepEqual(tracker.observe("session-c", base), []);

  // session-a should have been evicted; first observation is treated as baseline again.
  assert.deepEqual(tracker.observe("session-a", changed), []);
});
