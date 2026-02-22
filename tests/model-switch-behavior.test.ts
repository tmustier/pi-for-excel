import assert from "node:assert/strict";
import { test } from "node:test";

import {
  MODEL_SWITCH_BEHAVIOR_SETTING_KEY,
  getStoredModelSwitchBehavior,
  isModelSwitchBehavior,
  normalizeModelSwitchBehavior,
  setStoredModelSwitchBehavior,
  shouldForkModelSwitch,
  type ModelSwitchBehavior,
} from "../src/models/switch-behavior.ts";

class MemoryStore {
  private readonly values = new Map<string, unknown>();

  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.has(key) ? this.values.get(key) ?? null : null);
  }

  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

void test("normalizeModelSwitchBehavior defaults to inPlace", () => {
  assert.equal(normalizeModelSwitchBehavior(undefined), "inPlace");
  assert.equal(normalizeModelSwitchBehavior("unknown"), "inPlace");
});

void test("isModelSwitchBehavior only accepts known values", () => {
  assert.equal(isModelSwitchBehavior("inPlace"), true);
  assert.equal(isModelSwitchBehavior("fork"), true);
  assert.equal(isModelSwitchBehavior("forked"), false);
});

void test("stored model switch behavior defaults to inPlace", async () => {
  const store = new MemoryStore();
  const behavior = await getStoredModelSwitchBehavior(store);

  assert.equal(behavior, "inPlace");
});

void test("setStoredModelSwitchBehavior persists and returns the stored value", async () => {
  const store = new MemoryStore();

  const next = await setStoredModelSwitchBehavior(store, "fork");
  assert.equal(next, "fork");

  const stored = await store.get(MODEL_SWITCH_BEHAVIOR_SETTING_KEY);
  assert.equal(stored, "fork");
});

void test("shouldForkModelSwitch only forks for non-empty sessions with fork behavior", () => {
  const scenarios: Array<{ behavior: ModelSwitchBehavior; hasMessages: boolean; expected: boolean }> = [
    { behavior: "inPlace", hasMessages: false, expected: false },
    { behavior: "inPlace", hasMessages: true, expected: false },
    { behavior: "fork", hasMessages: false, expected: false },
    { behavior: "fork", hasMessages: true, expected: true },
  ];

  for (const scenario of scenarios) {
    assert.equal(
      shouldForkModelSwitch({
        behavior: scenario.behavior,
        hasMessages: scenario.hasMessages,
      }),
      scenario.expected,
    );
  }
});
