import assert from "node:assert/strict";
import { test } from "node:test";

import { createExecutionModeController } from "../src/execution/controller.ts";

function sharedStore() {
  const values = new Map<string, unknown>();
  return {
    get: (key: string) => Promise.resolve(values.get(key) ?? null),
    set: (key: string, value: unknown) => {
      values.set(key, value);
      return Promise.resolve();
    },
  };
}

/** Two ends of one channel, like BroadcastChannel between two taskpanes. */
function channelPair() {
  const listeners: Array<Array<() => void>> = [[], []];
  const end = (self: number, other: number) => ({
    postMessage: () => {
      for (const listener of listeners[other] ?? []) listener();
    },
    addEventListener: (_type: string, listener: () => void) => {
      listeners[self]?.push(listener);
    },
  });
  return [end(0, 1), end(1, 0)] as const;
}

void test("two taskpanes never run in different execution modes", async () => {
  const settings = sharedStore();
  const [channelA, channelB] = channelPair();
  // Pane B has no live channel: it must still decide approvals on the stored mode.
  const paneA = await createExecutionModeController({ settings, channel: channelA });
  const paneB = await createExecutionModeController({ settings, channel: null });
  const paneC = await createExecutionModeController({ settings, channel: channelB });

  await paneA.setMode("safe");
  assert.equal(await paneB.readMode(), "safe");
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(paneC.getMode(), "safe");

  await paneB.setMode("yolo");
  assert.equal(await paneA.readMode(), "yolo");
});
