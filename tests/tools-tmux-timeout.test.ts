import assert from "node:assert/strict";
import { test } from "node:test";

import { createTmuxTool } from "../src/tools/tmux.ts";

function firstText(result: { content: Array<{ type: string; text?: string }> }): string {
  return result.content.find((item) => item.type === "text")?.text ?? "";
}

function installPendingFetch(): () => void {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (_input, init): Promise<Response> => new Promise((_resolve, reject) => {
    const signal = init?.signal;
    if (signal?.aborted) {
      reject(new DOMException("aborted", "AbortError"));
      return;
    }
    signal?.addEventListener("abort", () => {
      reject(new DOMException("aborted", "AbortError"));
    }, { once: true });
  });
  return () => {
    globalThis.fetch = originalFetch;
  };
}

void test("tmux execute includes wait and capture budgets in its network timeout", async (t) => {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const restoreFetch = installPendingFetch();
  try {
    const tool = createTmuxTool({
      getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    });

    const execution = tool.execute("timeout-budgets", {
      action: "send_and_capture",
      session: "build",
      text: "npm test",
      wait_ms: 30_000,
      timeout_ms: 120_000,
    });
    await Promise.resolve();
    t.mock.timers.tick(155_000);

    const result = await execution;
    assert.equal(firstText(result), "Error: Tmux bridge request timed out after 155000ms.\nSkill: tmux-bridge");
  } finally {
    restoreFetch();
  }
});

void test("tmux execute applies its default timeout and maximum accepted budget", async (t) => {
  t.mock.timers.enable({ apis: ["setTimeout"] });
  const restoreFetch = installPendingFetch();
  try {
    const tool = createTmuxTool({
      getBridgeConfig: () => Promise.resolve({ url: "https://localhost:3341" }),
    });

    const defaultExecution = tool.execute("timeout-default", {
      action: "send_and_capture",
      session: "build",
      text: "echo ready",
    });
    await Promise.resolve();
    t.mock.timers.tick(15_000);
    const defaultResult = await defaultExecution;

    const cappedExecution = tool.execute("timeout-cap", {
      action: "send_and_capture",
      session: "build",
      text: "echo done",
      wait_ms: 120_000,
      timeout_ms: 120_000,
    });
    await Promise.resolve();
    t.mock.timers.tick(245_000);
    const cappedResult = await cappedExecution;

    assert.deepEqual(
      [firstText(defaultResult), firstText(cappedResult)],
      [
        "Error: Tmux bridge request timed out after 15000ms.\nSkill: tmux-bridge",
        "Error: Tmux bridge request timed out after 245000ms.\nSkill: tmux-bridge",
      ],
    );
  } finally {
    restoreFetch();
  }
});
