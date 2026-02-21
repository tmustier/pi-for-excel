import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@mariozechner/pi-agent-core";
import type { ImageContent } from "@mariozechner/pi-ai";

import { commandRegistry, type SlashCommand } from "../src/commands/types.ts";
import { createActionQueue } from "../src/taskpane/action-queue.ts";

class TestAgent extends Agent {
  promptCalls: string[] = [];
  waitForIdleCalls = 0;
  onPrompt?: (text: string) => void;
  onWaitForIdle?: () => Promise<void>;

  constructor() {
    super({
      initialState: {
        messages: [],
        tools: [],
      },
    });
  }

  override waitForIdle(): Promise<void> {
    this.waitForIdleCalls += 1;
    if (this.onWaitForIdle) {
      return this.onWaitForIdle();
    }

    return Promise.resolve();
  }

  override prompt(message: AgentMessage | AgentMessage[]): Promise<void>;
  override prompt(input: string, images?: ImageContent[]): Promise<void>;
  override prompt(
    input: string | AgentMessage | AgentMessage[],
    _images?: ImageContent[],
  ): Promise<void> {
    if (typeof input === "string") {
      this.promptCalls.push(input);
      this.onPrompt?.(input);
    }

    return Promise.resolve();
  }
}

function createDeferred(): { promise: Promise<void>; resolve: () => void } {
  let resolvePromise: (() => void) | undefined;

  const promise = new Promise<void>((resolve) => {
    resolvePromise = resolve;
  });

  return {
    promise,
    resolve: () => {
      resolvePromise?.();
    },
  };
}

async function waitForCondition(predicate: () => boolean, timeoutMs = 1500): Promise<void> {
  const start = Date.now();

  while (!predicate()) {
    if (Date.now() - start > timeoutMs) {
      throw new Error("Timed out waiting for condition");
    }

    await new Promise<void>((resolve) => {
      setTimeout(resolve, 5);
    });
  }
}

async function withRegisteredCommand(command: SlashCommand, run: () => Promise<void>): Promise<void> {
  const previous = commandRegistry.get(command.name);
  commandRegistry.register(command);

  try {
    await run();
  } finally {
    if (previous) {
      commandRegistry.register(previous);
    } else {
      commandRegistry.unregister(command.name);
    }
  }
}

function createUserMessage(text: string, timestamp: number): AgentMessage {
  return {
    role: "user",
    content: text,
    timestamp,
  };
}

void test("queued prompt survives compact replaceMessages and runs after compact", async () => {
  const agent = new TestAgent();
  const compactGate = createDeferred();
  const compactStarted = createDeferred();
  const compactFinished = createDeferred();

  const executionOrder: string[] = [];
  const busyIndicators: Array<{ label: string | null; hint: string | null }> = [];
  const queueSnapshots: Array<Array<{ type: "prompt" | "command"; label: string; text: string }>> = [];

  await withRegisteredCommand(
    {
      name: "compact",
      description: "compact",
      source: "builtin",
      execute: async () => {
        executionOrder.push("compact:start");
        compactStarted.resolve();

        await compactGate.promise;

        // Simulate compaction replacing message history while a prompt is queued.
        agent.replaceMessages([
          createUserMessage("compaction summary", Date.now()),
        ]);

        executionOrder.push("compact:end");
        compactFinished.resolve();
      },
    },
    async () => {
      const queue = createActionQueue({
        agent,
        autoCompactEnabled: false,
        sidebar: {
          setBusyIndicator: (label, hint) => {
            busyIndicators.push({ label, hint: hint ?? null });
          },
        },
        queueDisplay: {
          setActionQueue: (items) => {
            queueSnapshots.push(items.map((item) => ({ ...item })));
          },
        },
      });

      agent.onPrompt = (text) => {
        executionOrder.push(`prompt:${text}`);
      };

      queue.enqueueCommand("compact", "");

      await compactStarted.promise;
      assert.equal(queue.isBusy(), true);

      queue.enqueuePrompt("after compact");

      compactGate.resolve();

      await compactFinished.promise;
      await waitForCondition(() => agent.promptCalls.length === 1 && !queue.isBusy());

      assert.deepEqual(agent.promptCalls, ["after compact"]);
      assert.deepEqual(executionOrder, [
        "compact:start",
        "compact:end",
        "prompt:after compact",
      ]);

      assert.equal(busyIndicators[0]?.label, "Compacting context…");
      assert.equal(
        busyIndicators[0]?.hint,
        "Send messages and Pi will see them after compaction",
      );
      assert.equal(busyIndicators[busyIndicators.length - 1]?.label, null);

      const queuedPromptWasShown = queueSnapshots.some((snapshot) =>
        snapshot.some((item) => item.type === "prompt" && item.text === "after compact")
      );
      assert.equal(queuedPromptWasShown, true);

      queue.shutdown();
    },
  );
});

void test("ordered queue runs compact, prompt, then compact", async () => {
  const agent = new TestAgent();
  const compactRuns: Array<{ start: ReturnType<typeof createDeferred>; gate: ReturnType<typeof createDeferred> }> = [
    { start: createDeferred(), gate: createDeferred() },
    { start: createDeferred(), gate: createDeferred() },
  ];

  let compactRunIndex = 0;
  const executionOrder: string[] = [];

  await withRegisteredCommand(
    {
      name: "compact",
      description: "compact",
      source: "builtin",
      execute: async () => {
        const current = compactRunIndex;
        compactRunIndex += 1;

        executionOrder.push(`compact:${current + 1}:start`);
        compactRuns[current]?.start.resolve();
        await compactRuns[current]?.gate.promise;
        executionOrder.push(`compact:${current + 1}:end`);
      },
    },
    async () => {
      const queue = createActionQueue({
        agent,
        autoCompactEnabled: false,
        sidebar: {
          setBusyIndicator: () => {
            // not needed in this test
          },
        },
        queueDisplay: {
          setActionQueue: () => {
            // not needed in this test
          },
        },
      });

      agent.onPrompt = (text) => {
        executionOrder.push(`prompt:${text}`);
      };

      queue.enqueueCommand("compact", "");
      queue.enqueuePrompt("middle prompt");
      queue.enqueueCommand("compact", "");

      await compactRuns[0].start.promise;
      compactRuns[0].gate.resolve();

      await waitForCondition(() => agent.promptCalls.includes("middle prompt"));
      await compactRuns[1].start.promise;
      compactRuns[1].gate.resolve();

      await waitForCondition(() => !queue.isBusy());

      assert.deepEqual(executionOrder, [
        "compact:1:start",
        "compact:1:end",
        "prompt:middle prompt",
        "compact:2:start",
        "compact:2:end",
      ]);

      queue.shutdown();
    },
  );
});

void test("drainQueuedActions clears pending prompts and commands in FIFO order", async () => {
  const agent = new TestAgent();
  const waitGate = createDeferred();

  const queueSnapshots: Array<Array<{ type: "prompt" | "command"; label: string; text: string }>> = [];

  agent.onWaitForIdle = () => waitGate.promise;

  const queue = createActionQueue({
    agent,
    autoCompactEnabled: false,
    sidebar: {
      setBusyIndicator: () => {
        // no-op
      },
    },
    queueDisplay: {
      setActionQueue: (items) => {
        queueSnapshots.push(items.map((item) => ({ ...item })));
      },
    },
  });

  queue.enqueueCommand("compact", "");
  queue.enqueuePrompt("after compact");
  queue.enqueueCommand("compact", "deep");

  await waitForCondition(() => agent.waitForIdleCalls > 0);

  const drained = queue.drainQueuedActions();
  assert.deepEqual(drained, [
    { type: "command", name: "compact", args: "" },
    { type: "prompt", text: "after compact" },
    { type: "command", name: "compact", args: "deep" },
  ]);

  const latestSnapshot = queueSnapshots[queueSnapshots.length - 1] ?? [];
  assert.deepEqual(latestSnapshot, []);

  waitGate.resolve();
  await waitForCondition(() => !queue.isBusy());

  queue.shutdown();
});
