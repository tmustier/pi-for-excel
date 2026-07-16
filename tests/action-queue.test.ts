import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, ImageContent, Model, Usage } from "@earendil-works/pi-ai";

import { commandRegistry, type SlashCommand } from "../src/commands/types.ts";
import { createActionQueue } from "../src/taskpane/action-queue.ts";

class TestAgent extends Agent {
  promptCalls: string[] = [];
  continueCalls = 0;
  waitForIdleCalls = 0;
  onPrompt?: (text: string) => void;
  onWaitForIdle?: () => Promise<void>;

  constructor(model?: Model<Api>) {
    super({
      initialState: {
        ...(model ? { model } : {}),
        messages: [],
        tools: [],
      },
    });
  }

  override continue(): Promise<void> {
    this.continueCalls += 1;
    return Promise.resolve();
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

const EMPTY_USAGE: Usage = {
  input: 0,
  output: 0,
  cacheRead: 0,
  cacheWrite: 0,
  totalTokens: 0,
  cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
};

function createSmallContextModel(): Model<Api> {
  return {
    id: "small-model",
    name: "Small Model",
    api: "openai-completions",
    provider: "custom-gateway",
    baseUrl: "https://gateway.example.invalid",
    reasoning: false,
    input: ["text"],
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
    contextWindow: 65_536,
    maxTokens: 4096,
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
        agent.state.messages = [
          createUserMessage("compaction summary", Date.now()),
        ];

        executionOrder.push("compact:end");
        compactFinished.resolve();
      },
    },
    async () => {
      const queue = createActionQueue({
        agent,
        autoCompactEnabled: false,
        runCompact: async () => {},
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
        runCompact: async () => {},
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
    runCompact: async () => {},
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

void test("prompt ending in context overflow triggers compact-and-retry once", async () => {
  const model = createSmallContextModel();
  const agent = new TestAgent(model);

  const toolResult: AgentMessage = {
    role: "toolResult",
    toolCallId: "call-1",
    toolName: "read_range",
    content: [{ type: "text", text: "rows..." }],
    isError: false,
    timestamp: 2,
  };

  agent.onPrompt = (text) => {
    // Simulate a run that ends in a provider context-overflow error.
    agent.state.messages = [
      createUserMessage(text, 1),
      toolResult,
      {
        role: "assistant",
        content: [{ type: "text", text: "" }],
        api: model.api,
        provider: model.provider,
        model: model.id,
        usage: EMPTY_USAGE,
        stopReason: "error",
        errorMessage: "Requested token count exceeds the model's maximum context length of 65536 tokens",
        timestamp: 3,
      },
    ];
  };

  let compactRuns = 0;
  const queue = createActionQueue({
    agent,
    autoCompactEnabled: true,
    runCompact: () => {
      compactRuns += 1;
      agent.state.messages = [createUserMessage("compaction summary", 4), toolResult];
      return Promise.resolve();
    },
    sidebar: {
      setBusyIndicator: () => {
        // no-op
      },
    },
    queueDisplay: {
      setActionQueue: () => {
        // no-op
      },
    },
  });

  queue.enqueuePrompt("analyze this data");

  await waitForCondition(() => agent.continueCalls === 1 && !queue.isBusy());

  assert.equal(compactRuns, 1);
  assert.deepEqual(agent.promptCalls, ["analyze this data"]);
  // Failed assistant message was removed from the retry context.
  assert.equal(
    agent.state.messages.some((m) => m.role === "assistant"),
    false,
  );

  queue.shutdown();
});

void test("shutdown uninstalls the mid-turn prepareNextTurn hook", () => {
  const agent = new TestAgent();
  const queue = createActionQueue({
    agent,
    autoCompactEnabled: true,
    runCompact: async () => {},
    sidebar: {
      setBusyIndicator: () => {
        // no-op
      },
    },
    queueDisplay: {
      setActionQueue: () => {
        // no-op
      },
    },
  });

  assert.equal(typeof agent.prepareNextTurn, "function");
  queue.shutdown();
  assert.equal(agent.prepareNextTurn, undefined);
});
