import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";

import type { ActionQueue } from "../src/taskpane/action-queue.ts";
import type { QueueDisplay } from "../src/taskpane/queue-display.ts";
import { createTabState } from "../src/taskpane/tab-state.ts";
import {
  SessionRuntimeManager,
  type SessionRuntime,
} from "../src/taskpane/session-runtime-manager.ts";
import type { SessionPersistenceController } from "../src/taskpane/sessions.ts";

interface RuntimeHarness {
  runtime: SessionRuntime;
  emitPersistenceChange: () => void;
  getPersistenceSubscriberCount: () => number;
  getDisposeCount: () => number;
  getRefreshCount: () => number;
}

function createAgent(): Agent {
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  return new Agent({
    initialState: {
      model: faux.getModel(),
      messages: [],
      tools: [],
    },
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
  });
}

function createRuntimeHarness(runtimeId: string): RuntimeHarness {
  const persistenceListeners = new Set<() => void>();
  let disposeCount = 0;
  let refreshCount = 0;
  const agent = createAgent();

  const persistence: SessionPersistenceController = {
    getSessionId: () => `session-${runtimeId}`,
    getSessionTitle: () => "",
    getSessionCreatedAt: () => "2026-01-01T00:00:00.000Z",
    hasExplicitTitle: () => false,
    startNewSession: () => {},
    renameSession: () => Promise.resolve(),
    applyLoadedSession: () => Promise.resolve(),
    restoreLatestSession: () => Promise.resolve(false),
    saveSession: () => Promise.resolve(),
    subscribe: (listener) => {
      persistenceListeners.add(listener);
      return () => {
        persistenceListeners.delete(listener);
      };
    },
    dispose: () => {},
  };

  const actionQueue: ActionQueue = {
    enqueuePrompt: () => {},
    enqueueCommand: () => {},
    editLast: () => false,
    stop: () => ({ texts: [], cancelledCommands: [] }),
    isBusy: () => false,
    shutdown: () => {},
  };

  const queueDisplay: QueueDisplay = {
    add: () => {},
    clear: () => {},
    drainQueuedMessages: () => [],
    removeQueuedMessage: () => {},
    setRowHandlers: () => {},
    lastEditableRow: () => null,
    setActionQueue: () => {},
    attach: () => {},
    detach: () => {},
  };

  return {
    runtime: {
      runtimeId,
      agent,
      actionQueue,
      queueDisplay,
      tab: createTabState(),
      persistence,
      lockState: "idle",
      refreshCapabilities: () => {
        refreshCount += 1;
        return Promise.resolve();
      },
      dispose: () => {
        disposeCount += 1;
      },
    },
    emitPersistenceChange: () => {
      for (const listener of persistenceListeners) listener();
    },
    getPersistenceSubscriberCount: () => persistenceListeners.size,
    getDisposeCount: () => disposeCount,
    getRefreshCount: () => refreshCount,
  };
}

void test("closing a runtime unsubscribes lifecycle persistence events before disposal", () => {
  const first = createRuntimeHarness("first");
  const second = createRuntimeHarness("second");
  const manager = new SessionRuntimeManager({
    createRuntime: () => Promise.reject(new Error("factory should not run")),
  });

  manager.registerRuntime(first.runtime, { activate: true });
  manager.registerRuntime(second.runtime, { activate: false });

  let snapshotCount = 0;
  manager.subscribe(() => {
    snapshotCount += 1;
  });
  first.emitPersistenceChange();
  assert.equal(snapshotCount, 2);
  assert.equal(first.getPersistenceSubscriberCount(), 1);

  manager.closeRuntime("first");
  const countAfterClose = snapshotCount;
  assert.equal(first.getPersistenceSubscriberCount(), 0);
  assert.equal(first.getDisposeCount(), 1);
  assert.equal(manager.getActiveRuntime()?.runtimeId, "second");

  first.emitPersistenceChange();
  assert.equal(snapshotCount, countAfterClose);
});

void test("capability refresh is owned by the lifecycle and visits every runtime", async () => {
  const first = createRuntimeHarness("first");
  const second = createRuntimeHarness("second");
  const manager = new SessionRuntimeManager({
    createRuntime: () => Promise.reject(new Error("factory should not run")),
  });
  manager.registerRuntime(first.runtime, { activate: true });
  manager.registerRuntime(second.runtime, { activate: false });

  await manager.refreshCapabilities();
  assert.equal(first.getRefreshCount(), 1);
  assert.equal(second.getRefreshCount(), 1);
});

void test("capability refresh requested during a pass runs again and a listener-triggered refresh is not lost", async () => {
  const first = createRuntimeHarness("first");
  const manager = new SessionRuntimeManager({
    createRuntime: () => Promise.reject(new Error("factory should not run")),
  });
  manager.registerRuntime(first.runtime, { activate: true });

  let release: () => void = () => {};
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  let refreshCount = 0;
  first.runtime.refreshCapabilities = async () => {
    refreshCount += 1;
    if (refreshCount === 1) await gate;
  };

  const inFlight = manager.refreshCapabilities();
  const coalesced = manager.refreshCapabilities();
  release();
  await Promise.all([inFlight, coalesced]);
  assert.equal(refreshCount, 2, "a request during an in-flight pass must run a second pass");

  let listenerRequests = 0;
  let armed = false;
  const unsubscribe = manager.subscribe(() => {
    if (!armed || listenerRequests > 0) return;
    listenerRequests += 1;
    void manager.refreshCapabilities();
  });
  armed = true;
  await manager.refreshCapabilities();
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  unsubscribe();
  assert.equal(listenerRequests, 1);
  assert.equal(refreshCount, 4, "a refresh requested from a snapshot listener must execute");
});

void test("a run that ends in a background tab marks it until the tab is opened; a user stop leaves no mark", () => {
  const first = createRuntimeHarness("first");
  const second = createRuntimeHarness("second");
  const busy = new Map<string, boolean>();
  const outcomes = new Map<string, "finished" | "failed" | null>();
  for (const harness of [first, second]) {
    const runtimeId = harness.runtime.runtimeId;
    harness.runtime.actionQueue.isBusy = () => busy.get(runtimeId) ?? false;
  }
  const settled: Array<{ runtimeId: string; outcome: string; wasActive: boolean; title: string }> = [];
  const manager = new SessionRuntimeManager({
    createRuntime: () => Promise.reject(new Error("factory should not run")),
    getRunOutcome: (runtime) => outcomes.get(runtime.runtimeId) ?? null,
    onRunSettled: (event) => settled.push(event),
  });
  manager.registerRuntime(first.runtime, { activate: true });
  manager.registerRuntime(second.runtime, { activate: false });
  const attentionOf = (runtimeId: string) =>
    manager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.attention;

  // Background run finishes.
  busy.set("second", true);
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("second"), "none");
  busy.set("second", false);
  outcomes.set("second", "finished");
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("second"), "finished");
  assert.deepEqual(settled, [{ runtimeId: "second", outcome: "finished", wasActive: false, title: "Chat 2" }]);

  // Settles once: further emits while idle report nothing new.
  manager.notifyRuntimeChanged();
  assert.equal(settled.length, 1);

  // Failure overrides; a new run clears the mark while it works.
  busy.set("second", true);
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("second"), "none");
  busy.set("second", false);
  outcomes.set("second", "failed");
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("second"), "failed");

  // Opening the tab clears it.
  manager.switchRuntime("second");
  assert.equal(attentionOf("second"), "none");

  // A user stop in the (now background) first tab: no mark, no event.
  settled.length = 0;
  busy.set("first", true);
  manager.notifyRuntimeChanged();
  busy.set("first", false);
  outcomes.set("first", null);
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("first"), "none");
  assert.equal(settled.length, 0);

  // The active tab finishing is reported as active and never marked.
  busy.set("second", true);
  manager.notifyRuntimeChanged();
  busy.set("second", false);
  outcomes.set("second", "finished");
  manager.notifyRuntimeChanged();
  assert.equal(attentionOf("second"), "none");
  assert.deepEqual(settled.map((event) => event.wasActive), [true]);
});
