import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";

import type { ActionQueue } from "../src/taskpane/action-queue.ts";
import type { QueueDisplay } from "../src/taskpane/queue-display.ts";
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
    drainQueuedActions: () => [],
    isBusy: () => false,
    shutdown: () => {},
  };

  const queueDisplay: QueueDisplay = {
    add: () => {},
    clear: () => {},
    drainQueuedMessages: () => [],
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
