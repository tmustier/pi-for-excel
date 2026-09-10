import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";
import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import type { ActionQueue } from "../src/taskpane/action-queue.ts";
import type { QueueDisplay } from "../src/taskpane/queue-display.ts";
import { SessionRuntimeManager, type SessionRuntime } from "../src/taskpane/session-runtime-manager.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, DynamicValue>>();

  private store(name: string): Map<string, DynamicValue> {
    const existing = this.stores.get(name);
    if (existing) return existing;
    const created = new Map<string, DynamicValue>();
    this.stores.set(name, created);
    return created;
  }

  get<T = DynamicValue>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    return Promise.resolve(value === undefined ? null : structuredClone(value) as T);
  }

  set<T = DynamicValue>(storeName: string, key: string, value: T): Promise<void> {
    this.store(storeName).set(key, structuredClone(value));
    return Promise.resolve();
  }

  delete(storeName: string, key: string): Promise<void> {
    this.store(storeName).delete(key);
    return Promise.resolve();
  }

  keys(storeName: string, prefix?: string): Promise<string[]> {
    const keys = [...this.store(storeName).keys()];
    return Promise.resolve(prefix ? keys.filter((key) => key.startsWith(prefix)) : keys);
  }

  getAllFromIndex<T = DynamicValue>(storeName: string): Promise<T[]> {
    return Promise.resolve([...this.store(storeName).values()].map((value) => structuredClone(value)) as T[]);
  }

  clear(storeName: string): Promise<void> {
    this.store(storeName).clear();
    return Promise.resolve();
  }

  has(storeName: string, key: string): Promise<boolean> {
    return Promise.resolve(this.store(storeName).has(key));
  }

  transaction<T>(
    _storeNames: string[],
    _mode: "readonly" | "readwrite",
    operation: (tx: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    return operation({
      get: <V = DynamicValue>(storeName: string, key: string) => this.get<V>(storeName, key),
      set: <V = DynamicValue>(storeName: string, key: string, value: V) => this.set(storeName, key, value),
      delete: (storeName: string, key: string) => this.delete(storeName, key),
    });
  }

  getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return Promise.resolve({ usage: 0, quota: 0, percent: 0 });
  }

  requestPersistence(): Promise<boolean> {
    return Promise.resolve(true);
  }
}

const actionQueue: ActionQueue = {
  enqueuePrompt: () => {},
  enqueueCommand: () => {},
  drainQueuedActions: () => [],
  isBusy: () => false,
  shutdown: () => {},
};

function createQueueDisplay(): QueueDisplay {
  return {
    add: () => {},
    clear: () => {},
    drainQueuedMessages: () => [],
    setActionQueue: () => {},
    attach: () => {},
    detach: () => {},
  };
}

function createEnvironment() {
  const backend = new MemoryStorageBackend();
  const { sessions, settings } = initAppStorage("session-runtime-lifecycle", backend);
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  let nextRuntime = 1;

  const createRuntime = async (autoRestoreLatest: boolean): Promise<SessionRuntime> => {
    const agent = new Agent({
      initialState: {
        model: getBuiltinModel("openai", "gpt-5.6-sol"),
        messages: [],
        tools: [],
      },
      streamFn: (model, context, options) => models.streamSimple(model, context, options),
    });
    const persistence = await setupSessionPersistence({
      agent,
      sessions,
      settings,
      models,
      autoRestoreLatest,
      spreadsheetHost: {
        getWorkbookContext: () => Promise.resolve({
          workbookId: "url_sha256:runtime-lifecycle-workbook",
          workbookName: null,
          source: "document.url",
        }),
        sessionStorage: settingsBackedSessionStorage,
      },
    });
    const runtimeId = `runtime-${nextRuntime}`;
    nextRuntime += 1;
    return {
      runtimeId,
      agent,
      persistence,
      actionQueue,
      queueDisplay: createQueueDisplay(),
      lockState: "idle",
      refreshCapabilities: () => Promise.resolve(),
      dispose: () => persistence.dispose(),
    };
  };

  const manager = new SessionRuntimeManager({
    createRuntime: (options) => createRuntime(options.autoRestoreLatest),
  });
  return { manager, createRuntime };
}

void test("runtime tab snapshots follow create, switch, and close lifecycle", async () => {
  const environment = createEnvironment();
  const first = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const second = await environment.manager.createRuntime({ activate: false, autoRestoreLatest: false });

  assert.deepEqual(environment.manager.snapshot(), {
    activeRuntimeId: first.runtimeId,
    tabs: [
      {
        runtimeId: first.runtimeId,
        title: "Chat 1",
        isActive: true,
        isStreaming: false,
        isBusy: false,
        lockState: "idle",
      },
      {
        runtimeId: second.runtimeId,
        title: "Chat 2",
        isActive: false,
        isStreaming: false,
        isBusy: false,
        lockState: "idle",
      },
    ],
  });

  environment.manager.switchRuntime(second.runtimeId);
  environment.manager.closeRuntime(second.runtimeId);
  assert.deepEqual(environment.manager.snapshot().tabs.map((tab) => ({
    runtimeId: tab.runtimeId,
    isActive: tab.isActive,
  })), [{ runtimeId: first.runtimeId, isActive: true }]);
});

void test("fork preference updates an empty session in place and survives restart", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const sessionId = runtime.persistence.getSessionId();

  const result = await environment.manager.selectModel({
    runtimeId: runtime.runtimeId,
    nextModel: getBuiltinModel("openai-codex", "gpt-5.6-sol"),
    behavior: "fork",
  });
  assert.equal(result.outcome, "updated");
  assert.equal(environment.manager.snapshot().tabs.length, 1);
  runtime.agent.state.messages.push({
    role: "user",
    content: "message after choosing the model",
    timestamp: 1_789_030_800_000,
  });
  await runtime.persistence.saveSession({ force: true });
  runtime.dispose();

  const restarted = await environment.createRuntime(true);
  assert.equal(restarted.persistence.getSessionId(), sessionId);
  assert.equal(restarted.agent.state.model.provider, "openai-codex");
  restarted.dispose();
});
