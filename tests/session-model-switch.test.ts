import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai";
import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import type { ActionQueue } from "../src/taskpane/action-queue.ts";
import type { QueueDisplay } from "../src/taskpane/queue-display.ts";
import {
  SessionRuntimeManager,
  type SessionRuntime,
} from "../src/taskpane/session-runtime-manager.ts";
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

interface RuntimeEnvironment {
  manager: SessionRuntimeManager;
  sessions: ReturnType<typeof initAppStorage>["sessions"];
  createRuntime: (autoRestoreLatest?: boolean) => Promise<SessionRuntime>;
  respond: (text: string) => void;
}

function createEnvironment(backend: StorageBackend): RuntimeEnvironment {
  const { sessions, settings } = initAppStorage("session-model-switch-test", backend);
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  let nextRuntimeNumber = 1;

  const createRuntime = async (autoRestoreLatest = false): Promise<SessionRuntime> => {
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
          workbookId: "url_sha256:model-switch-workbook",
          workbookName: null,
          source: "document.url",
        }),
        sessionStorage: settingsBackedSessionStorage,
      },
    });
    const runtimeId = `runtime-${nextRuntimeNumber}`;
    nextRuntimeNumber += 1;
    return {
      runtimeId,
      agent,
      actionQueue,
      queueDisplay: createQueueDisplay(),
      persistence,
      lockState: "idle",
      refreshCapabilities: () => Promise.resolve(),
      dispose: () => persistence.dispose(),
    };
  };

  const manager = new SessionRuntimeManager({
    createRuntime: () => createRuntime(false),
  });
  return {
    manager,
    sessions,
    createRuntime,
    respond: (text: string) => faux.setResponses([fauxAssistantMessage(text)]),
  };
}

async function seedConversation(environment: RuntimeEnvironment): Promise<SessionRuntime> {
  const runtime = await environment.createRuntime();
  environment.manager.registerRuntime(runtime, { activate: true });
  environment.respond("original answer");
  await runtime.agent.prompt("original question");
  await runtime.persistence.saveSession({ force: true });
  return runtime;
}

async function restartSession(
  environment: RuntimeEnvironment,
  sessionId: string,
): Promise<SessionRuntime> {
  const saved = await environment.sessions.loadSession(sessionId);
  assert.ok(saved);
  const restarted = await environment.createRuntime();
  await restarted.persistence.applyLoadedSession(saved);
  return restarted;
}

void test("in-place model switch keeps the session and restores its selected model", async () => {
  const environment = createEnvironment(new MemoryStorageBackend());
  const original = await seedConversation(environment);
  const originalSessionId = original.persistence.getSessionId();

  const result = await environment.manager.selectModel({
    runtimeId: original.runtimeId,
    nextModel: getBuiltinModel("openai-codex", "gpt-5.6-sol"),
    behavior: "inPlace",
  });
  assert.equal(result.outcome, "updated");
  original.dispose();

  const restarted = await restartSession(environment, originalSessionId);
  assert.equal(restarted.persistence.getSessionId(), originalSessionId);
  assert.equal(restarted.agent.state.model.provider, "openai-codex");
  assert.match(JSON.stringify(restarted.agent.state.messages), /original question/u);
});

void test("fork model switch preserves the source session and restores a new selected-model session", async () => {
  const environment = createEnvironment(new MemoryStorageBackend());
  const original = await seedConversation(environment);
  const originalSessionId = original.persistence.getSessionId();

  const result = await environment.manager.selectModel({
    runtimeId: original.runtimeId,
    nextModel: getBuiltinModel("openai-codex", "gpt-5.6-sol"),
    behavior: "fork",
  });
  assert.equal(result.outcome, "forked");
  if (result.outcome !== "forked") return;
  const forkSessionId = result.runtime.persistence.getSessionId();
  assert.notEqual(forkSessionId, originalSessionId);
  original.dispose();
  result.runtime.dispose();

  const restartedSource = await restartSession(environment, originalSessionId);
  const restartedFork = await restartSession(environment, forkSessionId);
  assert.equal(restartedSource.agent.state.model.provider, "openai");
  assert.equal(restartedFork.agent.state.model.provider, "openai-codex");
  assert.match(JSON.stringify(restartedFork.agent.state.messages), /original question/u);
});
