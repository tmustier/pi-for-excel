import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai";

import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";
import { workbookLatestSessionKey } from "../src/workbook/session-association.ts";

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

  getAllFromIndex<T = DynamicValue>(
    storeName: string,
    _indexName: string,
    direction: "asc" | "desc" = "asc",
  ): Promise<T[]> {
    const values = [...this.store(storeName).values()];
    values.sort((left, right) => {
      const leftValue = readLastModified(left);
      const rightValue = readLastModified(right);
      return direction === "desc"
        ? rightValue.localeCompare(leftValue)
        : leftValue.localeCompare(rightValue);
    });
    return Promise.resolve(values.map((value) => structuredClone(value)) as T[]);
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

function readLastModified(value: DynamicValue): string {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return "";
  const candidate = value as { lastModified?: DynamicValue };
  return typeof candidate.lastModified === "string" ? candidate.lastModified : "";
}

const workbookA = "url_sha256:workbook-a";
const sessionA = "00000000-0000-4000-8000-000000000001";
const sessionB = "00000000-0000-4000-8000-000000000002";

async function createRuntime(backend: StorageBackend, workbookId: string | null) {
  const { sessions, settings } = initAppStorage("session-restart-selection", backend);
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  const agent = new Agent({
    initialState: { model: faux.getModel(), messages: [], tools: [] },
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
  });
  const persistence = await setupSessionPersistence({
    agent,
    sessions,
    settings,
    models,
    spreadsheetHost: {
      getWorkbookContext: () => Promise.resolve({
        workbookId,
        workbookName: null,
        source: workbookId ? "document.url" : "unknown",
      }),
      sessionStorage: settingsBackedSessionStorage,
    },
  });
  return { agent, faux, persistence, sessions, settings };
}

async function seedSession(
  backend: StorageBackend,
  id: string,
  marker: string,
  lastModified: string,
): Promise<void> {
  const runtime = await createRuntime(backend, null);
  runtime.agent.state.messages = [{
    role: "user",
    content: marker,
    timestamp: Date.parse(lastModified),
  }];
  await runtime.sessions.saveSession(id, runtime.agent.state, {
    id,
    title: marker,
    createdAt: lastModified,
    lastModified,
    messageCount: 1,
    usage: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    thinkingLevel: "off",
    preview: marker,
  });
  runtime.persistence.dispose();
}

void test("restart selection follows workbook identity and persisted latest pointers", async () => {
  const scenarios: Array<{
    name: string;
    workbookId: string | null;
    workbookPointer: string | null;
    expectedSessionId: string | null;
  }> = [
    {
      name: "known workbook prefers its pointer over a newer global session",
      workbookId: workbookA,
      workbookPointer: sessionA,
      expectedSessionId: sessionA,
    },
    {
      name: "known workbook without a pointer restores nothing",
      workbookId: workbookA,
      workbookPointer: null,
      expectedSessionId: null,
    },
    {
      name: "known workbook ignores a blank pointer",
      workbookId: workbookA,
      workbookPointer: "   ",
      expectedSessionId: null,
    },
    {
      name: "unknown workbook restores the global latest session",
      workbookId: null,
      workbookPointer: sessionA,
      expectedSessionId: sessionB,
    },
  ];

  for (const scenario of scenarios) {
    const backend = new MemoryStorageBackend();
    await seedSession(backend, sessionA, "workbook-a conversation", "2026-09-10T09:00:00.000Z");
    await seedSession(backend, sessionB, "newer global conversation", "2026-09-10T10:00:00.000Z");
    if (scenario.workbookPointer !== null) {
      await backend.set("settings", workbookLatestSessionKey(workbookA), scenario.workbookPointer);
    }

    const restarted = await createRuntime(backend, scenario.workbookId);
    const restored = await restarted.persistence.restoreLatestSession();

    assert.equal(restored, scenario.expectedSessionId !== null, scenario.name);
    if (scenario.expectedSessionId !== null) {
      assert.equal(restarted.persistence.getSessionId(), scenario.expectedSessionId, scenario.name);
    } else {
      assert.equal(restarted.agent.state.messages.length, 0, scenario.name);
    }
    restarted.persistence.dispose();
  }
});

void test("a newer saved conversation becomes the workbook's restart target", async () => {
  const backend = new MemoryStorageBackend();
  const first = await createRuntime(backend, workbookA);
  first.faux.setResponses([fauxAssistantMessage("first answer")]);
  await first.agent.prompt("first conversation");
  await first.persistence.saveSession({ force: true });
  const firstId = first.persistence.getSessionId();
  first.persistence.startNewSession();
  first.agent.state.messages = [];
  first.faux.setResponses([fauxAssistantMessage("second answer")]);
  await first.agent.prompt("second conversation");
  await first.persistence.saveSession({ force: true });
  const secondId = first.persistence.getSessionId();
  first.persistence.dispose();

  const restarted = await createRuntime(backend, workbookA);
  assert.equal(await restarted.persistence.restoreLatestSession(), true);
  assert.notEqual(secondId, firstId);
  assert.equal(restarted.persistence.getSessionId(), secondId);
  assert.match(JSON.stringify(restarted.agent.state.messages), /second conversation/u);
});
