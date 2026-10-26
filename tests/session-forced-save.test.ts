import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxProvider } from "@earendil-works/pi-ai";

import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, unknown>>();

  private store(name: string): Map<string, unknown> {
    const existing = this.stores.get(name);
    if (existing) return existing;
    const created = new Map<string, unknown>();
    this.stores.set(name, created);
    return created;
  }

  get<T = unknown>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    return Promise.resolve(value === undefined ? null : structuredClone(value) as T);
  }

  set<T = unknown>(storeName: string, key: string, value: T): Promise<void> {
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

  getAllFromIndex<T = unknown>(storeName: string): Promise<T[]> {
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
      get: <V = unknown>(storeName: string, key: string) => this.get<V>(storeName, key),
      set: <V = unknown>(storeName: string, key: string, value: V) => this.set(storeName, key, value),
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

async function createRuntime(
  backend: StorageBackend,
  autoRestoreLatest = false,
  onSavedAsCopy?: (args: { fromSessionId: string; toSessionId: string; title: string }) => void,
) {
  const { sessions, settings } = initAppStorage("session-forced-save-test", backend);
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
    autoRestoreLatest,
    spreadsheetHost: {
      getWorkbookContext: () => Promise.resolve({
        workbookId: "url_sha256:forced-save-workbook",
        workbookName: null,
        source: "document.url",
      }),
      sessionStorage: settingsBackedSessionStorage,
    },
    ...(onSavedAsCopy ? { onSavedAsCopy } : {}),
  });
  return { agent, persistence, sessions };
}

void test("a forced save before the first assistant response survives a taskpane restart", async () => {
  const backend = new MemoryStorageBackend();
  const beforeRestart = await createRuntime(backend);
  const sessionId = beforeRestart.persistence.getSessionId();
  beforeRestart.agent.state.messages.push({
    role: "user",
    content: "waiting for the first answer",
    timestamp: 1_789_030_800_000,
  });

  await beforeRestart.persistence.saveSession({ force: true });
  beforeRestart.persistence.dispose();

  const afterRestart = await createRuntime(backend, true);
  assert.equal(afterRestart.persistence.getSessionId(), sessionId);
  assert.match(JSON.stringify(afterRestart.agent.state.messages), /waiting for the first answer/u);
});

void test("one session open in two taskpanes keeps both conversations", async () => {
  const backend = new MemoryStorageBackend();
  const paneA = await createRuntime(backend);
  paneA.agent.state.messages.push({ role: "user", content: "first turn", timestamp: 1 });
  await paneA.persistence.saveSession({ force: true });
  const sharedId = paneA.persistence.getSessionId();

  const copies: string[] = [];
  const paneB = await createRuntime(backend, false, ({ toSessionId }) => copies.push(toSessionId));
  const loaded = await paneB.sessions.get(sharedId);
  assert.ok(loaded);
  await paneB.persistence.applyLoadedSession(loaded);

  paneA.agent.state.messages.push({ role: "user", content: "Write 11 in K1", timestamp: 2 });
  paneB.agent.state.messages.push({ role: "user", content: "Write 22 in K2", timestamp: 3 });
  await paneA.persistence.saveSession({ force: true });
  await paneB.persistence.saveSession({ force: true });

  // Pane A's turn is not overwritten by pane B's older copy...
  assert.match(JSON.stringify((await paneA.sessions.get(sharedId))?.messages), /Write 11 in K1/u);
  // ...and pane B's turn is kept as a new session, which pane B now continues.
  assert.equal(copies.length, 1);
  assert.equal(paneB.persistence.getSessionId(), copies[0]);
  const copy = await paneB.sessions.get(copies[0] ?? "");
  assert.match(JSON.stringify(copy?.messages), /Write 22 in K2/u);
  assert.doesNotMatch(JSON.stringify(copy?.messages), /Write 11 in K1/u);

  // Later saves from either pane go through normally.
  paneB.agent.state.messages.push({ role: "user", content: "next", timestamp: 4 });
  await paneB.persistence.saveSession({ force: true });
  assert.equal(copies.length, 1);
  paneA.persistence.dispose();
  paneB.persistence.dispose();
});
