import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  type Models,
} from "@earendil-works/pi-ai";

import {
  APP_STORAGE_DATABASE_NAME,
  APP_STORAGE_DATABASE_VERSION,
  getAppStorageConfig,
  initAppStorage,
} from "../src/storage/init-app-storage.ts";
import type {
  StorageBackend,
  StorageTransaction,
} from "../src/storage/local/types.ts";
import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";
import {
  sessionWorkbookKey,
  workbookLatestSessionKey,
} from "../src/workbook/session-association.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, DynamicValue>>();

  private store(name: string): Map<string, DynamicValue> {
    let store = this.stores.get(name);
    if (!store) {
      store = new Map<string, DynamicValue>();
      this.stores.set(name, store);
    }
    return store;
  }

  get<T = DynamicValue>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    // This boundary mirrors IndexedDB: callers own the requested persisted type.
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
    indexName: string,
    direction: "asc" | "desc" = "asc",
  ): Promise<T[]> {
    assert.equal(indexName, "lastModified");
    const values = [...this.store(storeName).values()].map((value) => structuredClone(value));
    values.sort((left, right) => {
      const leftModified = readLastModified(left);
      const rightModified = readLastModified(right);
      return direction === "desc"
        ? rightModified.localeCompare(leftModified)
        : leftModified.localeCompare(rightModified);
    });
    // This boundary mirrors IndexedDB: callers own the requested persisted type.
    return Promise.resolve(values as T[]);
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

interface RuntimeHarness {
  agent: Agent;
  controller: Awaited<ReturnType<typeof setupSessionPersistence>>;
  models: Models;
  sessions: ReturnType<typeof initAppStorage>["sessions"];
  settings: ReturnType<typeof initAppStorage>["settings"];
  respond: (text: string) => void;
}

async function createRuntime(
  backend: StorageBackend,
  workbookId: string | null,
): Promise<RuntimeHarness> {
  const { sessions, settings } = initAppStorage(APP_STORAGE_DATABASE_NAME, backend);
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const agent = new Agent({
    initialState: { model, messages: [], tools: [] },
    streamFn: (requestModel, context, options) => models.streamSimple(requestModel, context, options),
  });
  const controller = await setupSessionPersistence({
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

  return {
    agent,
    controller,
    models,
    sessions,
    settings,
    respond: (text: string) => faux.setResponses([fauxAssistantMessage(text)]),
  };
}

async function promptAndWaitForSave(runtime: RuntimeHarness, prompt: string): Promise<string> {
  runtime.respond(`reply:${prompt}`);
  await runtime.agent.prompt(prompt);
  const sessionId = runtime.controller.getSessionId();

  for (let attempt = 0; attempt < 50; attempt += 1) {
    const saved = await runtime.sessions.loadSession(sessionId);
    if (saved?.messages.length === 2) return sessionId;
    await new Promise<void>((resolve) => setTimeout(resolve, 1));
  }

  assert.fail(`session ${sessionId} was not persisted after agent message events`);
}

function transcriptText(runtime: RuntimeHarness): string {
  return JSON.stringify(runtime.agent.state.messages);
}

void test("restart restores workbook A without exposing its session to workbook B", async () => {
  const backend = new MemoryStorageBackend();
  const beforeRestart = await createRuntime(backend, "url_sha256:workbook-a");
  const sessionId = await promptAndWaitForSave(beforeRestart, "alpha-only");
  beforeRestart.controller.dispose();

  const workbookAAfterRestart = await createRuntime(backend, "url_sha256:workbook-a");
  assert.equal(await workbookAAfterRestart.controller.restoreLatestSession(), true);
  assert.equal(workbookAAfterRestart.controller.getSessionId(), sessionId);
  assert.match(transcriptText(workbookAAfterRestart), /alpha-only/);

  const workbookBAfterRestart = await createRuntime(backend, "url_sha256:workbook-b");
  assert.equal(await workbookBAfterRestart.controller.restoreLatestSession(), false);
  assert.equal(workbookBAfterRestart.agent.state.messages.length, 0);
});

void test("restart keeps two workbooks' sessions isolated", async () => {
  const backend = new MemoryStorageBackend();
  const workbookA = await createRuntime(backend, "url_sha256:workbook-a");
  const sessionA = await promptAndWaitForSave(workbookA, "workbook-a-message");
  const workbookB = await createRuntime(backend, "url_sha256:workbook-b");
  const sessionB = await promptAndWaitForSave(workbookB, "workbook-b-message");
  workbookA.controller.dispose();
  workbookB.controller.dispose();

  const restartedA = await createRuntime(backend, "url_sha256:workbook-a");
  const restartedB = await createRuntime(backend, "url_sha256:workbook-b");
  assert.equal(await restartedA.controller.restoreLatestSession(), true);
  assert.equal(await restartedB.controller.restoreLatestSession(), true);
  assert.equal(restartedA.controller.getSessionId(), sessionA);
  assert.equal(restartedB.controller.getSessionId(), sessionB);
  assert.match(transcriptText(restartedA), /workbook-a-message/);
  assert.doesNotMatch(transcriptText(restartedA), /workbook-b-message/);
  assert.match(transcriptText(restartedB), /workbook-b-message/);
  assert.doesNotMatch(transcriptText(restartedB), /workbook-a-message/);
});

void test("corrupt persisted session is skipped and a recoverable sibling remains", async () => {
  const backend = new MemoryStorageBackend();
  const validRuntime = await createRuntime(backend, "url_sha256:workbook-a");
  const validSessionId = await promptAndWaitForSave(validRuntime, "recoverable-sibling");
  validRuntime.controller.dispose();

  const corruptSessionId = "00000000-0000-4000-8000-000000000099";
  await backend.set("sessions", corruptSessionId, "{malformed-json");
  const restarted = await createRuntime(backend, "url_sha256:workbook-a");
  await restarted.settings.set(
    workbookLatestSessionKey("url_sha256:workbook-a"),
    corruptSessionId,
  );

  assert.equal(await restarted.controller.restoreLatestSession(), false);
  assert.equal(restarted.agent.state.messages.length, 0);
  const recoverable = await restarted.sessions.loadSession(validSessionId);
  assert.match(JSON.stringify(recoverable?.messages), /recoverable-sibling/);
});

void test("legacy pre-workbook-association session restores through global fallback", async () => {
  const backend = new MemoryStorageBackend();
  const legacyRuntime = await createRuntime(backend, null);
  const legacySessionId = await promptAndWaitForSave(legacyRuntime, "legacy-global-session");
  legacyRuntime.controller.dispose();

  const restarted = await createRuntime(backend, null);
  assert.equal(await restarted.settings.get(sessionWorkbookKey(legacySessionId)), null);
  assert.equal(await restarted.controller.restoreLatestSession(), true);
  assert.equal(restarted.controller.getSessionId(), legacySessionId);
  assert.match(transcriptText(restarted), /legacy-global-session/);
});

void test("disposed persistence drops later agent message events", async () => {
  const backend = new MemoryStorageBackend();
  const runtime = await createRuntime(backend, "url_sha256:workbook-a");
  const sessionId = await promptAndWaitForSave(runtime, "saved-before-dispose");
  runtime.controller.dispose();

  runtime.respond("reply:must-not-persist");
  await runtime.agent.prompt("must-not-persist");
  await new Promise<void>((resolve) => setTimeout(resolve, 10));

  const restarted = await createRuntime(backend, "url_sha256:workbook-a");
  assert.equal(await restarted.controller.restoreLatestSession(), true);
  assert.equal(restarted.controller.getSessionId(), sessionId);
  assert.match(transcriptText(restarted), /saved-before-dispose/);
  assert.doesNotMatch(transcriptText(restarted), /must-not-persist/);
});

void test("database identity and persisted store/key formats stay stable", async () => {
  const config = getAppStorageConfig();
  assert.equal(config.dbName, APP_STORAGE_DATABASE_NAME);
  assert.equal(APP_STORAGE_DATABASE_NAME, "pi-for-excel");
  assert.equal(config.version, APP_STORAGE_DATABASE_VERSION);
  assert.equal(APP_STORAGE_DATABASE_VERSION, 2);
  assert.deepEqual(
    config.stores.map((store) => store.name),
    ["settings", "provider-keys", "sessions", "sessions-metadata", "custom-providers", "model-catalogs"],
  );
  assert.deepEqual(config.stores.find((store) => store.name === "sessions"), {
    name: "sessions",
    keyPath: "id",
    indices: [{ name: "lastModified", keyPath: "lastModified" }],
  });

  const backend = new MemoryStorageBackend();
  const runtime = await createRuntime(backend, "url_sha256:workbook-a");
  const sessionId = await promptAndWaitForSave(runtime, "format-contract");
  const settingKeys = await runtime.settings.list();
  assert.deepEqual(settingKeys.sort(), [
    sessionWorkbookKey(sessionId),
    workbookLatestSessionKey("url_sha256:workbook-a"),
  ].sort());
  assert.equal(await backend.has("sessions", sessionId), true);
  assert.equal(await backend.has("sessions-metadata", sessionId), true);
});
