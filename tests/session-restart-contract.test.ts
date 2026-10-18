import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { DurableSessions } from "../src/durable/sessions.ts";
import { importWorkbookSessions } from "../src/durable/session-import.ts";
import { APP_STORAGE_DATABASE_NAME, APP_STORAGE_DATABASE_VERSION, getAppStorageConfig, initAppStorage } from "../src/storage/init-app-storage.ts";
import type { SessionData, StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import { partitionSessionIdsByWorkbook, sessionWorkbookKey, workbookLatestSessionKey } from "../src/workbook/session-association.ts";
import { createDurableWorkbookAssociationResolver } from "../src/workbook/durable-association.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, unknown>>();
  private store(name: string): Map<string, unknown> {
    let store = this.stores.get(name);
    if (!store) { store = new Map(); this.stores.set(name, store); }
    return store;
  }
  get<T = unknown>(name: string, key: string): Promise<T | null> {
    // Same typed persistence boundary as IndexedDB.
    return Promise.resolve(structuredClone(this.store(name).get(key) ?? null) as T | null);
  }
  set<T = unknown>(name: string, key: string, value: T): Promise<void> {
    this.store(name).set(key, structuredClone(value)); return Promise.resolve();
  }
  delete(name: string, key: string): Promise<void> { this.store(name).delete(key); return Promise.resolve(); }
  keys(name: string, prefix = ""): Promise<string[]> { return Promise.resolve([...this.store(name).keys()].filter((key) => key.startsWith(prefix))); }
  getAllFromIndex<T = unknown>(name: string): Promise<T[]> {
    // Tests enumerate snapshots by ID, not index order.
    return Promise.resolve(structuredClone([...this.store(name).values()]) as T[]);
  }
  clear(name: string): Promise<void> { this.store(name).clear(); return Promise.resolve(); }
  has(name: string, key: string): Promise<boolean> { return Promise.resolve(this.store(name).has(key)); }
  transaction<T>(_names: string[], _mode: "readonly" | "readwrite", operation: (tx: StorageTransaction) => Promise<T>): Promise<T> {
    return operation({ get: <V = unknown>(name: string, key: string) => this.get<V>(name, key),
      set: <V = unknown>(name: string, key: string, value: V) => this.set(name, key, value), delete: (name, key) => this.delete(name, key) });
  }
  getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> { return Promise.resolve({ usage: 0, quota: 0, percent: 0 }); }
  requestPersistence(): Promise<boolean> { return Promise.resolve(true); }
}

// Harness close normally owns storage close; reopening here models browser restart.
class ReopenableMemory extends MemoryStorage { override async close(): Promise<void> {} }
const model = fauxProvider().getModel();
function snapshot(id: string, text: string): SessionData {
  return { id, title: text, model, thinkingLevel: "off", createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-01T00:00:00.000Z",
    messages: [{ role: "user", content: text, timestamp: 1 }] };
}
async function saveLegacy(backend: StorageBackend, id: string, workbookId: string | null, text: string): Promise<void> {
  const app = initAppStorage(APP_STORAGE_DATABASE_NAME, backend);
  await backend.set("sessions", id, snapshot(id, text));
  await backend.set("sessions-metadata", id, { id, title: text, lastModified: "2026-01-01T00:00:00.000Z" });
  if (workbookId !== null) await app.settings.set(sessionWorkbookKey(id), workbookId);
  await app.settings.set(workbookLatestSessionKey(workbookId), id);
}
async function open(storage = new ReopenableMemory()): Promise<{ harness: Harness; sessions: DurableSessions }> {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  faux.setResponses([fauxAssistantMessage("native reply")]);
  models.setProvider(faux.provider);
  const harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  return { harness, sessions: new DurableSessions(harness, models, model) };
}
async function migrate(harness: Harness, backend: StorageBackend, workbookId: string | null): Promise<void> {
  const app = initAppStorage(APP_STORAGE_DATABASE_NAME, backend);
  await importWorkbookSessions(harness, app.sessions, app.settings, workbookId, context);
}

void test("workbook session partition fails closed when one association cannot be read", async () => {
  const settings = { get: (key: string): Promise<unknown> => key === sessionWorkbookKey("unreadable")
    ? Promise.reject(new Error("association read failed")) : Promise.resolve(key === sessionWorkbookKey("matching") ? "workbook-a" : "workbook-b"),
  set: (): Promise<void> => Promise.resolve() };
  await assert.rejects(partitionSessionIdsByWorkbook(settings, ["matching", "unreadable", "foreign"], "workbook-a"), /association read failed/u);
});

void test("transient migration failure preserves the session imported on the next restart", async () => {
  class ReadFailure extends MemoryStorageBackend {
    fail = true;
    override get<T = unknown>(name: string, key: string): Promise<T | null> {
      if (name === "settings" && key === sessionWorkbookKey("original") && this.fail) {
        this.fail = false; return Promise.reject(new Error("association read failed"));
      }
      return super.get<T>(name, key);
    }
  }
  const backend = new ReadFailure();
  await saveLegacy(backend, "original", "workbook-a", "persisted before failure");
  const storage = new ReopenableMemory();
  const first = await open(storage);
  try { await assert.rejects(migrate(first.harness, backend, "workbook-a"), /association read failed/u); }
  finally { await first.harness.close(context); }
  const restarted = await open(storage);
  try {
    await migrate(restarted.harness, backend, "workbook-a");
    assert.match(JSON.stringify((await restarted.sessions.loadSession("original"))?.messages), /persisted before failure/u);
    assert.equal((await restarted.sessions.getAllMetadata()).length, 1);
  } finally { await restarted.harness.close(context); }
});

void test("restart keeps two workbooks isolated and exposes no A session to B", async () => {
  const backend = new MemoryStorageBackend();
  await saveLegacy(backend, "a", "workbook-a", "alpha-only");
  await saveLegacy(backend, "b", "workbook-b", "beta-only");
  for (const [workbook, own, foreign, text] of [["workbook-a", "a", "b", "alpha-only"], ["workbook-b", "b", "a", "beta-only"]] as const) {
    const storage = new ReopenableMemory();
    const first = await open(storage);
    await migrate(first.harness, backend, workbook);
    const id = (await first.sessions.resolve(own))?.id;
    await first.harness.close(context);
    const restarted = await open(storage);
    try {
      await migrate(restarted.harness, backend, workbook);
      assert.equal((await restarted.sessions.resolve(own))?.id, id);
      assert.equal(await restarted.sessions.loadSession(foreign), null);
      assert.match(JSON.stringify((await restarted.sessions.loadSession(own))?.messages), new RegExp(text, "u"));
      assert.equal((await restarted.sessions.getAllMetadata()).length, 1);
    } finally { await restarted.harness.close(context); }
  }
});

void test("corrupt migration fails closed while a recoverable sibling remains", async () => {
  const backend = new MemoryStorageBackend();
  await saveLegacy(backend, "valid", "workbook-a", "recoverable-sibling");
  const runtime = await open();
  try {
    await migrate(runtime.harness, backend, "workbook-a");
    await backend.set("sessions", "corrupt", "{malformed-json");
    await initAppStorage(APP_STORAGE_DATABASE_NAME, backend).settings.set(sessionWorkbookKey("corrupt"), "workbook-a");
    await assert.rejects(migrate(runtime.harness, backend, "workbook-a"));
    assert.equal(await runtime.sessions.loadSession("corrupt"), null);
    assert.match(JSON.stringify((await runtime.sessions.loadSession("valid"))?.messages), /recoverable-sibling/u);
    assert.equal(await backend.get("sessions", "corrupt"), "{malformed-json");
  } finally { await runtime.harness.close(context); }
});

void test("failed workbook identity read imports no global session and writes no pointer", async () => {
  const backend = new MemoryStorageBackend();
  await saveLegacy(backend, "global", null, "foreign-global-session");
  const before = await backend.keys("settings");
  const runtime = await open();
  try {
    const resolve = createDurableWorkbookAssociationResolver({
      host: { kind: "browser", getWorkbookContext: () => Promise.reject(new Error("workbook identity read failed")) },
      settings: initAppStorage(APP_STORAGE_DATABASE_NAME, backend).settings,
    });
    await assert.rejects(resolve(), /workbook identity read failed/u);
    assert.deepEqual(await runtime.sessions.getAllMetadata(), []);
    assert.deepEqual(await backend.keys("settings"), before);
    await migrate(runtime.harness, backend, null);
    assert.match(JSON.stringify((await runtime.sessions.loadSession("global"))?.messages), /foreign-global-session/u);
  } finally { await runtime.harness.close(context); }
});

void test("legacy pre-association sessions survive native catalogue restart without latest-pointer dependence", async () => {
  const backend = new MemoryStorageBackend();
  await saveLegacy(backend, "legacy", null, "legacy-global-session");
  const storage = new ReopenableMemory();
  const first = await open(storage);
  await migrate(first.harness, backend, null);
  const id = (await first.sessions.resolve("legacy"))?.id;
  await first.harness.close(context);
  await backend.delete("settings", workbookLatestSessionKey(null));
  const restarted = await open(storage);
  try {
    assert.equal((await restarted.sessions.resolve("legacy"))?.id, id);
    assert.match(JSON.stringify((await restarted.sessions.loadSession("legacy"))?.messages), /legacy-global-session/u);
  } finally { await restarted.harness.close(context); }
});

void test("closed Harness rejects later input without changing persisted history", async () => {
  const storage = new ReopenableMemory();
  const first = await open(storage);
  const conversation = await first.sessions.create(model, "off");
  await conversation.submit({ type: "write", entry: { kind: "test.history", model: snapshot("saved", "saved-before-dispose").messages } }, context);
  await first.harness.close(context);
  await assert.rejects(conversation.submit({ type: "input", content: "must-not-persist" }, context));
  const restarted = await open(storage);
  try {
    const transcript = JSON.stringify((await restarted.sessions.loadSession(String(conversation.id)))?.messages);
    assert.match(transcript, /saved-before-dispose/u);
    assert.doesNotMatch(transcript, /must-not-persist/u);
  } finally { await restarted.harness.close(context); }
});

void test("native prompt and assistant reply are committed before restart", async () => {
  const storage = new ReopenableMemory();
  const first = await open(storage);
  const conversation = await first.sessions.create(model, "off");
  try {
    const submission = await conversation.submit({ type: "input", content: "persist this turn" }, context);
    assert.equal((await submission.wait(context)).status, "done");
  } finally { await first.harness.close(context); }
  const restarted = await open(storage);
  try {
    const saved = await restarted.sessions.loadSession(String(conversation.id));
    assert.deepEqual(saved?.messages.map((message) => message.role), ["user", "assistant"]);
    assert.match(JSON.stringify(saved?.messages), /persist this turn/u);
    assert.match(JSON.stringify(saved?.messages), /native reply/u);
    assert.equal(saved?.model.id, model.id);
  } finally { await restarted.harness.close(context); }
});

void test("persisted model catalogues retain valid input limits and drop malformed entries", async () => {
  const { modelCatalogs } = initAppStorage(APP_STORAGE_DATABASE_NAME, new MemoryStorageBackend());
  const inputLimits = { maxRequestBytes: 8_388_608, images: { maxPerRequest: 6,
    resize: { maxWidth: 900, maxHeight: 700, maxBytes: 1_048_576, jpegQuality: 75 } } };
  await modelCatalogs.write("faux", { models: [{ ...model, id: "valid", inputLimits },
    { ...model, id: "invalid", inputLimits: { images: { resize: { jpegQuality: 101 } } } }] });
  const restored = await modelCatalogs.read("faux");
  assert.ok(restored);
  assert.deepEqual(restored.models.map((entry) => entry.id), ["valid"]);
  assert.deepEqual(restored.models[0]?.inputLimits, inputLimits);
});

void test("database identity and legacy store/key formats stay stable through native import", async () => {
  const config = getAppStorageConfig();
  assert.equal(config.dbName, APP_STORAGE_DATABASE_NAME);
  assert.equal(APP_STORAGE_DATABASE_NAME, "pi-for-excel");
  assert.equal(config.version, APP_STORAGE_DATABASE_VERSION);
  assert.equal(APP_STORAGE_DATABASE_VERSION, 2);
  assert.deepEqual(config.stores.map((store) => store.name), ["settings", "provider-keys", "sessions", "sessions-metadata", "custom-providers", "model-catalogs"]);
  assert.deepEqual(config.stores.find((store) => store.name === "sessions"), { name: "sessions", keyPath: "id", indices: [{ name: "lastModified", keyPath: "lastModified" }] });
  const backend = new MemoryStorageBackend();
  await saveLegacy(backend, "format-contract", "url_sha256:workbook-a", "format-contract");
  const original = await backend.get("sessions", "format-contract");
  const keys = [sessionWorkbookKey("format-contract"), workbookLatestSessionKey("url_sha256:workbook-a")].sort();
  assert.deepEqual((await backend.keys("settings")).sort(), keys);
  const runtime = await open();
  try {
    await migrate(runtime.harness, backend, "url_sha256:workbook-a");
    assert.ok(await runtime.sessions.loadSession("format-contract"));
    assert.equal(await backend.has("sessions", "format-contract"), true);
    assert.equal(await backend.has("sessions-metadata", "format-contract"), true);
    assert.deepEqual(await backend.get("sessions", "format-contract"), original);
    assert.deepEqual((await backend.keys("settings")).sort(), keys);
  } finally { await runtime.harness.close(context); }
});
