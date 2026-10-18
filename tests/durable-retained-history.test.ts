import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import type { Api, Model } from "@earendil-works/pi-ai";
import { createRegistry, defineTask, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { createRetainedHistoryCallbacks } from "../src/durable/retained-history.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { MutationDoc, PendingMutationsDoc } from "../src/durable/runtime.ts";
import { createDurableWorkbookAssociationResolver, listRetainedDurableAssociations } from "../src/workbook/durable-association.ts";
import type { SettingsWriter } from "../src/storage/local/settings-store.ts";

const model: Model<Api> = { id: "test", name: "Test", provider: "test", api: "openai-completions", baseUrl: "",
  reasoning: false, input: ["text"], cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, contextWindow: 32000, maxTokens: 4000 };

class RetainedMemoryStorage extends MemoryStorage {
  readonly databaseId = "old-generation";
  readonly persistence = "best-effort" as const;
  checkpoint(): Promise<{ sequence: number; pruned: number }> { return Promise.resolve({ sequence: 0, pruned: 0 }); }
  override async close(): Promise<void> { /* Browser persistence survives closing its owning Harness. */ }
}

for (const state of ["complete", "task", "mutation", "unmigrated", "unrelated"] as const) void test(`retained WPS history: paused inspection, copy and deletion (${state})`, async () => {
  const pending = state !== "complete" && state !== "unrelated";
  const values = new Map<string, unknown>();
  const settings: SettingsWriter = { get: (key) => Promise.resolve(values.get(key)),
    set: (key, value) => { values.set(key, structuredClone(value)); return Promise.resolve(); } };
  const host = { kind: "wps" as const, getWorkbookContext: () => Promise.resolve({ workbookId: null, workbookName: "Book", source: "unknown" as const }) };
  const old = await createDurableWorkbookAssociationResolver({ host, settings, createId: () => "old" })();
  await old.rememberGeneration("old-generation");
  const current = await createDurableWorkbookAssociationResolver({ host, settings, createId: () => "new" })();
  await current.rememberGeneration("new-generation");
  const models = createModels();
  const storage = new RetainedMemoryStorage();
  const source = await Harness.open(storage, { models, registry: createRegistry() }, context);
  const sourceSessions = new DurableSessions(source, models, model);
  const conversation = await sourceSessions.create(model, "off", "Retained chat");
  const other = state === "unrelated" ? await sourceSessions.create(model, "off", "Pending chat") : null;
  // Simulate a prototype generation which wrote history before indexing existed.
  sourceSessions.dispose();
  await source.commit(async (tx) => {
    if (state !== "unmigrated") (await tx.doc(PendingMutationsDoc)).migrated = true;
    if (state === "mutation") (await tx.doc(PendingMutationsDoc)).ids.uncertain = true;
    await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "Old history", timestamp: 1 }] });
    if (state === "task") await tx.createTask(defineTask({ name: "must-not-run", version: 1, initial: () => ({ phase: "start" }),
      phases: { start: () => Promise.reject(new Error("Foreign work executed")) } }), {},
    { conversationId: conversation.id, ownership: { kind: "conversation" } });
    if (other) {
      const task = await tx.createTask(defineTask({ name: "must-not-run", version: 1, initial: () => ({ phase: "start" }),
        phases: { start: () => Promise.reject(new Error("Foreign work executed")) } }), {},
      { conversationId: other.id, ownership: { kind: "conversation" } });
      (await tx.doc(PendingMutationsDoc)).ids.bound = true;
      (await tx.doc(MutationDoc, "bound", null)).binding = { taskId: task, conversationId: other.id, rawCallId: "call" };
    }
  }, context);
  sourceSessions.dispose();
  await source.close(context);
  const target = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(target, models, model);
  let deletes = 0;
  let rejectDelete = true;
  const history = createRetainedHistoryCallbacks({ current, harness: target, sessions, models, defaultModel: model, settings,
    openStorage: (name, options) => {
      assert.equal(name, old.databaseName);
      assert.equal(options?.expectedDatabaseId, "old-generation");
      return Promise.resolve(storage);
    },
    deleteStorage: (options) => {
      assert.equal(options.name, old.databaseName);
      assert.equal(options.databaseId, "old-generation");
      assert.equal(options.authorization, "user-confirmed-obsolete-workbook");
      if (rejectDelete) return Promise.reject(new Error("Deletion blocked"));
      deletes++;
      return Promise.resolve();
    } });
  try {
    assert.equal((await history.listScopes()).find((row) => row.id === old.scopeId)?.identityDurable, false);
    const expected = [{ id: String(conversation.id), name: "Retained chat", completed: !pending },
      ...(other ? [{ id: String(other.id), name: "Pending chat", completed: false }] : [])];
    assert.deepEqual((await history.inspectSessions(old.scopeId)).sort((a, b) => a.id.localeCompare(b.id)), expected);
    assert.deepEqual((await history.inspectSessions(old.scopeId)).sort((a, b) => a.id.localeCompare(b.id)), expected);
    if (other) await assert.rejects(history.copyCompletedHistory(old.scopeId, String(other.id)), /Pending/);
    if (pending) {
      await assert.rejects(history.copyCompletedHistory(old.scopeId, String(conversation.id)), /Pending/);
      assert.equal((await sessions.getAllMetadata()).length, 0);
    } else {
      await history.copyCompletedHistory(old.scopeId, String(conversation.id));
      await history.copyCompletedHistory(old.scopeId, String(conversation.id));
      const copied = await sessions.getAllMetadata();
      assert.equal(copied.length, 1);
      assert.equal((await sessions.loadSession(copied[0]?.id ?? ""))?.messages.length, 1);
      assert.deepEqual((await target.inspect(context)).tasks, []);
      assert.deepEqual((await target.inspect(context)).submissions, []);
    }
    if (other) {
      const retained = await Harness.open(storage, { models, registry: createRegistry() }, context);
      assert.equal((await retained.inspect(context)).tasks.length, 1);
      assert.equal((await retained.inspect(context)).scheduling, "paused");
      assert.deepEqual((await retained.snapshot(PendingMutationsDoc, context))?.ids, { bound: true });
      assert.equal((await retained.snapshot(MutationDoc, "bound", context))?.binding?.conversationId, other.id);
      await retained.close(context);
    }
    assert.equal((await target.inspect(context)).scheduling, "paused");
    await assert.rejects(history.deleteGeneration(current.scopeId), /current/);
    assert.equal(deletes, 0);
    await assert.rejects(history.deleteGeneration(old.scopeId), /Deletion blocked/);
    assert.equal((await listRetainedDurableAssociations(settings)).length, 2);
    rejectDelete = false;
    await history.deleteGeneration(old.scopeId);
    assert.equal(deletes, 1);
    assert.equal((await listRetainedDurableAssociations(settings)).length, 1);
    assert.equal(await settings.get(`workbook.durable.v1.${old.scopeId}`), null);
    assert.equal((await history.listScopes())[0]?.current, true);
  } finally { sessions.dispose(); await target.close(context); }
});
