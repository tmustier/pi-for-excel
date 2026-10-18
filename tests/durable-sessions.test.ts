import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import type { Api, Model } from "@earendil-works/pi-ai";
import { CompactionTask, createRegistry, Harness, LiveDoc, MemoryStorage, type DocumentId } from "@earendil-works/pi-durable";
import { DurableSessions, SessionInfoDoc } from "../src/durable/sessions.ts";
import { importLegacySession, LegacyContextEntry, LegacySessionDoc, LegacyImportInfoDoc } from "../src/durable/session-import.ts";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import type { SessionData } from "../src/storage/local/types.ts";

const model: Model<Api> = {
  id: "unlisted", name: "Original model", provider: "test", api: "openai-completions", baseUrl: "https://example.com",
  reasoning: true, input: ["text"], cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, contextWindow: 32000, maxTokens: 4000,
};

void test("catalogue preserves imported IDs, originals, archived export, forks and soft deletion", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, { ...model, id: "different-fallback" });
  const session: SessionData = {
    id: "old-tab-id", title: "Original title", createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-02-01T00:00:00.000Z",
    model, thinkingLevel: "high", messages: [
      { role: "archivedMessages", timestamp: 3, archivedChatMessageCount: 1,
        archivedMessages: [{ role: "user", content: "Archived forever", timestamp: 1 },
          { role: "assistant", content: [{ type: "text", text: "Old answer" }], timestamp: 2,
            api: "openai-completions", provider: "test", model: model.id, stopReason: "stop",
            usage: { input: 10, output: 2, cacheRead: 0, cacheWrite: 0, totalTokens: 12,
              cost: { input: 1, output: 2, cacheRead: 0, cacheWrite: 0, total: 3 } } }] },
      { role: "user", content: "Active", timestamp: 4 },
    ],
  };
  try {
    const conversation = await importLegacySession(harness, session, context);
    const loaded = await catalogue.loadSession(session.id);
    assert.deepEqual(loaded, { ...session, model: { ...model, baseUrl: "" }, id: String(conversation.id) });
    assert.equal((await catalogue.metadata(session.id))?.usage.totalTokens, 12);
    assert.equal((await catalogue.metadata(session.id))?.usage.cost.total, 3);
    await catalogue.renameSession(session.id, "Renamed");
    assert.equal((await catalogue.metadata(session.id))?.title, "Renamed");
    assert.equal((await catalogue.loadSession(session.id))?.createdAt, session.createdAt);
    const fork = await catalogue.fork(session.id, undefined, "Branch");
    assert.equal((await harness.commit((tx) => tx.conversation(fork.id), context))?.parent?.conversationId, conversation.id);
    assert.deepEqual((await catalogue.loadSession(String(fork.id)))?.messages, session.messages);
    assert.equal((await catalogue.metadata(String(fork.id)))?.usage.totalTokens, 12);
    assert.equal((await catalogue.resolve(session.id))?.id, conversation.id);
    await catalogue.delete(session.id);
    assert.equal(await catalogue.loadSession(session.id), null);
    assert.equal(await catalogue.metadata(session.id), null);
    assert.equal((await catalogue.getAllMetadata()).length, 1);
    assert.equal((await catalogue.resolve(session.id))?.id, conversation.id);
    assert.equal((await harness.snapshot(SessionInfoDoc, conversation.id, context))?.deleted, true);
    assert.equal((await harness.snapshot(LegacySessionDoc, conversation.id, context))?.originalSessionJson, JSON.stringify(session));
    assert.ok(await catalogue.loadSession(String(fork.id)));
  } finally { await harness.close(context); }
});

void test("catalogue pages all conversations and exports all history after reset", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const conversation = await catalogue.create(model, "off");
    await harness.commit(async (tx) => {
      for (let index = 0; index < 215; index++) {
        await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: `message-${index}`, timestamp: index + 1 }] });
      }
    }, context);
    await conversation.reset("Fresh start", context);
    assert.equal((await catalogue.loadSession(String(conversation.id)))?.messages.length, 216);
    assert.equal((await catalogue.metadata(String(conversation.id)))?.title, "message-0");
    for (let index = 0; index < 204; index++) await catalogue.create(model, "off", `Session ${index}`);
    assert.equal((await catalogue.getAllMetadata()).length, 205);
    assert.equal(await catalogue.loadSession("missing"), null);
    await assert.rejects(catalogue.renameSession("missing", "Nope"), /not found/);
  } finally { await harness.close(context); }
});

void test("prototype originals survive projection, reset, export and native forks", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  const original: SessionData = { id: "prototype", title: "Prototype", model, thinkingLevel: "off",
    createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-01T00:00:00.000Z",
    messages: [{ role: "archivedMessages", archivedChatMessageCount: 1, timestamp: 2,
      archivedMessages: [{ role: "user", content: "original archived payload", timestamp: 1 }] },
    { role: "user", content: "not the shaped contribution", timestamp: 3 }] };
  try {
    const conversation = await catalogue.create(model, "off");
    await harness.commit(async (tx) => {
      Object.assign(await tx.doc(LegacySessionDoc, conversation.id), { legacySessionId: original.id,
        title: original.title, createdAt: original.createdAt, lastModified: original.lastModified,
        originalSessionJson: JSON.stringify(original) });
      await tx.appendEntry(LegacyContextEntry, conversation.id, { model: [{ role: "user", content: "shaped", timestamp: 3 }] });
    }, context);
    const presenter = await ConversationPresenter.open(harness, conversation, models, model);
    try {
      assert.deepEqual(presenter.current.messages, [{ role: "user", content: "shaped", timestamp: 3 }]);
      assert.deepEqual((await presenter.archivePage()).messages, [original.messages[0]?.role === "archivedMessages"
        ? original.messages[0].archivedMessages[0] : undefined, original.messages[1]]);
      assert.deepEqual((await catalogue.loadSession(original.id))?.messages, original.messages);
      const fork = await catalogue.fork(original.id);
      assert.deepEqual((await catalogue.loadSession(String(fork.id)))?.messages, original.messages);
      await conversation.reset(undefined, context);
      await presenter.refresh();
      assert.equal(presenter.current.hasArchive, true);
      assert.equal((await presenter.archivePage()).messages.length, 2);
      assert.deepEqual((await catalogue.loadSession(original.id))?.messages, original.messages);
    } finally { await presenter.dispose(); }
  } finally { await harness.close(context); }
});

void test("presenter pages archived history explicitly and projects prompt patches and compaction measurements", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const conversation = await catalogue.create(model, "off");
    await harness.commit(async (tx) => {
      for (let index = 0; index < 215; index++) await tx.appendEntry(conversation.id, {
        kind: "test.message", model: [{ role: "user", content: `message-${index}`, timestamp: index }] });
    }, context);
    await conversation.reset(undefined, context);
    await harness.commit(async (tx) => {
      await tx.appendEntry(conversation.id, { kind: "pi.system", model: [{ role: "system", content: "",
        sections: { first: "one", removed: "old" }, timestamp: 216 }] });
      await tx.appendEntry(conversation.id, { kind: "pi.system", model: [{ role: "system", content: "",
        sections: { removed: null, second: "two" }, timestamp: 217 }] });
      await tx.appendEntry(conversation.id, { kind: "pi.compaction", data: { reason: "manual", tokensBefore: 12345 },
        model: [{ role: "user", content: "summary", timestamp: 218 }] });
    }, context);
    const presenter = await ConversationPresenter.open(harness, conversation, models, model);
    try {
      assert.equal(presenter.current.hasArchive, true);
      const archived = [];
      let cursor;
      do {
        const page = await presenter.archivePage(cursor);
        archived.push(...page.messages);
        cursor = page.next;
      } while (cursor !== undefined);
      assert.equal(archived.length, 215);
      assert.equal(presenter.current.systemPrompt, "one\n\ntwo");
      assert.deepEqual(presenter.current.messages.at(-1), { role: "compactionSummary", summary: "summary", tokensBefore: 12345, timestamp: 218 });
    } finally { await presenter.dispose(); }
    await harness.commit(async (tx) => {
      (await tx.doc(LegacySessionDoc, conversation.id)).originalSessionJson = JSON.stringify({ model: { ...model, contextWindow: "invalid" } });
      (await tx.doc(SessionInfoDoc, conversation.id)).fallbackModel = null;
    }, context);
    assert.deepEqual((await catalogue.loadSession(String(conversation.id)))?.model, { ...model, baseUrl: "" });
  } finally { await harness.close(context); }
});

void test("presenter cannot replace a newer refresh with a delayed older context", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const conversation = await catalogue.create(model, "off");
    let release: (() => void) | undefined;
    let captured: (() => void) | undefined;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const ready = new Promise<void>((resolve) => { captured = resolve; });
    let delay = false;
    const delayedConversation: typeof conversation = { ...conversation, id: conversation.id,
      agent: conversation.agent.bind(conversation), entries: conversation.entries.bind(conversation),
      context: async (readContext) => {
      const view = await conversation.context(readContext);
      if (delay) {
        delay = false;
        captured?.();
        await gate;
      }
      return view;
    } };
    const presenter = await ConversationPresenter.open(harness, delayedConversation, models, model);
    try {
      delay = true;
      const older = presenter.refresh();
      await ready;
      await harness.commit(async (tx) => {
        await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "newest", timestamp: 1 }] });
      }, context);
      await presenter.refresh();
      assert.equal(presenter.current.messages.at(-1)?.role, "user");
      release?.();
      await older;
      assert.deepEqual(presenter.current.messages.at(-1), { role: "user", content: "newest", timestamp: 1 });
    } finally { release?.(); await presenter.dispose(); }
  } finally { await harness.close(context); }
});

void test("task-owned subagents are not user catalogue sessions", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const parent = await catalogue.create(model, "off", "User session");
    const child = await harness.commit(async (tx) => {
      const taskId = await tx.createTask(CompactionTask, { reason: "manual" }, {
        ownership: { kind: "conversation" }, conversationId: parent.id,
      });
      return tx.createConversation({ ownership: { kind: "task", taskId } });
    }, context);
    assert.equal(await catalogue.resolve(String(child.id)), null);
    assert.equal(await catalogue.loadSession(String(child.id)), null);
    assert.deepEqual((await catalogue.getAllMetadata()).map((session) => session.id), [String(parent.id)]);
  } finally { await harness.close(context); }
});

void test("fork after compaction includes the newer head marker and all originals", async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const conversation = await catalogue.create(model, "off");
    await harness.commit(async (tx) => {
      await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "old", timestamp: 1 }] });
      const kept = await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "kept", timestamp: 2 }] });
      await tx.appendEntry(conversation.id, { kind: "pi.compaction", head: kept.id,
        model: [{ role: "user", content: "summary", timestamp: 3 }], data: { reason: "manual", tokensBefore: 42 } });
    }, context);
    const fork = await catalogue.fork(String(conversation.id));
    assert.deepEqual((await catalogue.loadSession(String(fork.id)))?.messages,
      (await catalogue.loadSession(String(conversation.id)))?.messages);
    assert.deepEqual((await fork.context(context)).messages, (await conversation.context(context)).messages);
  } finally { await harness.close(context); }
});

void test("commits buffered before watch start remain in the UI archive", { timeout: 5000 }, async () => {
  const models = createModels();
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  try {
    const conversation = await catalogue.create(model, "off");
    let first = true;
    const duringOpen: typeof conversation = { ...conversation, id: conversation.id,
      agent: conversation.agent.bind(conversation), entries: conversation.entries.bind(conversation),
      context: async (readContext) => {
        const view = await conversation.context(readContext);
        if (first) {
          first = false;
          await harness.commit(async (tx) => {
            await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "buffered original", timestamp: 1 }] });
          }, context);
          await conversation.reset("new context", context);
        }
        return view;
      } };
    const presenter = await ConversationPresenter.open(harness, duringOpen, models, model);
    try {
      await new Promise<void>((resolve) => {
        presenter.subscribe((snapshot) => {
          if (snapshot.messages.some((message) => message.role === "user" && message.content === "new context")) resolve();
        });
      });
      assert.deepEqual((await presenter.archivePage()).messages,
        [{ role: "user", content: "buffered original", timestamp: 1 }]);
    } finally { await presenter.dispose(); }
  } finally { await harness.close(context); }
});

class InstrumentedStorage extends MemoryStorage {
  persistedWrites: string[] = [];
  entryQueries: { limit: number; minEntryId: number | undefined }[] = [];
  archivedEntriesRead = 0;
  originalDocumentsRead = 0;
  private legacyIds = new Set<DocumentId>();

  override async commit(...args: Parameters<MemoryStorage["commit"]>) {
    const seq = await super.commit(...args);
    this.persistedWrites.push(JSON.stringify(args[0]));
    return seq;
  }

  override async scanEntries(...args: Parameters<MemoryStorage["scanEntries"]>) {
    this.entryQueries.push({ limit: args[1], minEntryId: args[0].minEntryId });
    const page = await super.scanEntries(...args);
    this.archivedEntriesRead += page.items.filter((entry) => entry.model?.some((message) =>
      message.role === "user" && typeof message.content === "string" && message.content.startsWith("cold-"))).length;
    return page;
  }

  override async findDocument(...args: Parameters<MemoryStorage["findDocument"]>) {
    const record = await super.findDocument(...args);
    if (args[0].kind === LegacySessionDoc.definition.kind && record) this.legacyIds.add(record.id);
    return record;
  }

  override async document(...args: Parameters<MemoryStorage["document"]>) {
    if (this.legacyIds.has(args[0])) this.originalDocumentsRead++;
    return super.document(...args);
  }

  // A persistent test backend: reopen a fresh Harness over the same stored records.
  override async close(): Promise<void> {}
}

void test("cold 1200-entry history stays cold through metadata, tab headers and 40 streaming partials; explicit archive/export remain complete", { timeout: 10000 }, async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  let harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  let catalogue = new DurableSessions(harness, models, model);
  const imported: SessionData = { id: "cold-import", title: "Cold import", model, thinkingLevel: "off",
    createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-02T00:00:00.000Z",
    messages: Array.from({ length: 1200 }, (_, index) => ({ role: "user", content: `cold-${index}`, timestamp: index })) };
  const original = await importLegacySession(harness, imported, context);
  await original.reset("Active", context);
  await catalogue.metadata(String(original.id));
  catalogue.dispose();
  await harness.close(context);
  harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  catalogue = new DurableSessions(harness, models, model);
  const conversation = await catalogue.resolve(String(original.id));
  assert.ok(conversation);
  storage.archivedEntriesRead = storage.originalDocumentsRead = 0;
  await catalogue.reconcileMetadata();
  const presenter = await ConversationPresenter.open(harness, conversation, models, model);
  try {
    for (let index = 0; index < 4; index++) {
      assert.equal((await catalogue.getAllMetadata())[0]?.messageCount, 1201);
      assert.equal((await catalogue.header(String(conversation.id)))?.title, "Cold import");
    }
    const projection = presenter.current.messages;
    for (let index = 0; index < 40; index++) {
      const observed = new Promise<void>((resolve) => {
        const unsubscribe = presenter.onEvent((event) => {
          if (event.type === "message_start" || event.type === "message_update") { unsubscribe(); resolve(); }
        });
      });
      await harness.commit(async (tx) => {
        (await tx.doc(LiveDoc, conversation.id)).generation = { attempt: 0, message: {
          role: "assistant", api: "openai-completions", provider: "test", model: model.id,
          content: [{ type: "text", text: `partial-${index}` }], stopReason: "stop", timestamp: 2000,
          usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
            cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
        } };
      }, context);
      await observed;
      assert.equal(presenter.current.messages, projection, "partial updates retain the active projection");
      assert.deepEqual(presenter.current.streamingMessage?.content, [{ type: "text", text: `partial-${index}` }]);
    }
    assert.equal(storage.archivedEntriesRead, 0);
    assert.equal(storage.originalDocumentsRead, 0);
    const archived = [];
    let cursor;
    do {
      const page = await presenter.archivePage(cursor);
      assert.ok(page.messages.length <= 100);
      archived.push(...page.messages);
      cursor = page.next;
    } while (cursor);
    assert.deepEqual(archived, imported.messages);
    assert.equal((await catalogue.loadSession(String(conversation.id)))?.messages.length, 1201);
    assert.ok(storage.archivedEntriesRead >= 1200);
  } finally { await presenter.dispose(); catalogue.dispose(); await harness.close(context); }
});

void test("prototype originals migrate once into persistent 100-message pages without replacing the backup", async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  let harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  let catalogue = new DurableSessions(harness, models, model);
  const conversation = await catalogue.create(model, "off");
  const original: SessionData = { id: "prototype-large", title: "Prototype", model, thinkingLevel: "off",
    createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-02T00:00:00.000Z",
    messages: Array.from({ length: 550 }, (_, index) => ({ role: "user", content: `cold-${index}`, timestamp: index })) };
  const backup = JSON.stringify(original);
  await harness.commit(async (tx) => {
    Object.assign(await tx.doc(LegacySessionDoc, conversation.id), { legacySessionId: original.id,
      title: original.title, createdAt: original.createdAt, lastModified: original.lastModified, originalSessionJson: backup });
    await tx.appendEntry(LegacyContextEntry, conversation.id, { model: [{ role: "user", content: "shaped", timestamp: 551 }] });
  }, context);
  assert.equal((await catalogue.metadata(String(conversation.id)))?.messageCount, 550);
  catalogue.dispose();
  await harness.close(context);
  harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  catalogue = new DurableSessions(harness, models, model);
  const reopened = await catalogue.resolve(String(conversation.id));
  assert.ok(reopened);
  storage.originalDocumentsRead = storage.archivedEntriesRead = 0;
  const presenter = await ConversationPresenter.open(harness, reopened, models, model);
  try {
    for (let index = 0; index < 5; index++) {
      await presenter.refresh(false);
      assert.equal((await catalogue.getAllMetadata())[0]?.messageCount, 550);
    }
    const archived = [];
    let cursor;
    do {
      const page = await presenter.archivePage(cursor);
      assert.ok(page.messages.length <= 100);
      archived.push(...page.messages);
      cursor = page.next;
    } while (cursor);
    assert.deepEqual(archived, original.messages);
    assert.equal(storage.originalDocumentsRead, 0, "paging migrated originals never hydrates the backup");
    assert.equal(storage.archivedEntriesRead, 0);
    assert.deepEqual((await catalogue.loadSession(String(conversation.id)))?.messages, original.messages);
    assert.equal((await harness.snapshot(LegacySessionDoc, conversation.id, context))?.originalSessionJson, backup);
  } finally { await presenter.dispose(); catalogue.dispose(); await harness.close(context); }
});

void test("imported archive buckets are stored out of active entries and paged only on request, with lossless export", async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  const harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  const catalogue = new DurableSessions(harness, models, model);
  const archived: SessionData["messages"] = Array.from({ length: 350 }, (_, index) =>
    ({ role: "user", content: `cold-${index}`, timestamp: index }));
  const session: SessionData = { id: "archive-bucket", title: "Bucket", model, thinkingLevel: "off",
    createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-02T00:00:00.000Z",
    messages: [{ role: "archivedMessages", archivedMessages: archived, archivedChatMessageCount: 350, timestamp: 351 },
      { role: "user", content: "Active only", timestamp: 352 }] };
  const conversation = await importLegacySession(harness, session, context);
  const presenter = await ConversationPresenter.open(harness, conversation, models, model);
  try {
    assert.doesNotMatch(JSON.stringify((await conversation.context(context)).entries), /cold-/);
    assert.deepEqual(presenter.current.messages, [session.messages[1]]);
    assert.equal(presenter.current.hasArchive, true);
    storage.originalDocumentsRead = 0;
    const messages = [];
    let cursor;
    do {
      const page = await presenter.archivePage(cursor);
      assert.ok(page.messages.length <= 100);
      messages.push(...page.messages);
      cursor = page.next;
    } while (cursor);
    assert.deepEqual(messages, archived);
    assert.equal(storage.originalDocumentsRead, 0);
    assert.equal((await catalogue.metadata(session.id))?.messageCount, 351);
    assert.deepEqual((await catalogue.loadSession(session.id))?.messages, session.messages);
  } finally { await presenter.dispose(); catalogue.dispose(); await harness.close(context); }
});

void test("explicit index migration and crash-tail recovery are paged and idempotent; normal listing never backfills history", async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  const harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  const conversation = await harness.createConversation({ ownership: { kind: "ownerless" } }, context);
  await harness.commit(async (tx) => {
    for (let index = 0; index < 450; index++) await tx.appendEntry(conversation.id, {
      kind: "test.message", model: [{ role: "user", content: `cold-${index}`, timestamp: index }] });
  }, context);
  await conversation.reset("Active", context);
  let catalogue = new DurableSessions(harness, models, model);
  try {
    await catalogue.reconcileMetadata();
    assert.equal((await catalogue.metadata(String(conversation.id)))?.messageCount, 451);
    storage.archivedEntriesRead = 0;
    await catalogue.reconcileMetadata();
    assert.equal((await catalogue.getAllMetadata())[0]?.title, "cold-0");
    assert.equal(storage.archivedEntriesRead, 0);
    catalogue.dispose();
    // Simulate a committed tail whose asynchronous secondary index never ran.
    await harness.commit(async (tx) => {
      await tx.appendEntry(conversation.id, { kind: "test.message", model: [{ role: "user", content: "Recovered tail", timestamp: 500 }] });
    }, context);
    catalogue = new DurableSessions(harness, models, model);
    await catalogue.reconcileMetadata();
    assert.equal((await catalogue.metadata(String(conversation.id)))?.messageCount, 452);
    assert.equal(storage.archivedEntriesRead, 0);
  } finally { catalogue.dispose(); await harness.close(context); }
});

void test("new session docs and exports omit model credentials; imported backup alone stays lossless", async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  const harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  const secretModel = { ...model, headers: { Authorization: "Bearer SENTINEL_AUTH", "x-token": "SENTINEL_HEADER" },
    baseUrl: "https://example.com?token=SENTINEL_URL" };
  const catalogue = new DurableSessions(harness, models, secretModel);
  try {
    const native = await catalogue.create(secretModel, "off");
    const safe = JSON.stringify(await harness.snapshot(SessionInfoDoc, native.id, context));
    assert.doesNotMatch(safe, /SENTINEL|Authorization|headers|https:/);
    assert.doesNotMatch(JSON.stringify(await catalogue.loadSession(String(native.id))), /SENTINEL|Authorization|headers/);
    await harness.commit(async (tx) => {
      await tx.appendEntry(native.id, { kind: "test.message", model: [{ role: "user", content: "First", timestamp: 1 }] });
    }, context);
    const fork = await catalogue.fork(String(native.id), secretModel);
    assert.doesNotMatch(JSON.stringify(await harness.snapshot(SessionInfoDoc, fork.id, context)), /SENTINEL|Authorization|headers/);
    assert.doesNotMatch(storage.persistedWrites.join("\n"), /SENTINEL|Authorization|headers/, "all native store writes are credential-free");
    const imported: SessionData = { id: "secret-import", title: "Imported", model: secretModel, thinkingLevel: "off",
      createdAt: "2026-01-01T00:00:00.000Z", lastModified: "2026-01-01T00:00:00.000Z", messages: [] };
    const legacy = await importLegacySession(harness, imported, context);
    assert.doesNotMatch(JSON.stringify(await harness.snapshot(LegacyImportInfoDoc, legacy.id, context)), /SENTINEL|Authorization|headers/);
    assert.doesNotMatch(JSON.stringify(await catalogue.loadSession(imported.id)), /SENTINEL|Authorization|headers/);
    assert.equal((await harness.snapshot(LegacySessionDoc, legacy.id, context))?.originalSessionJson, JSON.stringify(imported));
    assert.equal(secretModel.headers.Authorization, "Bearer SENTINEL_AUTH", "runtime input was not mutated");
  } finally { catalogue.dispose(); await harness.close(context); }
});

void test("empty drafts restore but stay out of history; first user before assistant recovers with bounded queries", async () => {
  const models = createModels();
  const storage = new InstrumentedStorage();
  let harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  let catalogue = new DurableSessions(harness, models, model);
  const empty = await catalogue.create(model, "off");
  const named = await catalogue.create(model, "off", "Named draft");
  const pending = await catalogue.create(model, "off");
  catalogue.dispose();
  // Crash window: user entry committed, no assistant and no secondary index.
  await harness.commit(async (tx) => {
    await tx.appendEntry(pending.id, { kind: "test.message", model: [{ role: "user", content: "Pending first prompt", timestamp: 1 }] });
  }, context);
  await harness.close(context);
  harness = await Harness.open(storage, { models, registry: createRegistry() }, context);
  catalogue = new DurableSessions(harness, models, model);
  try {
    storage.entryQueries = [];
    await catalogue.reconcileMetadata();
    assert.ok(storage.entryQueries.length > 0);
    assert.ok(storage.entryQueries.every((query) => query.limit <= 200 && query.minEntryId !== undefined));
    storage.entryQueries = [];
    assert.ok(await catalogue.header(String(empty.id)), "unused tab is still durable");
    const history = await catalogue.getAllMetadata();
    assert.deepEqual(new Set(history.map((item) => item.id)), new Set([String(named.id), String(pending.id)]));
    const first = history.find((item) => item.id === String(pending.id));
    assert.equal(first?.title, "Pending first prompt");
    assert.equal(first?.messageCount, 1);
    assert.equal(first?.preview, "Pending first prompt");
    assert.equal(storage.entryQueries.length, 0, "listing and restoration headers never query transcript entries");
    assert.deepEqual((await catalogue.loadSession(String(pending.id)))?.messages,
      [{ role: "user", content: "Pending first prompt", timestamp: 1 }]);
  } finally { catalogue.dispose(); await harness.close(context); }
});
