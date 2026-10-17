import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import type { Api, Model } from "@earendil-works/pi-ai";
import { CompactionTask, createRegistry, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { DurableSessions, SessionInfoDoc } from "../src/durable/sessions.ts";
import { importLegacySession, LegacyContextEntry, LegacySessionDoc } from "../src/durable/session-import.ts";
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
    assert.deepEqual(loaded, { ...session, id: String(conversation.id) });
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
      assert.deepEqual(presenter.current.messages, original.messages);
      assert.deepEqual((await catalogue.loadSession(original.id))?.messages, original.messages);
      const fork = await catalogue.fork(original.id);
      assert.deepEqual((await catalogue.loadSession(String(fork.id)))?.messages, original.messages);
      await conversation.reset(undefined, context);
      await presenter.refresh();
      assert.deepEqual(presenter.current.messages[0], { role: "archivedMessages", timestamp: 0,
        archivedChatMessageCount: 1, archivedMessages: original.messages });
      assert.deepEqual((await catalogue.loadSession(original.id))?.messages, original.messages);
    } finally { await presenter.dispose(); }
  } finally { await harness.close(context); }
});

void test("presenter reads all committed history, prompt patches and compaction measurements", async () => {
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
      const archive = presenter.current.messages[0];
      assert.equal(archive?.role, "archivedMessages");
      if (archive?.role === "archivedMessages") assert.equal(archive.archivedMessages.length, 215);
      assert.equal(presenter.current.systemPrompt, "one\n\ntwo");
      assert.deepEqual(presenter.current.messages.at(-1), { role: "compactionSummary", summary: "summary", tokensBefore: 12345, timestamp: 218 });
    } finally { await presenter.dispose(); }
    await harness.commit(async (tx) => {
      (await tx.doc(LegacySessionDoc, conversation.id)).originalSessionJson = JSON.stringify({ model: { ...model, contextWindow: "invalid" } });
      (await tx.doc(SessionInfoDoc, conversation.id)).fallbackModel = null;
    }, context);
    assert.deepEqual((await catalogue.loadSession(String(conversation.id)))?.model, model);
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
          const archive = snapshot.messages[0];
          if (archive?.role === "archivedMessages" && JSON.stringify(archive).includes("buffered original")) resolve();
        });
      });
      const archive = presenter.current.messages[0];
      assert.equal(archive?.role, "archivedMessages");
      if (archive?.role === "archivedMessages") assert.deepEqual(archive.archivedMessages,
        [{ role: "user", content: "buffered original", timestamp: 1 }]);
    } finally { await presenter.dispose(); }
  } finally { await harness.close(context); }
});
