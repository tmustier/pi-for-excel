import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { AgentDoc, createRegistry, Harness, MemoryStorage, ProviderDoc } from "@earendil-works/pi-durable";
import { importLegacySession, LegacyContextEntry, LegacySessionDoc } from "../src/durable/session-import.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import type { SessionData } from "../src/storage/local/types.ts";

const context = BACKGROUND_CONTEXT;

function fixture(): SessionData {
  return {
    id: "11111111-1111-4111-8111-111111111111",
    title: "Original title", createdAt: "2026-01-01T00:00:00Z", lastModified: "2026-02-01T00:00:00Z",
    thinkingLevel: "high",
    model: {
      id: "import-model", name: "Import model", provider: "test", api: "openai-completions",
      baseUrl: "https://example.com", reasoning: true, input: ["text"],
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 }, contextWindow: 32000, maxTokens: 4096,
    },
    messages: [
      { role: "archivedMessages", timestamp: 3, archivedChatMessageCount: 1, archivedMessages: [
        { role: "user", content: "ARCHIVE ONLY", timestamp: 1 },
        { role: "toolResult", toolName: "read_range", toolCallId: "old", content: [{ type: "text", text: "old result" }], details: { cells: [1, 2] }, isError: false, timestamp: 2 },
      ] },
      { role: "compactionSummary", summary: "COMPACTED SUMMARY", tokensBefore: 12345, timestamp: 4 },
      { role: "user", content: "Continue", timestamp: 5 },
    ],
  };
}

void test("imports real Durable context and originals atomically, with concurrent duplicate protection", async () => {
  const harness = await Harness.open(new MemoryStorage(), { models: createModels(), registry: createRegistry() }, context);
  try {
    const session = fixture();
    const before = structuredClone(session);
    const [conversation, duplicate] = await Promise.all([
      importLegacySession(harness, session, context), importLegacySession(harness, session, context),
    ]);
    assert.equal(conversation.id, duplicate.id);
    const view = await conversation.context(context);
    assert.deepEqual(view.messages, createConvertToLlm({ getContextWindow: () => 32000 })(session.messages));
    assert.equal(view.entries.length, 1);
    assert.ok(view.entries.every(LegacyContextEntry.is));
    assert.match(JSON.stringify(view.messages), /COMPACTED SUMMARY/);
    assert.doesNotMatch(JSON.stringify(view.messages), /ARCHIVE ONLY/);
    const original = await harness.snapshot(LegacySessionDoc, conversation.id, context);
    assert.deepEqual(original, {
      legacySessionId: session.id, title: session.title, createdAt: session.createdAt,
      lastModified: session.lastModified, originalSessionJson: JSON.stringify(session),
    });
    assert.deepEqual(await harness.snapshot(ProviderDoc, conversation.id, context), { sessionId: session.id });
    const agent = await harness.snapshot(AgentDoc, conversation.id, context);
    assert.deepEqual(agent?.model, { provider: "test", modelId: "import-model" });
    assert.equal(agent?.thinkingLevel, "high");
    assert.deepEqual(session, before);
    const repeated = await importLegacySession(harness, { ...session, title: "Do not overwrite", messages: [] }, context);
    assert.equal(repeated.id, conversation.id);
    assert.deepEqual(await repeated.context(context), view);
    assert.deepEqual(await harness.snapshot(LegacySessionDoc, repeated.id, context), original);
  } finally {
    await harness.close(context);
  }
});

void test("failed creation rolls back and permits a complete retry", async () => {
  let fail = true;
  const storage = new MemoryStorage();
  const harness = await Harness.open(storage, {
    models: createModels(), registry: createRegistry(),
    conversationCreated: () => { if (fail) throw new Error("creation rejected"); },
  }, context);
  try {
    await assert.rejects(importLegacySession(harness, fixture(), context), /creation rejected/);
    const page = await storage.scanConversations({}, 100, undefined, context);
    assert.equal(page.items.length, 0);
    fail = false;
    const conversation = await importLegacySession(harness, fixture(), context);
    assert.equal((await conversation.context(context)).messages.length, 2);
  } finally {
    await harness.close(context);
  }
});
