import assert from "node:assert/strict";
import { test } from "node:test";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, defineExtension, Harness, MemoryStorage, section } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { importLegacySession } from "../src/durable/session-import.ts";
import { DurableSessions } from "../src/durable/sessions.ts";

void test("internal Durable prompt entries remain model context, not chat messages, exports or message counts", async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: Infinity });
  models.setProvider(faux.provider);
  faux.setResponses([fauxAssistantMessage("Visible response")]);
  const registry = createRegistry();
  const extension = defineExtension({ name: "projection-test", sections: [
    section("internal-policy", () => "Internal policy only", { tag: false }),
  ] });
  registry.install(extension);
  const harness = await Harness.open(new MemoryStorage(), { models, registry, settings: { compaction: { enabled: false } } }, context);
  const sessions = new DurableSessions(harness, models, faux.getModel());
  let view: ConversationPresenter | undefined;
  try {
    const conversation = await sessions.create(faux.getModel(), "off");
    await conversation.configure({ extensions: [extension] }, context);
    view = await ConversationPresenter.open(harness, conversation, models, faux.getModel());
    const submitted = await conversation.submit({ type: "input", content: "Hello" }, context);
    assert.equal((await submitted.wait(context)).status, "done");
    await view.refresh();
    const exported = await sessions.loadSession(String(conversation.id));
    assert.ok(exported);
    assert.deepEqual(view.current.messages.map((message) => message.role), ["user", "assistant"]);
    assert.deepEqual(exported.messages.map((message) => message.role), ["user", "assistant"]);
    assert.equal((await sessions.metadata(String(conversation.id)))?.messageCount, 2);
    assert.match(view.current.systemPrompt ?? "", /Internal policy only/);
    assert.ok((await conversation.context(context)).messages.some((message) => message.role === "system"), "model context still includes its policy");
  } finally {
    await view?.dispose();
    sessions.dispose();
    await harness.close(context);
  }
});

void test("explicit system messages in original imported snapshots remain losslessly exportable", async () => {
  const models = createModels();
  const faux = fauxProvider();
  models.setProvider(faux.provider);
  const harness = await Harness.open(new MemoryStorage(), { models, registry: createRegistry() }, context);
  const sessions = new DurableSessions(harness, models, faux.getModel());
  try {
    const original = { role: "system" as const, content: "User-authored original context", timestamp: 1 };
    const conversation = await importLegacySession(harness, {
      id: "original-system-context", title: "Original", createdAt: new Date(1).toISOString(),
      lastModified: new Date(1).toISOString(), model: faux.getModel(), thinkingLevel: "off", messages: [original],
    }, context);
    assert.deepEqual((await sessions.loadSession(String(conversation.id)))?.messages, [original]);
  } finally { sessions.dispose(); await harness.close(context); }
});
