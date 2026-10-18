import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxProvider, fauxAssistantMessage } from "@earendil-works/pi-ai/providers/faux";
import { createRegistry, defineExtension, Harness, MemoryStorage } from "@earendil-works/pi-durable";
import { ConversationPresenter } from "../src/durable/conversation-view.ts";
import { createExtensionAPI, loadExtension } from "../src/commands/extension-api.ts";

void test("installed legacy extension reads committed streaming events and follows tabs alongside native API", async () => {
  const models = createModels();
  const faux = fauxProvider({ tokensPerSecond: 15, tokenSize: { min: 1, max: 1 } });
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const registry = createRegistry();
  const extension = defineExtension({ name: "tool-metadata", tools: [{ name: "legacy_visible_tool", description: "Read tool metadata", parameters: Type.Object({}),
    execute: () => Promise.resolve({ output: "unused" }) }] });
  registry.install(extension);
  const harness = await Harness.open(new MemoryStorage(), { models, registry }, context);
  const first = await harness.root(context, { agent: { model: { provider: model.provider, modelId: model.id } } });
  await first.configure({ extensions: [extension] }, context);
  const second = await harness.createConversation({ ownership: { kind: "ownerless" } }, context);
  await second.configure({ model: { provider: model.provider, modelId: model.id }, extensions: [] }, context);
  const one = await ConversationPresenter.open(harness, first, models, model);
  const two = await ConversationPresenter.open(harness, second, models, model);
  let active = one;
  const changes = new Set<() => void>();
  const api = createExtensionAPI({
    getActiveConversation: () => active,
    subscribeActiveConversation: (listener) => { changes.add(listener); return () => { changes.delete(listener); }; },
  });
  const ended: string[] = [];
  const deltas: string[] = [];
  const completed: string[] = [];
  let nativeEnds = 0;
  const handle = await loadExtension(api, (installed) => {
    const raw = installed.agent.raw;
    const stop = installed.onAgentEvent((event) => {
      assert.equal(raw.state.model.id, model.id);
      if (event.type === "message_end" && event.message.role === "assistant") ended.push(JSON.stringify(event.message.content));
      if (event.type === "message_update") {
        assert.equal(event.message.role, "assistant");
        const update = event.assistantMessageEvent;
        if (update.type === "text_delta") deltas.push(update.delta);
        if (update.type === "text_end") completed.push(update.content);
      }
    });
    const stopNative = installed.onConversationEvent((event) => {
      if (event.type === "message_end") {
        assert.ok(event.entry.id);
        if (event.entry.model?.some((message) => message.role === "user" || message.role === "assistant")) nativeEnds++;
      }
    });
    return () => { stop(); stopNative(); };
  });
  try {
    faux.setResponses([fauxAssistantMessage("First streamed reply")]);
    await first.submit({ type: "input", content: "first" }, context);
    await first.waitForIdle(context);
    await new Promise((resolve) => setTimeout(resolve, 150));
    assert.equal(ended.length, 1);
    assert.ok(deltas.length > 0, "Durable streaming deltas reach legacy assistantMessageEvent");
    const retained = api.agent.raw;
    assert.equal(retained.state.tools[0]?.name, "legacy_visible_tool");
    assert.equal("execute" in (retained.state.tools[0] ?? {}), false, "read-state tools cannot bypass native dispatch");
    retained.state.messages.length = 0;
    assert.ok(one.current.messages.length > 0, "read adapter cannot mutate presenter state");
    active = two;
    for (const listener of changes) listener();
    assert.equal(api.agent.conversation, second);
    assert.deepEqual(retained.state.tools, []);
    assert.deepEqual(retained.state.messages, two.current.messages);
    faux.setResponses([fauxAssistantMessage("Old tab ignored")]);
    await first.submit({ type: "input", content: "background" }, context);
    await first.waitForIdle(context);
    faux.setResponses([fauxAssistantMessage("Second tab reply")]);
    await second.submit({ type: "input", content: "second" }, context);
    await second.waitForIdle(context);
    await new Promise((resolve) => setTimeout(resolve, 150));
    assert.equal(ended.length, 2);
    assert.match(ended[1] ?? "", /Second tab reply/);
    assert.deepEqual(completed, ["First streamed reply", "Second tab reply"]);
    assert.equal(nativeEnds, 4);
    await handle.deactivate();
    assert.equal(changes.size, 0);
  } finally {
    await handle.deactivate();
    await one.dispose(); await two.dispose();
    await harness.close(context);
  }
});
