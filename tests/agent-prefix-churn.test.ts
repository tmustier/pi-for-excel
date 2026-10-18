import assert from "node:assert/strict";
import { test } from "node:test";

import { type AgentTool } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { MemoryStorage, ProviderDoc } from "@earendil-works/pi-durable";
import { openDurableRuntime, routeModels } from "../src/durable/runtime.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";
import { Type } from "typebox";

import {
  createOfficeStreamFn,
  getLastContext,
  getPayloadSnapshots,
} from "../src/auth/stream-proxy.ts";
import { setDebugEnabled } from "../src/debug/debug.ts";

void test("Agent transcript diagnostics report stable prefixes and model, prompt, and tool changes", async () => {
  const priorDocument = Reflect.get(globalThis, "document") as unknown;
  const priorWindow = Reflect.get(globalThis, "window") as unknown;
  Reflect.set(globalThis, "document", new EventTarget());
  const values = new Map<string, string>();
  Reflect.set(globalThis, "window", { localStorage: {
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => values.set(key, value),
  } });

  try {
    setDebugEnabled(true);
    const faux = fauxProvider({ models: [
      { id: "prefix-a", contextWindow: 200_000, maxTokens: 4_096 },
      { id: "prefix-b", contextWindow: 200_000, maxTokens: 4_096 },
    ] });
    faux.setResponses([
      fauxAssistantMessage("first"),
      fauxAssistantMessage("second"),
      fauxAssistantMessage("third"),
    ]);
    const models = createModels();
    models.setProvider(faux.provider);

    const parameters = Type.Object({ query: Type.String() });
    const searchTool: AgentTool<typeof parameters, undefined> = {
      name: "search_docs",
      label: "Search docs",
      description: "Search docs",
      parameters,
      execute: () => Promise.resolve({
        content: [{ type: "text", text: "unused" }],
        details: undefined,
      }),
    };
    const sessionId = "agent-transcript-prefix-regression";
    const runtime = await openDurableRuntime({ storage: new MemoryStorage(),
      models: routeModels(models, createOfficeStreamFn(() => Promise.resolve(undefined), models)),
      workbookId: "prefix", coordinator: createWorkbookCoordinator(), autoCompaction: false,
      mutationHost: { prepare: () => { throw new Error("Unexpected mutation"); }, observe: () => { throw new Error("Unexpected mutation"); } },
    });
    try {
      const conversation = await new DurableSessions(runtime.harness, models, faux.getModel()).create(faux.getModel("prefix-a"), "off");
      await runtime.harness.commit(async (tx) => { (await tx.doc(ProviderDoc, conversation.id)).sessionId = sessionId; }, context);
      await runtime.install(conversation, { tools: [searchTool], systemPrompt: "base" }, (history) => Promise.resolve(history));
      const prompt = async (content: string): Promise<void> => {
        const submission = await conversation.submit({ type: "input", content }, context);
        assert.equal((await submission.wait(context)).status, "done");
      };
      await prompt("first turn");
      await prompt("stable turn");
      await conversation.configure({ model: { provider: faux.getModel().provider, modelId: "prefix-b" } }, context);
      await runtime.install(conversation, { tools: [], systemPrompt: "base\n\nchanged" }, (history) => Promise.resolve(history));
      await prompt("changed turn");

      const snapshots = getPayloadSnapshots().filter(
        (snapshot) => snapshot.sessionId === sessionId,
      );
      assert.deepEqual(
        snapshots.map((snapshot) => snapshot.prefixChangeReasons),
        [[], [], ["model", "systemPrompt", "tools"]],
      );
      assert.equal(snapshots[0]?.systemChars, "base".length);
      assert.equal(snapshots[0]?.toolCount, 1);
      assert.equal(snapshots[0]?.toolBundle, "full");
      assert.equal(snapshots[0]?.toolsIncluded, true);
      assert.equal(snapshots[2]?.systemChars, "base\n\nchanged".length);
      assert.equal(snapshots[2]?.toolCount, 0);
      assert.equal(snapshots[2]?.toolBundle, "none");
      assert.equal(snapshots[2]?.toolsIncluded, false);

      const captured = getLastContext(sessionId);
      assert.equal(captured?.systemPrompt, "base\n\nchanged");
      assert.deepEqual(captured?.tools, []);
    } finally { await runtime.harness.close(context); }
  } finally {
    setDebugEnabled(false);
    if (priorDocument === undefined) Reflect.deleteProperty(globalThis, "document"); else Reflect.set(globalThis, "document", priorDocument);
    if (priorWindow === undefined) Reflect.deleteProperty(globalThis, "window"); else Reflect.set(globalThis, "window", priorWindow);
  }
});
