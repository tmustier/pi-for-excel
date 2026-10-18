import assert from "node:assert/strict";
import { test } from "node:test";

import { type AgentMessage, type AgentTool } from "@earendil-works/pi-agent-core";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { MemoryStorage } from "@earendil-works/pi-durable";
import { Type } from "typebox";
import type { Context } from "@earendil-works/pi-ai";
import { createModels } from "@earendil-works/pi-ai/models";
import { fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai/providers/faux";

import { effectiveToolOutputLimits } from "../src/context/window-budgets.ts";
import { openDurableRuntime } from "../src/durable/runtime.ts";
import { DurableSessions } from "../src/durable/sessions.ts";
import { createWorkbookCoordinator } from "../src/workbook/coordinator.ts";
import { applyToolOutputTruncation } from "../src/tools/output-truncation.ts";

function toolResult(id: number, text: string): Extract<AgentMessage, { role: "toolResult" }> {
  return { role: "toolResult", toolCallId: String(id), toolName: "read_range", content: [{ type: "text", text }], isError: false, timestamp: id };
}

void test("agent requests apply full, scaled, and floor budgets for each model window class", async () => {
  const cases = [
    { name: "128k", window: 128_000, kept: 6, maxBytes: 51_200, maxLines: 2_000 },
    { name: "larger", window: 200_000, kept: 6, maxBytes: 51_200, maxLines: 2_000 },
    { name: "unknown", window: undefined, kept: 6, maxBytes: 51_200, maxLines: 2_000 },
    { name: "invalid", window: 0, kept: 6, maxBytes: 51_200, maxLines: 2_000 },
    { name: "65k", window: 65_536, kept: 3, maxBytes: 26_214, maxLines: 1_024 },
    { name: "tiny", window: 4_096, kept: 2, maxBytes: 8_192, maxLines: 200 },
  ] as const;

  for (const scenario of cases) {
    const limits = effectiveToolOutputLimits(scenario.window);
    const sourceTool: AgentTool = {
      name: "read_range",
      label: "Read range",
      description: "Returns rows",
      parameters: Type.Object({}),
      execute: () => Promise.resolve({ content: [{ type: "text", text: "x\n".repeat(30_000) }], details: {} }),
    };
    const wrapped = applyToolOutputTruncation([sourceTool], { limits: () => limits });
    const result = await wrapped[0]?.execute("large", {});
    if (!result) throw new Error("Expected wrapped tool result.");
    const truncation = Reflect.get(result.details, "outputTruncation") as unknown;
    if (typeof truncation !== "object" || truncation === null) throw new Error("Expected truncation metadata.");
    assert.equal(Reflect.get(truncation, "maxBytes"), scenario.maxBytes, scenario.name);
    assert.equal(Reflect.get(truncation, "maxLines"), scenario.maxLines, scenario.name);

    const faux = fauxProvider({ models: [{ id: `budget-${scenario.name}`, contextWindow: 128_000, maxTokens: 4_096 }] });
    faux.setResponses([fauxAssistantMessage("done")]);
    const model = faux.getModel();
    Reflect.set(model, "contextWindow", scenario.window);
    const models = createModels();
    const requests: Context[] = [];
    const messages = Array.from({ length: 7 }, (_, index) => toolResult(index + 1, "z".repeat(1_500)));
    models.setProvider({ ...faux.provider, streamSimple: (requestModel, transcript, options) => {
      requests.push(structuredClone(transcript));
      return faux.provider.streamSimple(requestModel, transcript, options);
    },
    });
    const runtime = await openDurableRuntime({ storage: new MemoryStorage(), models,
      workbookId: "budgets", coordinator: createWorkbookCoordinator(), autoCompaction: false,
      mutationHost: { prepare: () => { throw new Error("Unexpected mutation"); }, observe: () => { throw new Error("Unexpected mutation"); } },
    });
    try {
      const conversation = await new DurableSessions(runtime.harness, models, model).create(model, "off");
      await runtime.install(conversation, { tools: [], systemPrompt: "" }, (history) => Promise.resolve(history));
      await runtime.harness.commit(async (tx) => {
        for (const message of messages) await tx.appendEntry(conversation.id, { kind: "test.history", model: [
          { ...fauxAssistantMessage(""), stopReason: "toolUse", content: [{ type: "toolCall", id: message.toolCallId, name: message.toolName, arguments: {} }] }, message,
        ] });
      }, context);
      const submission = await conversation.submit({ type: "input", content: "continue" }, context);
      assert.equal((await submission.wait(context)).status, "done");
      const request = requests[0];
      if (!request) throw new Error("Expected outbound model request.");
      const verbatim = request.messages.filter((message) => message.role === "toolResult" && JSON.stringify(message.content).includes("z".repeat(1_400)));
      assert.equal(verbatim.length, scenario.kept, scenario.name);
    } finally { await runtime.harness.close(context); }
  }
});
