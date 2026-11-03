import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider, type FauxResponseStep } from "@earendil-works/pi-ai";

import { trackTurns } from "../src/taskpane/turn-tracker.ts";

function createAgent(responses: FauxResponseStep[]): Agent {
  const faux = fauxProvider({ models: [{ id: "turn-tracker", contextWindow: 200_000, maxTokens: 4_096 }] });
  faux.setResponses(responses);
  const models = createModels();
  models.setProvider(faux.provider);
  return new Agent({
    initialState: { model: faux.getModel(), messages: [], tools: [] },
    streamFn: (model, context, options) => models.streamSimple(model, context, options),
  });
}

void test("a credential problem recorded during one turn is not reported for the next turn", async () => {
  let turns: ReturnType<typeof trackTurns> | null = null;
  const agent = createAgent([
    () => {
      const running = turns?.running();
      if (running) running.missingCredentialsFor = "mistral";
      return fauxAssistantMessage("", { stopReason: "aborted" });
    },
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "HTTP 500" }),
  ]);
  turns = trackTurns(agent);

  await agent.prompt("first");
  assert.equal(turns.current().missingCredentialsFor, "mistral");

  await agent.prompt("second");
  assert.equal(turns.current().missingCredentialsFor, null);
  turns.dispose();
});

void test("a key lookup outside a turn has no turn to record into", async () => {
  const agent = createAgent([fauxAssistantMessage("Done")]);
  const turns = trackTurns(agent);
  assert.equal(turns.running(), null);

  await agent.prompt("hello");
  assert.equal(turns.running(), null);
  assert.equal(turns.current().expiredLoginFor, null);
  turns.dispose();
});
