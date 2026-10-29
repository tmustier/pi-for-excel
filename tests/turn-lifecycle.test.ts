import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  type Context,
  type FauxResponseStep,
} from "@earendil-works/pi-ai";

import { createActionQueue } from "../src/taskpane/action-queue.ts";
import { createQueueDisplay } from "../src/taskpane/queue-display.ts";
import { createTabState, type TabActivity } from "../src/taskpane/tab-state.ts";
import { describeTurnFailure } from "../src/taskpane/turn-error.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import { streamWithProviderFailure } from "../src/auth/provider-failure.ts";

function requestUserTexts(context: Context): string[] {
  return context.messages.flatMap((message) => {
    if (message.role !== "user") return [];
    return [typeof message.content === "string"
      ? message.content
      : message.content.map((part) => (part.type === "text" ? part.text : "")).join("")];
  });
}

const untilStopped: FauxResponseStep = (_context, options) => new Promise((resolve) => {
  const finish = () => resolve(fauxAssistantMessage("partial", { stopReason: "aborted" }));
  if (options?.signal?.aborted) finish();
  options?.signal?.addEventListener("abort", finish, { once: true });
});

function createTab(responses: FauxResponseStep[], opts?: {
  runCompact?: (args: string, signal: AbortSignal) => Promise<void>;
  transformContext?: (messages: AgentMessage[]) => Promise<AgentMessage[]>;
  fetch?: typeof globalThis.fetch;
}) {
  const faux = fauxProvider({ models: [{ id: "turn-test", contextWindow: 200_000, maxTokens: 4_096 }] });
  faux.setResponses(responses);
  const models = createModels();
  models.setProvider(faux.provider);
  const requests: Context[] = [];
  const agent = new Agent({
    initialState: { model: faux.getModel(), messages: [], tools: [] },
    convertToLlm: createConvertToLlm(),
    transformContext: opts?.transformContext,
    streamFn: (model, context, options) => {
      requests.push(structuredClone(context));
      return streamWithProviderFailure((requestOptions) => models.streamSimple(model, context, requestOptions), { ...options, ...(opts?.fetch ? { fetch: opts.fetch } : {}) });
    },
  });
  const tab = createTabState();
  const activities: Array<TabActivity | null> = [];
  tab.subscribe(() => {
    if (activities[activities.length - 1] !== tab.activity) activities.push(tab.activity);
  });
  const queueDisplay = createQueueDisplay({ agent, held: tab });
  const settled: boolean[] = [];
  const queue = createActionQueue({
    agent,
    tab,
    queueDisplay,
    autoCompactEnabled: false,
    runCompact: opts?.runCompact ?? (() => Promise.resolve()),
    onTurnSettled: ({ overflowAfterRecovery }) => settled.push(overflowAfterRecovery),
  });

  const steer = (text: string, followUp = false) => {
    const message: AgentMessage = { role: "user", content: [{ type: "text", text }], timestamp: Date.now() };
    if (followUp) agent.followUp(message); else agent.steer(message);
    queueDisplay.add(followUp ? "follow-up" : "steer", text, message);
  };

  return { agent, tab, queue, requests, activities, settled, steer };
}

async function waitFor(predicate: () => boolean, timeoutMs = 10_000): Promise<void> {
  const started = Date.now();
  while (!predicate()) {
    if (Date.now() - started > timeoutMs) throw new Error("Timed out waiting for condition");
    await new Promise<void>((resolve) => setTimeout(resolve, 5));
  }
}

void test("Stop keeps each waiting steer, follow-up and queued prompt as its own unsent row, and the next turn does not get them", async () => {
  const tab = createTab([untilStopped, fauxAssistantMessage("OK")]);
  tab.tab.setDraft("half-typed");

  tab.queue.enqueuePrompt("Count to 40");
  await waitFor(() => tab.agent.state.isStreaming);
  tab.steer("Then say BANANA.");
  tab.steer("Then say MANGO.", true);
  tab.queue.enqueueCommand("compact", "");

  const taken = tab.queue.stop();
  await waitFor(() => !tab.queue.isBusy());

  assert.deepEqual(taken, { texts: ["Then say BANANA.", "Then say MANGO."], cancelledCommands: ["/compact"] });
  assert.deepEqual(tab.tab.held, ["Then say BANANA.", "Then say MANGO."]);
  assert.equal(tab.tab.draft, "half-typed");
  assert.equal(tab.agent.hasQueuedMessages(), false);

  assert.equal(tab.queue.editLast(), true);
  assert.equal(tab.tab.draft, "Then say MANGO.\n\nhalf-typed");
  assert.deepEqual(tab.tab.held, ["Then say BANANA."]);

  tab.queue.enqueuePrompt("Reply with just OK.");
  await waitFor(() => tab.requests.length === 2 && !tab.queue.isBusy());

  assert.deepEqual(requestUserTexts(tab.requests[1] ?? { messages: [] }), ["Count to 40", "Reply with just OK."]);
  assert.equal(tab.agent.state.messages.at(-1)?.role, "assistant");
  tab.queue.shutdown();
});

void test("a failed turn keeps its undelivered steers as unsent rows instead of sending them in the next turn", async () => {
  const tab = createTab([
    async (_context, options) => {
      await new Promise<void>((resolve) => setTimeout(resolve, 30));
      return options?.signal?.aborted
        ? fauxAssistantMessage("", { stopReason: "aborted" })
        : fauxAssistantMessage("", { stopReason: "error", errorMessage: "400 invalid tool schema" });
    },
    fauxAssistantMessage("OK"),
  ]);

  tab.queue.enqueuePrompt("first");
  await waitFor(() => tab.agent.state.isStreaming);
  tab.steer("stale steer");
  await waitFor(() => !tab.queue.isBusy());

  assert.deepEqual(tab.tab.held, ["stale steer"]);
  assert.equal(tab.tab.draft, "");
  tab.queue.enqueuePrompt("second");
  await waitFor(() => tab.requests.length === 2 && !tab.queue.isBusy());
  assert.equal(requestUserTexts(tab.requests[1] ?? { messages: [] }).includes("stale steer"), false);
  tab.queue.shutdown();
});

void test("a transient provider failure is retried with a visible countdown and the turn completes", async () => {
  const selection: AgentMessage = { role: "user", content: [{ type: "text", text: "SELECTION A1" }], timestamp: 0 };
  const tab = createTab([
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "503 Service Unavailable" }),
    fauxAssistantMessage("recovered"),
  ], { transformContext: (messages) => Promise.resolve([...messages, selection]) });

  tab.queue.enqueuePrompt("hello");
  await waitFor(() => !tab.queue.isBusy());

  assert.equal(tab.requests.length, 2);
  assert.match(tab.activities.find((activity) => activity !== null)?.label ?? "", /^Retrying in 2s \(attempt 1 of 6\)/u);
  assert.equal(tab.tab.activity, null);
  const assistants = tab.agent.state.messages.filter((message) => message.role === "assistant");
  assert.deepEqual(assistants.map((message) => message.stopReason), ["stop"]);
  assert.deepEqual(tab.requests.map((request) => requestUserTexts(request)), [["hello", "SELECTION A1"], ["hello", "SELECTION A1"]]);
  const storedUserTexts = tab.agent.state.messages.flatMap((message) =>
    message.role === "user" && typeof message.content !== "string"
      ? [message.content.map((part) => (part.type === "text" ? part.text : "")).join("")]
      : []);
  assert.deepEqual(storedUserTexts, ["hello"]);
  tab.queue.shutdown();
});

void test("Stop during a retry wait keeps the failure and sends no further request", async () => {
  const tab = createTab([
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "429 rate limit exceeded" }),
    fauxAssistantMessage("should not be requested"),
  ]);

  tab.queue.enqueuePrompt("hello");
  await waitFor(() => tab.tab.activity !== null);
  tab.queue.stop();
  await waitFor(() => !tab.queue.isBusy());
  await new Promise<void>((resolve) => setTimeout(resolve, 2_500));

  assert.equal(tab.requests.length, 1);
  const last = tab.agent.state.messages.at(-1);
  assert.equal(last?.role === "assistant" ? last.errorMessage : null, "429 rate limit exceeded");
  tab.queue.shutdown();
});

void test("a partial answer that failed mid-stream survives a retry that fails before streaming", async () => {
  const tab = createTab([
    fauxAssistantMessage("partial answer", { stopReason: "error", errorMessage: "network error" }),
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "Failed to fetch: network error" }),
    fauxAssistantMessage("should not be requested"),
  ]);

  tab.queue.enqueuePrompt("hello");
  await waitFor(() => tab.requests.length === 2 && tab.tab.activity !== null);
  tab.queue.stop();
  await waitFor(() => !tab.queue.isBusy());

  assert.equal(tab.requests.length, 2);
  const partial = tab.agent.state.messages.find((message) => message.role === "interruptedResponse");
  assert.equal(partial?.role === "interruptedResponse" ? partial.message.content[0]?.type === "text" && partial.message.content[0].text : null, "partial answer");
  assert.ok(tab.requests[1]?.messages.some((message) => message.role === "assistant" && message.content.some((part) => part.type === "text" && part.text === "partial answer")), "retry receives the partial answer instead of restarting it");
  assert.equal(tab.agent.state.messages.filter((message) => message.role === "user").length, 1);
  tab.queue.shutdown();
});

void test("a quota or usage-limit failure is not retried", async () => {
  for (const errorMessage of ["insufficient_quota: billing", "You have hit your ChatGPT usage limit (plus plan). Try again in ~42 min."]) {
    const tab = createTab([fauxAssistantMessage("", { stopReason: "error", errorMessage })]);
    tab.queue.enqueuePrompt("hello");
    await waitFor(() => !tab.queue.isBusy());
    assert.equal(tab.requests.length, 1, errorMessage);
    assert.deepEqual(tab.activities.filter((activity) => activity !== null), [], errorMessage);
    assert.deepEqual(tab.settled, [false]);
    tab.queue.shutdown();
  }
});

void test("manual Retry resumes the failed point without another user message", async () => {
  const tab = createTab([
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "401 invalid API key" }),
    fauxAssistantMessage("recovered"),
  ]);
  tab.queue.enqueuePrompt("original request");
  await waitFor(() => !tab.queue.isBusy());
  assert.equal(tab.queue.retry(), true);
  assert.equal(tab.queue.retry(), false, "a double click cannot duplicate a retry");
  await waitFor(() => !tab.queue.isBusy());
  assert.deepEqual(tab.requests.map(requestUserTexts), [["original request"], ["original request"]]);
  assert.deepEqual(tab.agent.state.messages.filter((message) => message.role === "assistant").map((message) => message.stopReason), ["stop"], JSON.stringify(tab.agent.state.messages));
  tab.queue.shutdown();
});

for (const row of [
  { code: "rate_limit_exceeded", retryAfter: "7", countdown: 7 },
  { code: "rate_limit_exceeded", retryAfter: "120", countdown: 60 },
  { code: "usage_limit_reached", retryAfter: "1", countdown: null },
  { code: "insufficient_quota", retryAfter: "1", countdown: null },
] as const) {
  void test(`HTTP 429 ${row.code} with Retry-After ${row.retryAfter} uses provider facts, not rewritten wording`, async () => {
    const tab = createTab([async (_context, options) => {
      await options?.fetch?.("https://provider.test/response");
      return fauxAssistantMessage("", { stopReason: "error", errorMessage: "You have hit your ChatGPT usage limit." });
    }], {
      fetch: () => Promise.resolve(new Response(JSON.stringify({ error: { code: row.code } }), { status: 429, headers: { "Retry-After": row.retryAfter } })),
    });
    tab.queue.enqueuePrompt("hello");
    await waitFor(() => tab.tab.activity !== null || !tab.queue.isBusy());
    if (row.countdown !== null) {
      assert.match(tab.tab.activity?.label ?? "", new RegExp(`^Retrying in ${row.countdown}s`));
      tab.queue.stop();
    } else {
      assert.equal(tab.tab.activity, null);
    }
    await waitFor(() => !tab.queue.isBusy());
    assert.equal(tab.requests.length, 1);
    tab.queue.shutdown();
  });
}

void test("Stop cancels a running /compact for this tab and returns the prompt held behind it", async () => {
  let compactSignal: AbortSignal | null = null;
  const tab = createTab([fauxAssistantMessage("never")], {
    runCompact: (_args, signal) => {
      compactSignal = signal;
      return new Promise<void>((resolve) => signal.addEventListener("abort", () => resolve(), { once: true }));
    },
  });

  tab.queue.enqueueCommand("compact", "");
  tab.queue.enqueuePrompt("held behind compaction");
  await waitFor(() => compactSignal !== null);
  const taken = tab.queue.stop();
  await waitFor(() => !tab.queue.isBusy());

  assert.equal(compactSignal?.aborted, true);
  assert.deepEqual(taken.texts, ["held behind compaction"]);
  assert.deepEqual(tab.tab.held, ["held behind compaction"]);
  assert.equal(tab.requests.length, 0);
  tab.queue.shutdown();
});

void test("editing the last waiting steer takes only that one back; the earlier one is still delivered", async () => {
  const slow: FauxResponseStep = () => new Promise((resolve) => setTimeout(() => resolve(fauxAssistantMessage("one")), 50));
  const tab = createTab([slow, fauxAssistantMessage("two")]);

  tab.queue.enqueuePrompt("start");
  await waitFor(() => tab.agent.state.isStreaming);
  tab.steer("first steer");
  tab.steer("second steer");
  assert.equal(tab.queue.editLast(), true);
  await waitFor(() => !tab.queue.isBusy());

  assert.equal(tab.tab.draft, "second steer");
  assert.deepEqual(tab.tab.held, []);
  assert.deepEqual(requestUserTexts(tab.requests[1] ?? { messages: [] }), ["start", "first steer"]);
  tab.queue.shutdown();
});

void test("inline error summaries explain recovery without repeating provider details", () => {
  const failed = (errorMessage: string): AgentMessage[] => [
    fauxAssistantMessage("", { stopReason: "error", errorMessage }),
  ];
  const base = {
    isContextOverflow: false,
    overflowAfterRecovery: false,
    missingCredentialsFor: null,
    online: true,
    proxyUrl: null,
    defaultProxyUrl: "https://localhost:3003",
    proxyBlockedByBrowser: false,
  };

  assert.equal(describeTurnFailure({ ...base, messages: [fauxAssistantMessage("partial", { stopReason: "aborted" })] }), null);
  assert.equal(describeTurnFailure({ ...base, messages: [fauxAssistantMessage("fine")] }), null);
  assert.equal(
    describeTurnFailure({ ...base, messages: failed("Request cancelled: invalid tool schema") }),
    "Pi couldn’t finish this response. Try again or choose another model.",
  );
  assert.match(
    describeTurnFailure({ ...base, messages: failed("400 model_not_supported").map((message) => ({ ...message, provider: "github-copilot" })) }) ?? "",
    /Copilot plan.*GPT-5\.4 nano/u,
  );
  assert.equal(
    describeTurnFailure({ ...base, messages: failed("400 model_not_supported") }),
    "Pi couldn’t finish this response. Try again or choose another model.",
  );
  assert.equal(
    describeTurnFailure({ ...base, messages: failed("Provider is not configured: anthropic"), missingCredentialsFor: "anthropic" }),
    "API key required for anthropic.",
  );
  assert.equal(
    describeTurnFailure({ ...base, messages: failed("The anthropic login has expired and could not be refreshed."), expiredLoginFor: "Anthropic" }),
    "Your Anthropic login has expired and could not be renewed. Sign in again in Settings.",
  );
  assert.match(describeTurnFailure({ ...base, messages: failed("Failed to fetch"), online: false }) ?? "", /^You're offline/u);
  assert.match(
    describeTurnFailure({ ...base, messages: failed("Failed to fetch"), proxyUrl: "https://localhost:3003" }) ?? "",
    /through the proxy at https:\/\/localhost:3003/u,
  );
  assert.match(describeTurnFailure({ ...base, messages: failed("Failed to fetch") }) ?? "", /likely CORS/u);
  for (const proxyUrl of [null, "https://localhost:3003"]) {
    assert.match(
      describeTurnFailure({ ...base, messages: failed("Failed to fetch"), proxyUrl, proxyBlockedByBrowser: true }) ?? "",
      /Excel desktop/u,
    );
  }
  assert.match(
    describeTurnFailure({ ...base, messages: failed("context_length_exceeded"), isContextOverflow: true, overflowAfterRecovery: true }) ?? "",
    /^Context window still exceeded after compacting and retrying/u,
  );
  assert.doesNotMatch(
    describeTurnFailure({ ...base, messages: failed("context_length_exceeded"), isContextOverflow: true }) ?? "",
    /will try/u,
  );
});
