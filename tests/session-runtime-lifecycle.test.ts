import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent } from "@earendil-works/pi-agent-core";
import { createModels, fauxAssistantMessage, fauxProvider } from "@earendil-works/pi-ai";
import { getBuiltinModel } from "@earendil-works/pi-ai/providers/all";

import { settingsBackedSessionStorage } from "../src/host/session-storage.ts";
import { initAppStorage } from "../src/storage/init-app-storage.ts";
import type { SessionData, StorageBackend, StorageTransaction } from "../src/storage/local/types.ts";
import type { ActionQueue } from "../src/taskpane/action-queue.ts";
import type { QueueDisplay } from "../src/taskpane/queue-display.ts";
import { createTabState } from "../src/taskpane/tab-state.ts";
import { SessionRuntimeManager, type SessionRuntime } from "../src/taskpane/session-runtime-manager.ts";
import { createArchivedMessagesMessage } from "../src/messages/archived-history.ts";
import { createCompactionSummaryMessage } from "../src/messages/compaction.ts";
import { setupSessionPersistence } from "../src/taskpane/sessions.ts";

class MemoryStorageBackend implements StorageBackend {
  private readonly stores = new Map<string, Map<string, unknown>>();

  private store(name: string): Map<string, unknown> {
    const existing = this.stores.get(name);
    if (existing) return existing;
    const created = new Map<string, unknown>();
    this.stores.set(name, created);
    return created;
  }

  get<T = unknown>(storeName: string, key: string): Promise<T | null> {
    const value = this.store(storeName).get(key);
    return Promise.resolve(value === undefined ? null : structuredClone(value) as T);
  }

  set<T = unknown>(storeName: string, key: string, value: T): Promise<void> {
    this.store(storeName).set(key, structuredClone(value));
    return Promise.resolve();
  }

  delete(storeName: string, key: string): Promise<void> {
    this.store(storeName).delete(key);
    return Promise.resolve();
  }

  keys(storeName: string, prefix?: string): Promise<string[]> {
    const keys = [...this.store(storeName).keys()];
    return Promise.resolve(prefix ? keys.filter((key) => key.startsWith(prefix)) : keys);
  }

  getAllFromIndex<T = unknown>(storeName: string): Promise<T[]> {
    return Promise.resolve([...this.store(storeName).values()].map((value) => structuredClone(value)) as T[]);
  }

  clear(storeName: string): Promise<void> {
    this.store(storeName).clear();
    return Promise.resolve();
  }

  has(storeName: string, key: string): Promise<boolean> {
    return Promise.resolve(this.store(storeName).has(key));
  }

  transaction<T>(
    _storeNames: string[],
    _mode: "readonly" | "readwrite",
    operation: (tx: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    return operation({
      get: <V = unknown>(storeName: string, key: string) => this.get<V>(storeName, key),
      set: <V = unknown>(storeName: string, key: string, value: V) => this.set(storeName, key, value),
      delete: (storeName: string, key: string) => this.delete(storeName, key),
    });
  }

  getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return Promise.resolve({ usage: 0, quota: 0, percent: 0 });
  }

  requestPersistence(): Promise<boolean> {
    return Promise.resolve(true);
  }
}

const actionQueue: ActionQueue = {
  enqueuePrompt: () => {},
  enqueueCommand: () => {},
  editLast: () => false,
  stop: () => ({ texts: [], cancelledCommands: [] }),
  isBusy: () => false,
  shutdown: () => {},
};

function createQueueDisplay(): QueueDisplay {
  return {
    add: () => {},
    clear: () => {},
    drainQueuedMessages: () => [],
    removeQueuedMessage: () => {},
    setRowHandlers: () => {},
    lastEditableRow: () => null,
    setActionQueue: () => {},
    attach: () => {},
    detach: () => {},
  };
}

function createEnvironment() {
  const backend = new MemoryStorageBackend();
  const { sessions, settings } = initAppStorage("session-runtime-lifecycle", backend);
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  let nextRuntime = 1;

  const createRuntime = async (autoRestoreLatest: boolean): Promise<SessionRuntime> => {
    const agent = new Agent({
      initialState: {
        model: getBuiltinModel("openai", "gpt-5.6-sol"),
        messages: [],
        tools: [],
      },
      streamFn: (model, context, options) => models.streamSimple(model, context, options),
      getApiKey: () => "lifecycle-test-key",
    });
    const persistence = await setupSessionPersistence({
      agent,
      sessions,
      settings,
      resolveModel: (persisted) => models.getModel(persisted.provider, persisted.id),
      autoRestoreLatest,
      spreadsheetHost: {
        getWorkbookContext: () => Promise.resolve({
          workbookId: "url_sha256:runtime-lifecycle-workbook",
          workbookName: null,
          source: "document.url",
        }),
        sessionStorage: settingsBackedSessionStorage,
      },
    });
    const runtimeId = `runtime-${nextRuntime}`;
    nextRuntime += 1;
    return {
      runtimeId,
      agent,
      persistence,
      actionQueue,
      queueDisplay: createQueueDisplay(),
      tab: createTabState(),
      lockState: "idle",
      refreshCapabilities: () => Promise.resolve(),
      dispose: () => persistence.dispose(),
    };
  };

  const manager = new SessionRuntimeManager({
    createRuntime: (options) => createRuntime(options.autoRestoreLatest),
  });
  return { manager, createRuntime, sessions };
}

void test("runtime tab snapshots follow create, switch, and close lifecycle", async () => {
  const environment = createEnvironment();
  const first = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const second = await environment.manager.createRuntime({ activate: false, autoRestoreLatest: false });

  assert.deepEqual(environment.manager.snapshot(), {
    activeRuntimeId: first.runtimeId,
    tabs: [
      {
        runtimeId: first.runtimeId,
        sessionId: first.persistence.getSessionId(),
        tabNumber: 1,
        title: "Chat 1",
        isActive: true,
        isStreaming: false,
        isBusy: false,
        lockState: "idle",
        attention: "none",
      },
      {
        runtimeId: second.runtimeId,
        sessionId: second.persistence.getSessionId(),
        tabNumber: 2,
        title: "Chat 2",
        isActive: false,
        isStreaming: false,
        isBusy: false,
        lockState: "idle",
        attention: "none",
      },
    ],
  });

  environment.manager.switchRuntime(second.runtimeId);
  environment.manager.closeRuntime(second.runtimeId);
  assert.deepEqual(environment.manager.snapshot().tabs.map((tab) => ({
    runtimeId: tab.runtimeId,
    isActive: tab.isActive,
  })), [{ runtimeId: first.runtimeId, isActive: true }]);

  await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  assert.deepEqual(environment.manager.snapshotTabs().map((tab) => tab.title), ["Chat 1", "Chat 3"]);
});

void test("model selection updates an empty session in place and survives restart", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const sessionId = runtime.persistence.getSessionId();
  runtime.agent.state.messages.push({ role: "system", content: "", timestamp: 0 });

  const result = await environment.manager.selectModel({
    runtimeId: runtime.runtimeId,
    nextModel: getBuiltinModel("openai-codex", "gpt-5.6-sol"),
  });
  assert.equal(result.outcome, "updated");
  assert.equal(environment.manager.snapshot().tabs.length, 1);
  runtime.agent.state.messages.push({
    role: "user",
    content: "message after choosing the model",
    timestamp: 1_789_030_800_000,
  });
  await runtime.persistence.saveSession({ force: true });
  runtime.dispose();

  const restarted = await environment.createRuntime(true);
  assert.equal(restarted.persistence.getSessionId(), sessionId);
  assert.equal(restarted.agent.state.model.provider, "openai-codex");
  restarted.dispose();
});

async function saveConversation(runtime: SessionRuntime, text: string): Promise<SessionData> {
  runtime.agent.state.messages.push({ role: "user", content: text, timestamp: 1_789_030_800_000 });
  await runtime.persistence.saveSession({ force: true });
  return {
    id: runtime.persistence.getSessionId(),
    title: text,
    model: getBuiltinModel("openai", "gpt-5.6-sol"),
    thinkingLevel: "off",
    messages: structuredClone(runtime.agent.state.messages),
    createdAt: runtime.persistence.getSessionCreatedAt(),
    lastModified: new Date().toISOString(),
  };
}

void test("a session that is already open switches to its tab instead of opening a second one", async () => {
  const environment = createEnvironment();
  const first = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const sessionData = await saveConversation(first, "budget review");
  const second = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });

  for (const target of [{ kind: "new_tab" as const }, { kind: "resume" as const }]) {
    environment.manager.switchRuntime(second.runtimeId);
    const result = await environment.manager.openSession(sessionData, target);
    assert.equal(result.outcome, "focused");
    const tabs = environment.manager.snapshotTabs();
    assert.equal(tabs.length, 2);
    assert.equal(tabs.find((tab) => tab.isActive)?.runtimeId, first.runtimeId);
    assert.notEqual(second.persistence.getSessionId(), sessionData.id);
  }
});

void test("resuming beside a busy tab opens another tab and leaves its session in place", async () => {
  const environment = createEnvironment();
  const source = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const sessionData = await saveConversation(source, "saved elsewhere");
  const busy = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  environment.manager.closeRuntime(source.runtimeId);
  const busySessionId = busy.persistence.getSessionId();
  busy.actionQueue = { ...busy.actionQueue, isBusy: () => true };

  const result = await environment.manager.openSession(sessionData, { kind: "resume" });
  assert.equal(result.outcome, "opened");
  assert.equal(environment.manager.snapshotTabs().length, 2);
  assert.equal(busy.persistence.getSessionId(), busySessionId);
});

void test("resume reuses only an idle empty tab, never a conversation", async () => {
  const environment = createEnvironment();
  const source = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const saved = await saveConversation(source, "saved conversation");
  const empty = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  environment.manager.closeRuntime(source.runtimeId);
  empty.agent.state.messages.push({ role: "system", content: "hidden", timestamp: 0 });
  assert.equal((await environment.manager.openSession(saved, { kind: "resume" })).outcome, "replaced");
  assert.equal(empty.persistence.getSessionId(), saved.id);
  assert.equal(environment.manager.snapshotTabs().length, 1);

  const another = { ...saved, id: crypto.randomUUID(), title: "another conversation" };
  assert.equal((await environment.manager.openSession(another, { kind: "resume" })).outcome, "opened");
  assert.equal(empty.persistence.getSessionId(), saved.id);
  assert.equal(environment.manager.snapshotTabs().length, 2);
});

void test("first exchange generates one isolated title, persists it, and never overwrites a user rename", { timeout: 10_000 }, async () => {
  for (const renameDuringRequest of [false, true]) {
    const environment = createEnvironment();
    const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
    const faux = fauxProvider();
    const models = createModels();
    models.setProvider(faux.provider);
    let notifyTitleRequested = (): void => {};
    const titleRequested = new Promise<void>((resolve) => { notifyTitleRequested = resolve; });
    let answerTitle = (_title: string): void => {};
    const titleResponse = new Promise<string>((resolve) => { answerTitle = resolve; });
    let notifyTitleSettled = (): void => {};
    const titleSettled = new Promise<void>((resolve) => { notifyTitleSettled = resolve; });
    const sessionId = runtime.persistence.getSessionId();
    const chatPrefix = "Original chat instructions";
    const requests: string[] = [];
    runtime.agent.state.messages.push({ role: "system", content: chatPrefix, timestamp: 0 });
    runtime.agent.streamFunction = (model, context, options) => {
      requests.push(options?.sessionId ?? "");
      if (options?.sessionId === `${sessionId}::title`) {
        assert.doesNotMatch(JSON.stringify(context), /"tools":\[\{/);
        assert.doesNotMatch(JSON.stringify(context), /Original chat instructions/);
        assert.equal(options.maxTokens, 64);
      }
      return models.streamSimple(faux.getModel(), context, options);
    };
    faux.setResponses([
      fauxAssistantMessage("We can forecast Q3 revenue from historical sales."),
      async () => {
        notifyTitleRequested();
        return fauxAssistantMessage(await titleResponse);
      },
      fauxAssistantMessage("Next answer"),
    ]);
    const unsubscribe = runtime.persistence.subscribe(() => {
      if (runtime.persistence.getSessionTitle()) notifyTitleSettled();
    });
    await runtime.agent.prompt("Forecast Q3 revenue");
    await titleRequested;
    if (renameDuringRequest) await runtime.persistence.renameSession("My forecast");
    answerTitle("Quarterly Revenue Forecast");
    await titleSettled;
    await runtime.agent.prompt("Show the assumptions");
    await runtime.persistence.saveSession();
    const expected = renameDuringRequest ? "My forecast" : "Quarterly Revenue Forecast";
    assert.equal(environment.manager.snapshotTabs()[0]?.title, expected);
    assert.equal(requests.filter((id) => id.endsWith("::title")).length, 1);
    assert.equal(runtime.agent.sessionId, sessionId);
    assert.match(JSON.stringify(runtime.agent.state.messages), /Original chat instructions/);
    assert.doesNotMatch(JSON.stringify(runtime.agent.state.messages), /Write a short 3–5 word/);
    const saved = await environment.sessions.loadSession(sessionId);
    assert.equal(saved?.title, expected);
    assert.equal(saved?.titleChosen, renameDuringRequest);
    assert.equal(saved?.titleGenerationAttempted, true);
    unsubscribe();
    runtime.dispose();
  }
});

void test("failed title requests keep Chat N and are not charged again after resume", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  const requests: string[] = [];
  const stream: Agent["streamFunction"] = (_model, context, options) => {
    requests.push(options?.sessionId ?? "");
    return models.streamSimple(faux.getModel(), context, options);
  };
  runtime.agent.streamFunction = stream;
  faux.setResponses([
    fauxAssistantMessage("First answer"),
    fauxAssistantMessage("", { stopReason: "error", errorMessage: "Title unavailable" }),
    fauxAssistantMessage("Second answer"),
  ]);
  await runtime.agent.prompt("A question");
  await new Promise((resolve) => setTimeout(resolve, 20));
  await runtime.persistence.saveSession();
  assert.equal(environment.manager.snapshotTabs()[0]?.title, "Chat 1");
  const saved = await environment.sessions.loadSession(runtime.persistence.getSessionId());
  assert.ok(saved);
  assert.equal(saved.titleGenerationAttempted, true);
  assert.equal(saved.title, "");
  const resumed = await environment.createRuntime(false);
  await resumed.persistence.applyLoadedSession(saved);
  resumed.agent.streamFunction = stream;
  await resumed.agent.prompt("Continue");
  await resumed.persistence.saveSession();
  assert.equal(requests.filter((id) => id.endsWith("::title")).length, 1);
  runtime.dispose();
  resumed.dispose();
});

void test("a chosen tab title and the tab number survive reopening the session", async () => {
  const environment = createEnvironment();
  await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const named = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  await named.persistence.renameSession("Budget review");
  const sessionId = named.persistence.getSessionId();
  const tabNumber = environment.manager.getTabNumber(named.runtimeId) ?? 0;
  environment.manager.closeRuntime(named.runtimeId);

  const loaded = await environment.sessions.loadSession(sessionId);
  assert.ok(loaded);
  const result = await environment.manager.openSession(loaded, { kind: "new_tab", tabNumber });
  assert.equal(result.outcome, "opened");
  const tab = environment.manager.snapshotTabs().find((candidate) => candidate.sessionId === sessionId);
  assert.equal(tab?.title, "Budget review");
  assert.equal(tab?.tabNumber, 2);

  await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false, tabNumber: 7 });
  await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false, tabNumber: 7 });
  const titles = environment.manager.snapshotTabs().map((candidate) => candidate.title);
  assert.deepEqual(titles, ["Chat 1", "Budget review", "Chat 7", "Chat 8"]);
});

void test("a conversation is saved from the user's first message and counts only visible messages", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  runtime.agent.state.messages.push({ role: "system", content: "", timestamp: 0 });

  assert.equal(await runtime.persistence.saveSession(), false, "a blank tab is not saved without force");
  runtime.agent.state.messages.push({ role: "user", content: "first question", timestamp: 1 });
  assert.equal(await runtime.persistence.saveSession(), true);

  const metadata = await environment.sessions.getMetadata(runtime.persistence.getSessionId());
  assert.equal(metadata?.messageCount, 1);
});

void test("a compacted conversation, whose user messages are all archived, is still saved", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const question = { role: "user" as const, content: "first question", timestamp: 1 };
  runtime.agent.state.messages.push(question);
  assert.equal(await runtime.persistence.saveSession(), true);

  runtime.agent.state.messages = [
    createArchivedMessagesMessage({ existingArchivedMessages: [], newlyArchivedMessages: [question], timestamp: 2 }),
    createCompactionSummaryMessage({ summary: "The user asked a question.", tokensBefore: 43_000, timestamp: 3 }),
  ];
  assert.equal(await runtime.persistence.saveSession(), true);

  const saved = await environment.sessions.loadSession(runtime.persistence.getSessionId());
  assert.deepEqual(saved?.messages.map((message) => message.role), ["archivedMessages", "compactionSummary"]);
});

void test("an in-place model switch drops a thinking level the new model cannot use", async () => {
  const environment = createEnvironment();
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  runtime.agent.state.thinkingLevel = "high";

  const result = await environment.manager.selectModel({
    runtimeId: runtime.runtimeId,
    nextModel: getBuiltinModel("openai", "gpt-4.1"),
  });
  assert.equal(result.outcome, "updated");
  assert.equal(runtime.agent.state.thinkingLevel, "off");
});

void test("a reopened session drops a saved thinking level its model cannot use", async () => {
  const environment = createEnvironment();
  await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  const runtime = await environment.manager.createRuntime({ activate: true, autoRestoreLatest: false });
  runtime.agent.state.messages.push({ role: "user", content: "question", timestamp: 1 });
  runtime.agent.state.thinkingLevel = "high";
  await runtime.persistence.saveSession({ force: true });
  const sessionId = runtime.persistence.getSessionId();
  environment.manager.closeRuntime(runtime.runtimeId);

  const loaded = await environment.sessions.loadSession(sessionId);
  assert.ok(loaded);
  const result = await environment.manager.openSession(
    { ...loaded, model: getBuiltinModel("openai", "gpt-4.1-mini") },
    { kind: "new_tab" },
  );
  assert.equal(result.outcome, "opened");
  const reopened = environment.manager.getActiveRuntime();
  assert.equal(reopened?.agent.state.model.id, "gpt-4.1-mini");
  assert.equal(reopened?.agent.state.thinkingLevel, "off");
});
