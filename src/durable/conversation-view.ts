import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessage, Model, Models, ModelThinkingLevel } from "@earendil-works/pi-ai";
import {
  CompactionEntry, InboxDoc, LiveDoc, ProviderDoc, watchEvents,
  type AgentEvent, type AgentEventStream, type Conversation, type Cursor,
  type EntryRecord, type Harness, type InboxItem, type LiveState,
} from "@earendil-works/pi-durable";
import { COMPACTION_SUMMARY_PREFIX, COMPACTION_SUMMARY_SUFFIX } from "../messages/compaction.js";
import { extractTextFromContent } from "../utils/content.js";
import { LegacyContextEntry, LegacyMessageEntry, LegacySessionDoc, legacyOriginalMessages } from "./session-import.js";

/** Complete immutable timeline, including entries hidden by reset/compaction. */
export async function readConversationHistory(conversation: Conversation): Promise<EntryRecord[]> {
  const entries: EntryRecord[] = [];
  let cursor: Cursor | undefined;
  do {
    const page = await conversation.entries({ order: "ascending" }, 200, cursor, context);
    entries.push(...page.items);
    cursor = page.next;
  } while (cursor !== undefined);
  return entries;
}

/** UI projection only. Model context always comes from Conversation.context(). */
export function entryMessages(entry: EntryRecord, originalSessionJson?: string): AgentMessage[] {
  if (LegacyContextEntry.is(entry) && originalSessionJson) return legacyOriginalMessages(originalSessionJson);
  if (LegacyMessageEntry.is(entry) && entry.data) return [entry.data.message];
  if (CompactionEntry.is(entry)) {
    const summaryMessage = entry.model?.find((message) => message.role === "user");
    const text = summaryMessage ? extractTextFromContent(summaryMessage.content) : "";
    const summary = text.startsWith(COMPACTION_SUMMARY_PREFIX) && text.endsWith(COMPACTION_SUMMARY_SUFFIX)
      ? text.slice(COMPACTION_SUMMARY_PREFIX.length, -COMPACTION_SUMMARY_SUFFIX.length) : text;
    const tokensBefore = entry.data && "tokensBefore" in entry.data && typeof entry.data.tokensBefore === "number"
      ? entry.data.tokensBefore : 0;
    return [{ role: "compactionSummary", summary, tokensBefore, timestamp: summaryMessage?.timestamp ?? 0 }];
  }
  return [...(entry.model ?? [])];
}

export interface ConversationSnapshot {
  readonly messages: AgentMessage[];
  readonly model: Model<Api>;
  readonly thinkingLevel: ModelThinkingLevel;
  readonly busy: boolean;
  readonly streamingMessage: AssistantMessage | undefined;
  readonly pendingToolCalls: Set<string>;
  readonly inbox: readonly InboxItem[];
  readonly live: LiveState;
  readonly providerSessionId: string;
  /** Committed prompt sections; undefined until a prompt baseline is available. */
  readonly systemPrompt?: string | undefined;
}

/** Read-only presentation of committed Durable data, not an execution facade.
 * Consumers submit/configure/abort on the actual Conversation. Events are the
 * upstream Durable contract; no synthetic Agent events or mutable state exist.
 */
export class ConversationPresenter {
  private readonly listeners = new Set<(snapshot: ConversationSnapshot) => void>();
  private readonly eventListeners = new Set<(event: AgentEvent) => void>();
  private stream: AgentEventStream | undefined;
  private history: EntryRecord[] = [];
  private refreshVersion = 0;
  private snapshot: ConversationSnapshot;
  readonly harness: Harness;
  readonly conversation: Conversation;
  private readonly models: Models;

  private constructor(
    harness: Harness,
    conversation: Conversation,
    models: Models,
    fallbackModel: Model<Api>,
  ) {
    this.harness = harness;
    this.conversation = conversation;
    this.models = models;
    this.snapshot = {
      messages: [], model: fallbackModel, thinkingLevel: "off", busy: false,
      streamingMessage: undefined, pendingToolCalls: new Set(), inbox: [], live: {}, providerSessionId: "",
    };
  }

  static async open(harness: Harness, conversation: Conversation, models: Models, fallbackModel: Model<Api>): Promise<ConversationPresenter> {
    const presenter = new ConversationPresenter(harness, conversation, models, fallbackModel);
    presenter.stream = await watchEvents(harness, conversation.id, context);
    await presenter.refresh();
    presenter.stream.start(async (events) => {
      await presenter.refresh();
      for (const event of events) {
        for (const listener of presenter.eventListeners) listener(event);
      }
    });
    return presenter;
  }

  get current(): ConversationSnapshot { return this.snapshot; }

  async refresh(): Promise<void> {
    const version = ++this.refreshVersion;
    const [active, agent, live, inbox, provider, history, legacy] = await Promise.all([
      this.conversation.context(context), this.conversation.agent(context),
      this.harness.snapshot(LiveDoc, this.conversation.id, context),
      this.harness.snapshot(InboxDoc, this.conversation.id, context),
      this.harness.snapshot(ProviderDoc, this.conversation.id, context),
      readConversationHistory(this.conversation),
      this.harness.snapshot(LegacySessionDoc, this.conversation.id, context),
    ]);
    if (version !== this.refreshVersion) return;
    this.history = history;
    const activeIds = new Set(active.entries.map((entry) => entry.id));
    const project = (entry: EntryRecord): AgentMessage[] => entryMessages(entry, legacy?.originalSessionJson);
    const archived = this.history.filter((entry) => !activeIds.has(entry.id)).flatMap(project);
    const messages = active.entries.flatMap(project);
    // Do not flatten imported buckets or discard UI-only originals.
    if (archived.length) messages.unshift({ role: "archivedMessages", archivedMessages: archived,
      archivedChatMessageCount: archived.filter((message) => message.role === "user" || message.role === "assistant"
        || message.role === "user-with-attachments").length, timestamp: 0 });
    const sections = new Map<string, string>();
    let hasPrompt = false;
    for (const message of active.messages) {
      if (message.role !== "system" || !message.sections) continue;
      hasPrompt = true;
      for (const [key, value] of Object.entries(message.sections)) {
        if (value === null) sections.delete(key);
        else sections.set(key, value);
      }
    }
    this.snapshot = {
      messages,
      model: agent.model ? this.models.getModel(agent.model.provider, agent.model.modelId) ?? this.snapshot.model : this.snapshot.model,
      thinkingLevel: agent.thinkingLevel,
      busy: live?.run !== undefined,
      streamingMessage: live?.generation?.message,
      pendingToolCalls: new Set(live?.tools?.filter((tool) => tool.status !== "done").map((tool) => tool.callId)),
      inbox: inbox?.items ?? [], live: live ?? {}, providerSessionId: provider?.sessionId ?? "",
      systemPrompt: hasPrompt ? [...sections.values()].join("\n\n") : undefined,
    };
    for (const listener of this.listeners) listener(this.snapshot);
  }

  subscribe(listener: (snapshot: ConversationSnapshot) => void): () => void {
    this.listeners.add(listener);
    listener(this.snapshot);
    return () => { this.listeners.delete(listener); };
  }

  onEvent(listener: (event: AgentEvent) => void): () => void {
    this.eventListeners.add(listener);
    return () => { this.eventListeners.delete(listener); };
  }

  async dispose(): Promise<void> {
    this.listeners.clear();
    this.eventListeners.clear();
    await this.stream?.stop();
  }
}
