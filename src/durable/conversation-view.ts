import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessage, Model, Models, ModelThinkingLevel, Tool } from "@earendil-works/pi-ai";
import {
  CompactionEntry, InboxDoc, LiveDoc, ProviderDoc, watchEvents,
  type AgentEvent, type AgentEventStream, type Conversation, type Cursor,
  type EntryRecord, type Harness, type InboxItem, type LiveState,
} from "@earendil-works/pi-durable";
import { COMPACTION_SUMMARY_PREFIX, COMPACTION_SUMMARY_SUFFIX } from "../messages/compaction.js";
import { extractTextFromContent } from "../utils/content.js";
import { SessionMetadataDoc } from "./session-metadata.js";
import { ensureImportMetadata, LegacyContextEntry, LegacyImportInfoDoc, LegacyMessageEntry, LegacyTimelinePage, legacyOriginalMessages } from "./session-import.js";

/** Complete immutable timeline, including entries hidden by reset/compaction. */
export async function readConversationHistory(conversation: Conversation): Promise<EntryRecord[]> {
  const entries: EntryRecord[] = [];
  let cursor: Cursor | undefined;
  do {
    // Agent default (2026-10-18): 200 entries per explicit export page, not a history limit.
    const page = await conversation.entries({ order: "ascending" }, 200, cursor, context);
    entries.push(...page.items);
    cursor = page.next;
  } while (cursor !== undefined);
  return entries;
}

/** UI projection only. Model context always comes from Conversation.context(). */
export function entryMessages(entry: EntryRecord, originals?: string | readonly AgentMessage[]): AgentMessage[] {
  if (LegacyContextEntry.is(entry) && originals) return typeof originals === "string" ? legacyOriginalMessages(originals) : [...originals];
  if (LegacyMessageEntry.is(entry) && entry.data) {
    if (!entry.data.archived) return [entry.data.message];
    if (!originals) return [];
    const messages = typeof originals === "string" ? legacyOriginalMessages(originals) : originals;
    const original = messages[entry.data.messageIndex];
    if (!original) throw new Error("Missing imported archive original");
    return [original];
  }
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
  readonly hasArchive: boolean;
  readonly model: Model<Api>;
  readonly tools?: readonly Tool[];
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

export interface ArchiveCursor {
  readonly entries?: Cursor;
  readonly prototypePage?: number;
  readonly importPage?: number;
}

/** Read-only presentation of committed Durable data, not an execution facade.
 * Consumers submit/configure/abort on the actual Conversation. Events are the
 * upstream Durable contract; no synthetic Agent events or mutable state exist.
 */
export class ConversationPresenter {
  private readonly listeners = new Set<(snapshot: ConversationSnapshot) => void>();
  private readonly eventListeners = new Set<(event: AgentEvent) => void>();
  private stream: AgentEventStream | undefined;
  private activeIds = new Set<EntryRecord["id"]>();
  private prototypePages = 0;
  private importPages = 0;
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
      messages: [], hasArchive: false, model: fallbackModel, thinkingLevel: "off", busy: false,
      streamingMessage: undefined, pendingToolCalls: new Set(), inbox: [], live: {}, providerSessionId: "",
    };
  }

  static async open(harness: Harness, conversation: Conversation, models: Models, fallbackModel: Model<Api>): Promise<ConversationPresenter> {
    const presenter = new ConversationPresenter(harness, conversation, models, fallbackModel);
    await ensureImportMetadata(harness, conversation, context);
    const imported = await harness.snapshot(LegacyImportInfoDoc, conversation.id, context);
    presenter.prototypePages = imported?.prototypePages ?? 0;
    presenter.importPages = imported?.archivePages ?? 0;
    presenter.stream = await watchEvents(harness, conversation.id, context);
    await presenter.refresh();
    presenter.stream.start(async (events) => {
      const committed = events.some((event) => event.type === "snapshot" || event.type === "message_end" || event.type === "entry_appended");
      await presenter.refresh(committed);
      for (const event of events) {
        for (const listener of presenter.eventListeners) listener(event);
      }
    });
    return presenter;
  }

  get current(): ConversationSnapshot { return this.snapshot; }

  /** Explicit history access is paged and never retained by the live projection. */
  async archivePage(cursor?: ArchiveCursor): Promise<{ messages: AgentMessage[]; next: ArchiveCursor | undefined }> {
    const importPage = cursor?.importPage ?? (cursor === undefined ? 0 : this.importPages);
    if (importPage < this.importPages) {
      const page = await this.harness.snapshot(LegacyTimelinePage, this.conversation.id, `archive-${importPage}`, context);
      if (!page) throw new Error("Missing imported archive page");
      return { messages: page.messages, next: { importPage: importPage + 1 } };
    }
    const prototypePage = cursor?.prototypePage ?? (cursor === undefined ? 0 : this.prototypePages);
    if (prototypePage < this.prototypePages) {
      const page = await this.harness.snapshot(LegacyTimelinePage, this.conversation.id, String(prototypePage), context);
      if (!page) throw new Error("Missing prototype archive page");
      return { messages: page.messages, next: { prototypePage: prototypePage + 1 } };
    }
    // Agent default (2026-10-18): 100 entries per user-requested archive page.
    const page = await this.conversation.entries({ order: "ascending" }, 100, cursor?.entries, context);
    const archived = page.items.filter((entry) => !this.activeIds.has(entry.id) && !LegacyContextEntry.is(entry));
    return { messages: archived.flatMap((entry) => entryMessages(entry)), next: page.next === undefined ? undefined : { entries: page.next } };
  }

  async refresh(committed = true): Promise<void> {
    const version = ++this.refreshVersion;
    const [active, agent, live, inbox, provider, metadata] = await Promise.all([
      committed ? this.conversation.context(context) : undefined, this.conversation.agent(context),
      this.harness.snapshot(LiveDoc, this.conversation.id, context),
      this.harness.snapshot(InboxDoc, this.conversation.id, context),
      this.harness.snapshot(ProviderDoc, this.conversation.id, context),
      committed ? this.harness.snapshot(SessionMetadataDoc, this.conversation.id, context) : undefined,
    ]);
    if (version !== this.refreshVersion) return;
    if (active) this.activeIds = new Set(active.entries.map((entry) => entry.id));
    const messages = active ? active.entries.flatMap((entry) => entryMessages(entry)) : this.snapshot.messages;
    const sections = new Map<string, string>();
    let hasPrompt = false;
    for (const message of active?.messages ?? []) {
      if (message.role !== "system" || !message.sections) continue;
      hasPrompt = true;
      for (const [key, value] of Object.entries(message.sections)) {
        if (value === null) sections.delete(key);
        else sections.set(key, value);
      }
    }
    this.snapshot = {
      messages,
      hasArchive: active ? this.prototypePages > 0 || this.importPages > 0 || active.entries.some((entry) => entry.head !== undefined)
        || (metadata?.messageCount ?? 0) > messages.length : this.snapshot.hasArchive,
      model: agent.model ? this.models.getModel(agent.model.provider, agent.model.modelId) ?? this.snapshot.model : this.snapshot.model,
      tools: agent.tools.map(({ name, description, parameters }) => ({ name, description, parameters })),
      thinkingLevel: agent.thinkingLevel,
      busy: live?.run !== undefined,
      streamingMessage: live?.generation?.message,
      pendingToolCalls: new Set(live?.tools?.filter((tool) => tool.status !== "done").map((tool) => tool.callId)),
      inbox: inbox?.items ?? [], live: live ?? {}, providerSessionId: provider?.sessionId ?? "",
      systemPrompt: active ? (hasPrompt ? [...sections.values()].join("\n\n") : undefined) : this.snapshot.systemPrompt,
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
