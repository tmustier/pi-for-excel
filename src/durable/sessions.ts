import { copyJson, type JsonRepresentation, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, Model, Models, ModelThinkingLevel } from "@earendil-works/pi-ai";
import { configure, defineDoc, UsageDoc, type Conversation, type Cursor, type EntryId, type Harness } from "@earendil-works/pi-durable";
import type { SessionData, SessionMetadata } from "../storage/local/types.js";
import { extractTextFromContent } from "../utils/content.js";
import { entryMessages, readConversationHistory } from "./conversation-view.js";
import { ensureImportMetadata, LegacyContextEntry, LegacyImportInfoDoc, LegacyMessageEntry, LegacySessionDoc, legacyOriginalMessages } from "./session-import.js";
import { indexCommittedEntries, SessionMetadataDoc, sessionModelIdentity, persistedModel } from "./session-metadata.js";

export const SessionInfoDoc = defineDoc<{
  title: string;
  explicitTitle: boolean;
  createdAt: string;
  lastModified: string;
  deleted: boolean;
  fallbackModel: JsonRepresentation<Model<Api>> | null;
}>({
  kind: "excel.session-info", version: 1, scope: "conversation", history: "latest", fork: "current",
  initial: () => ({ title: "", explicitTitle: false, createdAt: "", lastModified: "", deleted: false, fallbackModel: null }),
});

function originals(messages: readonly AgentMessage[]): AgentMessage[] {
  return messages.flatMap((message) => message.role === "archivedMessages" ? originals(message.archivedMessages) : [message]);
}

/** Durable history catalogue. IDs exposed to callers are decimal Conversation IDs;
 * legacy IDs remain accepted aliases. Export is deliberately NOT the active UI projection.
 * No read failures are treated as missing data, and no legacy store is accessed.
 */
export class DurableSessions {
  private readonly harness: Harness;
  private readonly models: Models;
  private readonly fallbackModel: Model<Api>;
  private indexing: Promise<void> = Promise.resolve();
  private readonly unsubscribe: () => void;

  constructor(harness: Harness, models: Models, fallbackModel: Model<Api>) {
    this.harness = harness;
    this.models = models;
    this.fallbackModel = fallbackModel;
    this.unsubscribe = harness.subscribeCommits((publication) => {
      const entries = publication.changes.flatMap((change) => change.type === "entry" ? [change.value] : []);
      if (entries.length) {
        this.indexing = this.indexing.then(() => indexCommittedEntries(harness, entries, entryMessages));
        // Keep a failed index failed (metadata() surfaces it), without an
        // unhandled rejection when nobody is currently listing sessions.
        void this.indexing.catch(() => {});
      }
    });
  }

  dispose(): void { this.unsubscribe(); }

  /** Explicit startup migration/recovery, before allowing conversation writes.
   * Only the unindexed tail is read, one page at a time; no transcript cache is
   * built. Kept separate from metadata()/header() so tab/list reads stay cold.
   */
  async reconcileMetadata(): Promise<void> {
    await this.indexing;
    for (const conversation of await this.conversations()) {
      await ensureImportMetadata(this.harness, conversation, context);
      const summary = await this.harness.snapshot(SessionMetadataDoc, conversation.id, context);
      const minEntryId = ((summary?.lastEntryId ?? 0) + 1) as EntryId;
      let cursor: Cursor | undefined;
      do {
        // agent-default (not a user rule): 2026-10-18. Repair the index in 200-entry pages, without a history cap.
        const page = await conversation.entries({ minEntryId, order: "ascending" }, 200, cursor, context);
        if (page.items.length) await indexCommittedEntries(this.harness, page.items, (entry) =>
          summary?.legacyId && (LegacyMessageEntry.is(entry) || LegacyContextEntry.is(entry)) ? [] : entryMessages(entry));
        cursor = page.next;
      } while (cursor !== undefined);
    }
  }

  private async conversations(): Promise<Conversation[]> {
    const result: Conversation[] = [];
    let cursor: Cursor | undefined;
    do {
      // agent-default (not a user rule): 2026-10-18. Scan 200 catalogue rows per page.
      const page = await this.harness.commit((tx) => tx.scanConversations({}, 200, cursor), context);
      for (const record of page.items) {
        if (record.owner) continue;
        const conversation = await this.harness.conversation(record.id, context);
        if (!conversation) throw new Error(`Conversation disappeared: ${record.id}`);
        result.push(conversation);
      }
      cursor = page.next;
    } while (cursor !== undefined);
    return result;
  }

  async resolve(id: string): Promise<Conversation | null> {
    if (/^(0|[1-9]\d*)$/.test(id) && Number.isSafeInteger(Number(id))) {
      const record = await this.harness.commit((tx) => tx.conversation(Number(id) as Conversation["id"]), context);
      if (record) return record.owner ? null : (await this.harness.conversation(record.id, context)) ?? null;
    }
    const conversations = await this.conversations();
    const direct = conversations.find((conversation) => String(conversation.id) === id);
    if (direct) return direct;
    for (const conversation of conversations) {
      // Forks inherit import docs, but only the original owns the old tab ID.
      await ensureImportMetadata(this.harness, conversation, context);
      const summary = await this.harness.snapshot(SessionMetadataDoc, conversation.id, context);
      if (summary?.legacyId === id) {
        const raw = await this.harness.commit((tx) => tx.conversation(conversation.id), context);
        if (!raw?.parent) return conversation;
      }
    }
    return null;
  }

  async create(model: Model<Api>, thinkingLevel: ModelThinkingLevel, title?: string): Promise<Conversation> {
    return this.harness.createConversation({ ownership: { kind: "ownerless" }, init: async (tx, conversation) => {
      await configure(tx, conversation, { model: { provider: model.provider, modelId: model.id }, thinkingLevel });
      const info = await tx.doc(SessionInfoDoc, conversation);
      info.title = title ?? "";
      info.explicitTitle = title !== undefined;
      info.createdAt = info.lastModified = new Date().toISOString();
      info.fallbackModel = copyJson(sessionModelIdentity(model), { omitUndefinedProperties: true }) as JsonRepresentation<Model<Api>>;
    } }, context);
  }

  async loadSession(id: string): Promise<SessionData | null> {
    const conversation = await this.resolve(id);
    if (!conversation) return null;
    const [info, legacy, agent, history] = await Promise.all([
      this.harness.snapshot(SessionInfoDoc, conversation.id, context),
      this.harness.snapshot(LegacySessionDoc, conversation.id, context),
      conversation.agent(context), readConversationHistory(conversation),
    ]);
    if (info?.deleted) return null;
    const importedMessages = legacy?.originalSessionJson && history.some((entry) => LegacyContextEntry.is(entry)
      || (LegacyMessageEntry.is(entry) && entry.data?.archived)) ? legacyOriginalMessages(legacy.originalSessionJson) : undefined;
    const messages = history.flatMap((entry) => entryMessages(entry, importedMessages));
    const flat = originals(messages);
    const first = flat.find((message) => message.role === "user" || message.role === "user-with-attachments");
    const firstText = first && "content" in first ? extractTextFromContent(first.content) : "";
    const timestamps = flat.map((message) => message.timestamp).filter((time) => Number.isFinite(time));
    const createdAt = info?.createdAt || legacy?.createdAt || new Date(timestamps[0] ?? 0).toISOString();
    const lastModified = [info?.lastModified || legacy?.lastModified || createdAt,
      ...timestamps.map((time) => new Date(time).toISOString())].sort().at(-1) ?? createdAt;
    // This JSON was written by importLegacySession from typed SessionData. Only
    // its model fallback is read; history always comes from immutable entries.
    const original: JsonValue = legacy?.originalSessionJson ? copyJson(JSON.parse(legacy.originalSessionJson)) : null;
    const importedModel = original && typeof original === "object" && !Array.isArray(original) ? original.model : null;
    const fallback = persistedModel(info?.fallbackModel ?? null) ?? persistedModel(importedModel ?? null) ?? this.fallbackModel;
    const model = (agent.model ? this.models.getModel(agent.model.provider, agent.model.modelId) : undefined)
      ?? fallback;
    return { id: String(conversation.id), title: info?.explicitTitle ? info.title : legacy?.title || firstText.slice(0, 80),
      createdAt, lastModified, model: sessionModelIdentity(model), thinkingLevel: agent.thinkingLevel, messages };
  }

  /** Initial tabs need identity/configuration, not an exported transcript. */
  async header(id: string): Promise<Omit<SessionData, "messages"> | null> {
    const metadata = await this.metadata(id);
    if (!metadata) return null;
    const conversation = await this.resolve(metadata.id);
    if (!conversation) return null;
    const [agent, info, imported] = await Promise.all([conversation.agent(context),
      this.harness.snapshot(SessionInfoDoc, conversation.id, context),
      this.harness.snapshot(LegacyImportInfoDoc, conversation.id, context)]);
    const model = (agent.model ? this.models.getModel(agent.model.provider, agent.model.modelId) : undefined)
      ?? persistedModel(info?.fallbackModel ?? null) ?? persistedModel(imported?.model ?? null) ?? this.fallbackModel;
    return { id: metadata.id, title: metadata.title, createdAt: metadata.createdAt, lastModified: metadata.lastModified,
      model, thinkingLevel: metadata.thinkingLevel };
  }

  async metadata(id: string): Promise<SessionMetadata | null> {
    await this.indexing;
    const conversation = await this.resolve(id);
    if (!conversation) return null;
    await ensureImportMetadata(this.harness, conversation, context);
    const [info, summary, agent, nativeUsage] = await Promise.all([
      this.harness.snapshot(SessionInfoDoc, conversation.id, context),
      this.harness.snapshot(SessionMetadataDoc, conversation.id, context),
      conversation.agent(context), this.harness.snapshot(UsageDoc, conversation.id, context),
    ]);
    if (info?.deleted) return null;
    const data = summary ?? SessionMetadataDoc.definition.initial();
    // Native spend includes summarization attempts which have no message entry.
    // Prefer the native ledger when present; imported transcripts seed the index.
    const hasNativeUsage = nativeUsage && (Object.keys(nativeUsage.models).length > 0 || Object.keys(nativeUsage.tools).length > 0);
    const baseline = hasNativeUsage ? data.importedUsage ?? SessionMetadataDoc.definition.initial().usage : data.usage;
    const usage = { ...baseline, cost: { ...baseline.cost } };
    if (hasNativeUsage) for (const bucket of [nativeUsage.models, nativeUsage.tools]) for (const item of Object.values(bucket)) {
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "totalTokens"] as const) usage[key] += item[key];
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "total"] as const) usage.cost[key] += item.cost[key];
    }
    return { id: String(conversation.id), title: info?.explicitTitle ? info.title : data.title,
      createdAt: info?.createdAt || data.createdAt, lastModified: [info?.lastModified ?? "", data.lastModified].sort().at(-1) ?? "",
      thinkingLevel: agent.thinkingLevel, usage, messageCount: data.messageCount, preview: data.preview };
  }

  async getAllMetadata(): Promise<SessionMetadata[]> {
    const result: SessionMetadata[] = [];
    for (const conversation of await this.conversations()) {
      const metadata = await this.metadata(String(conversation.id));
      // Empty tabs remain addressable for restoration, but are not history.
      // Explicitly named drafts have a title and remain discoverable.
      if (metadata && (metadata.messageCount > 0 || metadata.title.trim())) result.push(metadata);
    }
    return result.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
  }

  async renameSession(id: string, title: string): Promise<void> {
    const conversation = await this.resolve(id);
    if (!conversation) throw new Error(`Session not found: ${id}`);
    await this.harness.commit(async (tx) => {
      const info = await tx.doc(SessionInfoDoc, conversation.id);
      info.title = title;
      info.explicitTitle = true;
      info.lastModified = new Date().toISOString();
    }, context);
  }

  async delete(id: string): Promise<void> {
    const conversation = await this.resolve(id);
    if (!conversation) return;
    await this.harness.commit(async (tx) => {
      const info = await tx.doc(SessionInfoDoc, conversation.id);
      info.deleted = true;
      info.lastModified = new Date().toISOString();
    }, context);
  }

  async fork(id: string, model?: Model<Api>, title?: string): Promise<Conversation> {
    const conversation = await this.resolve(id);
    const session = await this.header(id);
    if (!conversation || !session) throw new Error(`Session not found: ${id}`);
    const active = await conversation.context(context);
    // Context puts the head marker first, even when compaction appended it
    // after the kept tail. Fork at the newest committed entry, not UI order.
    const latest = active.entries.reduce<(typeof active.entries)[number] | undefined>(
      (newest, entry) => !newest || entry.id > newest.id ? entry : newest, undefined);
    if (!latest) throw new Error("Cannot fork a session without a visible entry");
    const selected = model ?? session.model;
    const inheritedUsage = (await this.metadata(id))?.usage;
    return conversation.fork(latest.id, { ownership: { kind: "ownerless" }, init: async (tx, fork) => {
      await configure(tx, fork, { model: { provider: selected.provider, modelId: selected.id } });
      const summary = await tx.doc(SessionMetadataDoc, fork);
      // pi.usage starts empty on a fork; inherited transcript spend is its baseline.
      const baseline = inheritedUsage ?? summary.usage;
      summary.importedUsage = { ...baseline, cost: { ...baseline.cost } };
      const info = await tx.doc(SessionInfoDoc, fork);
      info.title = title ?? session.title;
      info.explicitTitle = true;
      info.deleted = false;
      info.createdAt = info.lastModified = new Date().toISOString();
      info.fallbackModel = copyJson(sessionModelIdentity(selected), { omitUndefinedProperties: true }) as JsonRepresentation<Model<Api>>;
    } }, context);
  }
}
