import { copyJson, type JsonRepresentation, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, Model, Models, ModelThinkingLevel } from "@earendil-works/pi-ai";
import { configure, defineDoc, type Conversation, type Cursor, type Harness } from "@earendil-works/pi-durable";
import type { SessionData, SessionMetadata } from "../storage/local/types.js";
import { extractTextFromContent } from "../utils/content.js";
import { entryMessages, readConversationHistory } from "./conversation-view.js";
import { LegacySessionDoc } from "./session-import.js";
import { Type } from "typebox";
import { Value } from "typebox/value";
import { ModelCostSchema, ModelInputLimitsSchema, ModelPromptCacheSchema, ThinkingLevelMapSchema } from "@earendil-works/pi-ai/providers/model-schema";

const persistedModelSchema = Type.Object({
  id: Type.String(), name: Type.String(), api: Type.String(), provider: Type.String(), baseUrl: Type.String(),
  reasoning: Type.Boolean(), input: Type.Array(Type.Union([Type.Literal("text"), Type.Literal("image")])),
  cost: ModelCostSchema, contextWindow: Type.Number(), maxTokens: Type.Number(),
  type: Type.Optional(Type.Literal("chat")), headers: Type.Optional(Type.Record(Type.String(), Type.String())),
  inputLimits: Type.Optional(ModelInputLimitsSchema), promptCache: Type.Optional(ModelPromptCacheSchema),
  thinkingLevelMap: Type.Optional(ThinkingLevelMapSchema),
});

function persistedModel(value: JsonValue): Model<Api> | undefined {
  return Value.Check(persistedModelSchema, value) ? value : undefined;
}

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

  constructor(harness: Harness, models: Models, fallbackModel: Model<Api>) {
    this.harness = harness;
    this.models = models;
    this.fallbackModel = fallbackModel;
  }

  private async conversations(): Promise<Conversation[]> {
    const result: Conversation[] = [];
    let cursor: Cursor | undefined;
    do {
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
      const legacy = await this.harness.snapshot(LegacySessionDoc, conversation.id, context);
      if (legacy?.legacySessionId === id) {
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
      info.fallbackModel = copyJson(model, { omitUndefinedProperties: true }) as JsonRepresentation<Model<Api>>;
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
    const messages = history.flatMap((entry) => entryMessages(entry, legacy?.originalSessionJson));
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
      createdAt, lastModified, model, thinkingLevel: agent.thinkingLevel, messages };
  }

  async metadata(id: string): Promise<SessionMetadata | null> {
    const session = await this.loadSession(id);
    if (!session) return null;
    const messages = originals(session.messages);
    const usage: SessionMetadata["usage"] = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } };
    // Originals occur once in the timeline. Never also sum model contributions
    // (which can be reshaped copies of imported originals) or UI archive cards.
    for (const message of messages) {
      if (message.role !== "assistant") continue;
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "totalTokens"] as const) usage[key] += message.usage[key];
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "total"] as const) usage.cost[key] += message.usage.cost[key];
    }
    return { id: session.id, title: session.title, createdAt: session.createdAt, lastModified: session.lastModified,
      thinkingLevel: session.thinkingLevel, usage, messageCount: messages.length,
      preview: messages.filter((message) => message.role === "user" || message.role === "assistant")
        .map((message) => extractTextFromContent(message.content)).join("\n").slice(0, 2048) };
  }

  async getAllMetadata(): Promise<SessionMetadata[]> {
    const result: SessionMetadata[] = [];
    for (const conversation of await this.conversations()) {
      const metadata = await this.metadata(String(conversation.id));
      if (metadata) result.push(metadata);
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
    const session = await this.loadSession(id);
    if (!conversation || !session) throw new Error(`Session not found: ${id}`);
    const active = await conversation.context(context);
    // Context puts the head marker first, even when compaction appended it
    // after the kept tail. Fork at the newest committed entry, not UI order.
    const latest = active.entries.reduce<(typeof active.entries)[number] | undefined>(
      (newest, entry) => !newest || entry.id > newest.id ? entry : newest, undefined);
    if (!latest) throw new Error("Cannot fork a session without a visible entry");
    const selected = model ?? session.model;
    return conversation.fork(latest.id, { ownership: { kind: "ownerless" }, init: async (tx, fork) => {
      await configure(tx, fork, { model: { provider: selected.provider, modelId: selected.id } });
      const info = await tx.doc(SessionInfoDoc, fork);
      info.title = title ?? session.title;
      info.explicitTitle = true;
      info.deleted = false;
      info.createdAt = info.lastModified = new Date().toISOString();
      info.fallbackModel = copyJson(selected, { omitUndefinedProperties: true }) as JsonRepresentation<Model<Api>>;
    } }, context);
  }
}
