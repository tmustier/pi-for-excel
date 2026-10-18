import { defineDoc, type EntryRecord, type Harness } from "@earendil-works/pi-durable";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { Api, Model } from "@earendil-works/pi-ai";
import type { SessionData, SessionMetadata } from "../storage/local/types.js";
import { extractTextFromContent } from "../utils/content.js";
import type { JsonValue } from "@earendil-works/chord";
import { Type } from "typebox";
import { Value } from "typebox/value";
import { ModelCostSchema } from "@earendil-works/pi-ai/providers/model-schema";

const persistedModelSchema = Type.Object({
  id: Type.String(), name: Type.String(), api: Type.String(), provider: Type.String(), baseUrl: Type.String(),
  reasoning: Type.Boolean(), input: Type.Array(Type.Union([Type.Literal("text"), Type.Literal("image")])),
  cost: ModelCostSchema, contextWindow: Type.Number(), maxTokens: Type.Number(),
});

export function persistedModel(value: JsonValue): Model<Api> | undefined {
  return Value.Check(persistedModelSchema, value) ? sessionModelIdentity(value) : undefined;
}

/** Persistence/export fallback, never a routing configuration. Credentials and
 * endpoint URLs belong exclusively to the runtime Models registry. */
export function sessionModelIdentity(model: Model<Api>): Model<Api> {
  return { id: model.id, name: model.name, api: model.api, provider: model.provider,
    baseUrl: "", reasoning: model.reasoning, input: [...model.input],
    cost: { input: model.cost.input, output: model.cost.output,
      cacheRead: model.cost.cacheRead, cacheWrite: model.cost.cacheWrite },
    contextWindow: model.contextWindow, maxTokens: model.maxTokens };
}

export const SessionMetadataDoc = defineDoc<{
  legacyId: string;
  title: string;
  createdAt: string;
  lastModified: string;
  preview: string;
  messageCount: number;
  usage: SessionMetadata["usage"];
  importedUsage: SessionMetadata["usage"] | null;
  lastEntryId: number;
}>({
  kind: "excel.session-metadata", version: 1, scope: "conversation", history: "latest", fork: "current",
  initial: () => ({ legacyId: "", title: "", createdAt: "", lastModified: "", preview: "", messageCount: 0,
    lastEntryId: 0, importedUsage: null, usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } } }),
});

export function summarizeMessages(messages: readonly AgentMessage[], summary = SessionMetadataDoc.definition.initial()): ReturnType<typeof SessionMetadataDoc.definition.initial> {
  for (const message of messages) {
    if (message.role === "archivedMessages") { summarizeMessages(message.archivedMessages, summary); continue; }
    summary.messageCount++;
    const time = Number.isFinite(message.timestamp) ? new Date(message.timestamp).toISOString() : "";
    if (!summary.createdAt) summary.createdAt = time;
    if (time > summary.lastModified) summary.lastModified = time;
    if (message.role === "user" || message.role === "user-with-attachments" || message.role === "assistant") {
      const text = extractTextFromContent(message.content);
      if (!summary.title && message.role !== "assistant") summary.title = text.slice(0, 80);
      if (summary.preview.length < 2048) summary.preview = (summary.preview ? `${summary.preview}\n${text}` : text).slice(0, 2048);
    }
    if (message.role === "assistant") {
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "totalTokens"] as const) summary.usage[key] += message.usage[key];
      for (const key of ["input", "output", "cacheRead", "cacheWrite", "total"] as const) summary.usage.cost[key] += message.usage.cost[key];
    }
  }
  return summary;
}

export function importedMetadata(session: SessionData): ReturnType<typeof SessionMetadataDoc.definition.initial> {
  const summary = summarizeMessages(session.messages);
  return { ...summary, importedUsage: summary.usage, legacyId: session.id, title: session.title || summary.title,
    createdAt: session.createdAt, lastModified: session.lastModified };
}

/** Only committed entries are indexed; partial publications never touch the catalogue. */
export function indexCommittedEntries(harness: Harness, entries: readonly EntryRecord[], project: (entry: EntryRecord) => AgentMessage[]): Promise<void> {
  return harness.commit(async (tx) => {
    for (const entry of [...entries].sort((a, b) => a.id - b.id)) {
      const summary = await tx.doc(SessionMetadataDoc, entry.conversationId);
      if (entry.id <= summary.lastEntryId) continue;
      Object.assign(summary, summarizeMessages(project(entry), summary));
      summary.lastEntryId = entry.id;
    }
  }, context);
}
