import { copyJson, type Context, type JsonRepresentation, type JsonValue } from "@earendil-works/chord";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import {
  configure, defineDoc, defineDocFamily, defineEntry, ProviderDoc,
  type Conversation, type ConversationId, type Harness,
} from "@earendil-works/pi-durable";

import { createConvertToLlm } from "../messages/convert-to-llm.js";
import type { SessionData } from "../storage/local/types.js";
import type { SessionsStore } from "../storage/local/sessions-store.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";
import { getSessionWorkbookId } from "../workbook/session-association.js";
import { importedMetadata, SessionMetadataDoc, summarizeMessages, sessionModelIdentity, persistedModel } from "./session-metadata.js";

/** Lossless for the JSON-persisted SessionData contract, not arbitrary JS objects.
 * Keep originals separate from the shaped, model-facing import contribution.
 * Durable's conversation record has no title or timestamps of its own.
 */
export const LegacySessionDoc = defineDoc<{
  legacySessionId: string;
  title: string;
  createdAt: string;
  lastModified: string;
  originalSessionJson: string;
}>({
  kind: "excel.legacy-session", version: 1,
  scope: "conversation", history: "latest", fork: "current",
  initial: () => ({
    legacySessionId: "", title: "", createdAt: "", lastModified: "", originalSessionJson: "",
  }),
});

const ImportIndex = defineDocFamily<{ conversationId: ConversationId | null }, null>({
  kind: "excel.legacy-session-import", version: 1, scope: "session", family: true,
  initial: () => ({ conversationId: null }),
});

// Agent default (2026-10-18): 100 messages per cold import page, never a retention cap.
const IMPORT_PAGE_SIZE = 100;

export const LegacyContextEntry = defineEntry("excel.legacy-context");
export const LegacyImportInfoDoc = defineDoc<{
  checked: boolean;
  prototypePages: number;
  archivePages: number;
  model: JsonValue;
}>({
  kind: "excel.legacy-import-info", version: 1, scope: "conversation", history: "latest", fork: "current",
  initial: () => ({ checked: false, prototypePages: 0, archivePages: 0, model: null }),
});

export const LegacyTimelinePage = defineDocFamily<{ messages: JsonRepresentation<AgentMessage>[] }, null>({
  kind: "excel.legacy-timeline-page", version: 1, scope: "conversation", family: true,
  history: "latest", fork: "current", initial: () => ({ messages: [] }),
});

/** Upgrade prototype document-only originals once. The immutable backup is never retired.
 * Pages are UI-only: the original LegacyContextEntry still owns model context.
 */
export async function ensureImportMetadata(harness: Harness, conversation: Conversation, context: Context): Promise<void> {
  if ((await harness.snapshot(LegacyImportInfoDoc, conversation.id, context))?.checked) return;
  await harness.commit(async (tx) => {
    const info = await tx.doc(LegacyImportInfoDoc, conversation.id);
    if (info.checked) return;
    const legacy = await tx.doc(LegacySessionDoc, conversation.id);
    if (legacy.originalSessionJson) {
      const original: JsonValue = copyJson(JSON.parse(legacy.originalSessionJson));
      const model = original && typeof original === "object" && !Array.isArray(original) ? persistedModel(original.model ?? null) : undefined;
      info.model = model ? copyJson(model, { omitUndefinedProperties: true }) : null;
      const messages = originalMessages(original);
      const first = await tx.scanEntries({ conversationId: conversation.id, order: "ascending" }, 1);
      const prototype = first.items.some((entry) => LegacyContextEntry.is(entry));
      if (prototype) {
        let page: JsonRepresentation<AgentMessage>[] = [];
        for (const message of flattenedMessages(messages)) {
          page.push(copyJson(message, { omitUndefinedProperties: true }) as JsonRepresentation<AgentMessage>);
          if (page.length === IMPORT_PAGE_SIZE) {
            (await tx.doc(LegacyTimelinePage, conversation.id, String(info.prototypePages++), null)).messages = page;
            page = [];
          }
        }
        if (page.length) (await tx.doc(LegacyTimelinePage, conversation.id, String(info.prototypePages++), null)).messages = page;
      }
      const summary = await tx.doc(SessionMetadataDoc, conversation.id);
      const imported = summarizeMessages(messages);
      Object.assign(summary, imported, { importedUsage: imported.usage, legacyId: legacy.legacySessionId,
        lastEntryId: Math.max(summary.lastEntryId, first.items[0]?.id ?? 0),
        title: legacy.title || imported.title, createdAt: legacy.createdAt, lastModified: legacy.lastModified });
    }
    info.checked = true;
  }, context);
}
/** Prototype imports stored originals in the document, not in each entry. */
export function legacyOriginalMessages(originalSessionJson: string): AgentMessage[] {
  const original: JsonValue = copyJson(JSON.parse(originalSessionJson));
  return originalMessages(original);
}

function originalMessages(original: JsonValue): AgentMessage[] {
  if (!original || typeof original !== "object" || Array.isArray(original) || !Array.isArray(original.messages)) {
    throw new Error("Invalid legacy session messages");
  }
  // This document was written from SessionData by the original importer. Keep
  // custom message payloads intact rather than replaying its shaped model copy.
  return original.messages as JsonRepresentation<AgentMessage>[];
}

export const LegacyMessageEntry = defineEntry<{
  messageIndex: number;
  message: JsonRepresentation<AgentMessage>;
  archived?: boolean;
}>("excel.legacy-message");

function* flattenedMessages(messages: readonly AgentMessage[]): Generator<AgentMessage> {
  for (const message of messages) {
    if (message.role === "archivedMessages") yield* flattenedMessages(message.archivedMessages);
    else yield message;
  }
}

/** Import every snapshot for this workbook, not only open tabs or metadata rows.
 * A null scope is the unlinked history library, never another workbook's history.
 * Each session is atomic and indexed separately: an interrupted batch can restart
 * without duplicating completed imports. Originals and associations stay untouched.
 */
export async function importWorkbookSessions(
  harness: Harness,
  sessions: Pick<SessionsStore, "getSessionIds" | "get">,
  settings: SettingsWriter,
  workbookId: string | null,
  context: Context,
): Promise<Map<string, ConversationId>> {
  const imported = new Map<string, ConversationId>();
  for (const id of await sessions.getSessionIds()) {
    if (await getSessionWorkbookId(settings, id) !== workbookId) continue;
    const session = await sessions.get(id);
    if (!session) throw new Error(`Legacy session disappeared during migration: ${id}`);
    const conversation = await importLegacySession(harness, session, context);
    imported.set(id, conversation.id);
  }
  return imported;
}

/** Import an already-loaded legacy snapshot; never reads or writes legacy stores.
 * First import wins by legacy ID, even if a later snapshot differs. The index,
 * conversation, provider identity, settings, originals and context share one commit.
 * Call on the single owning Harness; no scheduling or provider request is started.
 */
export async function importLegacySession(
  harness: Harness,
  session: SessionData,
  context: Context,
): Promise<Conversation> {
  const originalSessionJson = JSON.stringify(session);
  const convert = createConvertToLlm({ getContextWindow: () => session.model.contextWindow });
  const id = await harness.commit(async (tx) => {
    const index = await tx.doc(ImportIndex, session.id, null);
    if (index.conversationId !== null) return index.conversationId;

    const conversation = await tx.createConversation({ ownership: { kind: "ownerless" } });
    await configure(tx, conversation.id, {
      model: { provider: session.model.provider, modelId: session.model.id },
      thinkingLevel: session.thinkingLevel,
    });
    // sessions.ts assigns agent.sessionId = SessionData.id. ProviderDoc is the
    // documented stable provider-facing identity, independent of Durable's ID.
    (await tx.doc(ProviderDoc, conversation.id)).sessionId = session.id;
    const original = await tx.doc(LegacySessionDoc, conversation.id);
    original.legacySessionId = session.id;
    original.title = session.title;
    original.createdAt = session.createdAt;
    original.lastModified = session.lastModified;
    original.originalSessionJson = originalSessionJson;
    // One entry per original message gives Durable real cut/fork boundaries.
    // Archive buckets contribute no model content. Their original message data
    // remains in the timeline for rendering/export independently of context.
    let lastEntryId = 0;
    const importInfo = await tx.doc(LegacyImportInfoDoc, conversation.id);
    let archivePage: JsonRepresentation<AgentMessage>[] = [];
    const flushArchive = async (): Promise<void> => {
      if (!archivePage.length) return;
      (await tx.doc(LegacyTimelinePage, conversation.id, `archive-${importInfo.archivePages++}`, null)).messages = archivePage;
      archivePage = [];
    };
    for (const [messageIndex, message] of session.messages.entries()) {
      // SessionData is the typed legacy persistence contract. copyJson validates
      // JSON serializability and detaches it without changing that shape.
      if (message.role === "archivedMessages") {
        for (const archived of flattenedMessages(message.archivedMessages)) {
          archivePage.push(copyJson(archived, { omitUndefinedProperties: true }) as JsonRepresentation<AgentMessage>);
          if (archivePage.length === IMPORT_PAGE_SIZE) await flushArchive();
        }
      }
      const archived = message.role === "archivedMessages";
      const stored = copyJson(archived ? { ...message, archivedMessages: [] } : message,
        { omitUndefinedProperties: true }) as JsonRepresentation<AgentMessage>;
      const entry = await tx.appendEntry(LegacyMessageEntry, conversation.id, {
        data: { messageIndex, message: stored, ...(archived ? { archived: true } : {}) }, model: convert([message]),
      });
      lastEntryId = entry.id;
    }
    await flushArchive();
    Object.assign(await tx.doc(SessionMetadataDoc, conversation.id), importedMetadata(session), { lastEntryId });
    importInfo.checked = true;
    importInfo.model = copyJson(sessionModelIdentity(session.model), { omitUndefinedProperties: true });
    index.conversationId = conversation.id;
    return conversation.id;
  }, context);
  const conversation = await harness.conversation(id, context);
  if (!conversation) throw new Error("Imported conversation is missing");
  return conversation;
}
