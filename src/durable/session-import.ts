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

export const LegacyContextEntry = defineEntry("excel.legacy-context");
/** Prototype imports stored originals in the document, not in each entry. */
export function legacyOriginalMessages(originalSessionJson: string): AgentMessage[] {
  const original: JsonValue = copyJson(JSON.parse(originalSessionJson));
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
}>("excel.legacy-message");

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
    for (const [messageIndex, message] of session.messages.entries()) {
      // SessionData is the typed legacy persistence contract. copyJson validates
      // JSON serializability and detaches it without changing that shape.
      const stored = copyJson(message, { omitUndefinedProperties: true }) as JsonRepresentation<AgentMessage>;
      await tx.appendEntry(LegacyMessageEntry, conversation.id, {
        data: { messageIndex, message: stored }, model: convert([message]),
      });
    }
    index.conversationId = conversation.id;
    return conversation.id;
  }, context);
  const conversation = await harness.conversation(id, context);
  if (!conversation) throw new Error("Imported conversation is missing");
  return conversation;
}
