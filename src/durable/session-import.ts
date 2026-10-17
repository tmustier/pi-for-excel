import type { Context } from "@earendil-works/chord";
import {
  configure, defineDoc, defineDocFamily, defineEntry, ProviderDoc,
  type Conversation, type ConversationId, type Harness,
} from "@earendil-works/pi-durable";

import { createConvertToLlm } from "../messages/convert-to-llm.js";
import type { SessionData } from "../storage/local/types.js";

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
  const messages = createConvertToLlm({ getContextWindow: () => session.model.contextWindow })(session.messages);
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
    await tx.appendEntry(LegacyContextEntry, conversation.id, { model: messages });
    index.conversationId = conversation.id;
    return conversation.id;
  }, context);
  const conversation = await harness.conversation(id, context);
  if (!conversation) throw new Error("Imported conversation is missing");
  return conversation;
}
