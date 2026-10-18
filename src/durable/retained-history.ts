import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { Api, Model, Models } from "@earendil-works/pi-ai";
import { createRegistry, Harness, type ConversationId, type Cursor } from "@earendil-works/pi-durable";
import type { SettingsWriter } from "../storage/local/settings-store.js";
import type { DurableHistoryCallbacks } from "../ui/durable-controls.js";
import { forgetRetainedDurableAssociation, listRetainedDurableAssociations,
  type DurableWorkbookAssociation } from "../workbook/durable-association.js";
import { openIndexedDbStorage, deleteDurableWorkbookStorage } from "./indexeddb-storage.js";
import { importLegacySession } from "./session-import.js";
import { DurableSessions } from "./sessions.js";
import { MutationDoc, PendingMutationsDoc } from "./runtime.js";
import { SessionMetadataDoc } from "./session-metadata.js";

/** Foreign generations are opened paused, without extensions or execution capabilities.
 * No progress-requesting Harness API may be called here. The caller owns current.
 */
export function createRetainedHistoryCallbacks(options: {
  current: DurableWorkbookAssociation;
  harness: Harness;
  sessions: DurableSessions;
  models: Models;
  defaultModel: Model<Api>;
  settings: SettingsWriter;
  /** Storage seams allow the same callback contract to be exercised in memory. */
  openStorage?: typeof openIndexedDbStorage;
  deleteStorage?: typeof deleteDurableWorkbookStorage;
}): DurableHistoryCallbacks {
  const find = async (id: string) => {
    const row = (await listRetainedDurableAssociations(options.settings)).find((item) => item.scopeId === id);
    if (!row?.databaseId) throw new Error("Retained generation is unavailable");
    return { ...row, databaseId: row.databaseId };
  };
  const visit = async <T>(id: string, run: (harness: Harness, sessions: DurableSessions) => Promise<T>): Promise<T> => {
    if (id === options.current.scopeId) return run(options.harness, options.sessions);
    const row = await find(id);
    const storage = await (options.openStorage ?? openIndexedDbStorage)(`pi-for-excel-durable:${id}`, {
      expectedDatabaseId: row.databaseId,
    });
    let harness: Harness | undefined;
    let sessions: DurableSessions | undefined;
    try {
      harness = await Harness.open(storage, { models: options.models, registry: createRegistry() }, context);
      if ((await harness.inspect(context)).scheduling !== "paused") throw new Error("History must remain paused");
      sessions = new DurableSessions(harness, options.models, options.defaultModel);
      // Prototype catalogues predate the metadata index. Repair only when a
      // document is missing; ordinary retained listings never read transcripts.
      let cursor: Cursor | undefined;
      let needsRepair = false;
      do {
        // Agent default (2026-10-18): 200 metadata checks per retained catalogue page.
        const page = await harness.commit((tx) => tx.scanConversations({}, 200, cursor), context);
        for (const conversation of page.items) {
          if (!conversation.owner && !await harness.snapshot(SessionMetadataDoc, conversation.id, context)) needsRepair = true;
        }
        cursor = page.next;
      } while (cursor !== undefined);
      if (needsRepair) await sessions.reconcileMetadata();
      return await run(harness, sessions);
    } finally {
      sessions?.dispose();
      if (harness) await harness.close(context);
      else await storage.close(context);
    }
  };
  const completion = async (harness: Harness) => {
    const inspection = await harness.inspect(context);
    const pending = await harness.snapshot(PendingMutationsDoc, context);
    let uncertain = pending?.migrated !== true;
    const blocked = new Set<string>();
    const block = async (id: ConversationId): Promise<void> => {
      const seen = new Set<ConversationId>();
      let next: ConversationId | undefined = id;
      while (next !== undefined) {
        if (seen.has(next)) { uncertain = true; return; }
        seen.add(next);
        const record = await harness.commit((tx) => tx.conversation(next as ConversationId), context);
        if (!record) { uncertain = true; return; }
        blocked.add(String(record.id));
        next = record.owner?.conversationId;
      }
    };
    for (const task of inspection.tasks) await block(task.record.conversationId);
    for (const submission of inspection.submissions) await block(submission.conversationId);
    for (const id of Object.keys(pending?.ids ?? {})) {
      const mutation = await harness.snapshot(MutationDoc, id, context);
      if (!mutation?.binding) uncertain = true;
      else await block(mutation.binding.conversationId);
    }
    return (id: string): boolean => !uncertain && !blocked.has(id);
  };
  return {
    listScopes: async () => (await listRetainedDurableAssociations(options.settings))
      .flatMap((row) => row.databaseId ? [{ id: row.scopeId, generation: row.databaseId, name: row.name,
        current: row.scopeId === options.current.scopeId,
        identityDurable: row.canonicalId !== null || row.scopeId.startsWith("doc_instance:") }] : []),
    inspectSessions: (id) => visit(id, async (harness, sessions) => {
      const completed = await completion(harness);
      return (await sessions.getAllMetadata()).map((session) => ({ id: session.id, name: session.title, completed: completed(session.id) }));
    }),
    copyCompletedHistory: async (id, sessionId) => {
      if (id === options.current.scopeId) throw new Error("Cannot copy current workbook history into itself");
      const row = await find(id);
      await visit(id, async (harness, sessions) => {
        const conversation = await sessions.resolve(sessionId);
        if (!conversation) throw new Error("Retained session is unavailable");
        if (!(await completion(harness))(String(conversation.id))) throw new Error("Pending or uncertain work cannot be copied");
        const session = await sessions.loadSession(sessionId);
        if (!session) throw new Error("Retained session is unavailable");
        // Structured provenance avoids collisions across generations and gives
        // importLegacySession its atomic first-import-wins deduplication key.
        await importLegacySession(options.harness, { ...session,
          id: JSON.stringify(["retained-workbook", row.scopeId, row.databaseId, session.id]) }, context);
      });
    },
    deleteGeneration: async (id) => {
      if (id === options.current.scopeId) throw new Error("Cannot delete current workbook generation");
      const row = await find(id);
      await (options.deleteStorage ?? deleteDurableWorkbookStorage)({ name: `pi-for-excel-durable:${id}`,
        databaseId: row.databaseId, authorization: "user-confirmed-obsolete-workbook" });
      // Failure leaves the expectation intact, never silently admits a fresh DB.
      await forgetRetainedDurableAssociation(options.settings, id, row.databaseId);
    },
  };
}
