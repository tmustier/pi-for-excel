import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { JsonValue } from "@earendil-works/chord";
import type { Api, Model, Models } from "@earendil-works/pi-ai";
import { Type } from "typebox";
import { Value } from "typebox/value";
import type { createOfficeStreamFn } from "../auth/stream-proxy.js";
import type { SpreadsheetHost } from "../host/types.js";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { SessionRuntime } from "../taskpane/session-runtime-manager.js";
import type { WorkbookCoordinator } from "../workbook/coordinator.js";
import { createDurableWorkbookAssociationResolver, officeDurableDocumentAssociation,
  type DurableWorkbookAssociation } from "../workbook/durable-association.js";
import { openIndexedDbStorage } from "./indexeddb-storage.js";
import { openDurableRuntime, routeModels } from "./runtime.js";
import { importLegacySession, importWorkbookSessions } from "./session-import.js";
import { createModelRecoveryLifecycle, type ModelRecoveryLifecycle } from "./model-recovery.js";
import { createWorkbookMutationHost } from "./mutation-host.js";
import { DurableSessions } from "./sessions.js";

const payloadSchema = Type.Object({
  text: Type.Optional(Type.String({ minLength: 1 })),
  requestId: Type.Optional(Type.String({ minLength: 1 })),
  point: Type.Optional(Type.Union([Type.Literal("beforeDispatch"), Type.Literal("afterDispatch"), Type.Literal("beforeResult"), Type.Null()])),
  mutationId: Type.Optional(Type.String({ minLength: 1 })),
  observationToken: Type.Optional(Type.String({ minLength: 1 })),
  sessionId: Type.Optional(Type.String({ minLength: 1 })),
}, { additionalProperties: false });

export interface DurableTaskpane {
  runtime: Awaited<ReturnType<typeof openDurableRuntime>>;
  sessions: DurableSessions;
  association: DurableWorkbookAssociation;
  resolveAssociation(): Promise<DurableWorkbookAssociation>;
  execute(command: string, payloadJson: string): Promise<JsonValue>;
}

/** The single workbook execution owner, also used by the verification bridge. */
export async function createDurableTaskpane(options: {
  getActiveRuntime: () => SessionRuntime | null;
  models: Models;
  streamFn: ReturnType<typeof createOfficeStreamFn>;
  host: SpreadsheetHost;
  coordinator: WorkbookCoordinator;
  defaultModel: Model<Api>;
  autoCompaction: boolean;
  autoResizeImages?: boolean;
  allowFaultInjection: boolean;
  onStorageWarning?: (message: string) => void;
}): Promise<DurableTaskpane> {
  const app = getAppStorage();
  const resolveAssociation = createDurableWorkbookAssociationResolver({
    host: options.host, settings: app.settings,
    document: options.host.kind === "office" ? officeDurableDocumentAssociation() : null,
  });
  const association = await resolveAssociation();
  const storage = await openIndexedDbStorage(association.databaseName,
    association.expectedDatabaseId === null ? {} : { expectedDatabaseId: association.expectedDatabaseId });
  let lifecycle: ModelRecoveryLifecycle | undefined;
  try {
    // Failure is fatal: never admit a task until its generation expectation is retained.
    await association.rememberGeneration(storage.databaseId);
    if (storage.persistence !== "persistent" || !association.documentBacked) {
      options.onStorageWarning?.("Chat history uses browser storage. Keep exports: browser data clearing can remove history; persistence is not a backup.");
    }
    const runtime = await openDurableRuntime({
      storage, models: routeModels(options.models, options.streamFn, () => lifecycle),
      workbookId: association.scopeId, coordinator: options.coordinator,
      mutationHost: createWorkbookMutationHost({ kind: options.host.kind,
        getWorkbookId: async () => (await resolveAssociation()).scopeId }),
      autoCompaction: options.autoCompaction, autoResizeImages: options.autoResizeImages ?? true,
      allowFaultInjection: options.allowFaultInjection,
      onOpened: (harness) => {
        lifecycle = createModelRecoveryLifecycle(() => harness.close(context));
        const bound = lifecycle;
        void harness.closed.then(() => bound.dispose());
      },
    });
    try {
      const workbook = await options.host.getWorkbookContext();
      await importWorkbookSessions(runtime.harness, app.sessions, app.settings, workbook.workbookId, context);
      const sessions = new DurableSessions(runtime.harness, options.models, options.defaultModel);
      await sessions.reconcileMetadata();
      void runtime.harness.closed.then(() => sessions.dispose());
      return { runtime, sessions, association, resolveAssociation,
        execute: async (command, payloadJson) => {
          const payload: unknown = JSON.parse(payloadJson);
          if (!Value.Check(payloadSchema, payload)) throw new Error("Invalid Durable command payload");
          if ((await resolveAssociation()).scopeId !== association.scopeId) throw new Error("Workbook changed; reload before continuing");
          const active = options.getActiveRuntime();
          if (!active) throw new Error("No active conversation");
          const conversation = active.view.conversation;
          switch (command) {
            case "durableStatus": return runtime.status(conversation);
            case "durableFault":
              if (payload.point === undefined) throw new Error("durableFault requires point");
              runtime.mutations.configureFault(payload.point);
              return runtime.mutations.status();
            case "durableResumeFault": runtime.mutations.resume(); return { resumed: true };
            case "durableSubmit": {
              if (!payload.text || !payload.requestId) throw new Error("durableSubmit requires text and requestId");
              await active.refreshCapabilities();
              const submission = await conversation.submit({ type: "input", content: payload.text, requestId: payload.requestId }, context);
              return { submissionId: submission.id };
            }
            case "durableImport": {
              if (!payload.sessionId) throw new Error("durableImport requires sessionId");
              const saved = await app.sessions.get(payload.sessionId);
              if (!saved) throw new Error("Legacy session not found");
              return { conversationId: (await importLegacySession(runtime.harness, saved, context)).id };
            }
            case "durableReconcile": {
              if (!payload.mutationId) throw new Error("durableReconcile requires mutationId");
              return { ...await runtime.mutations.reconcile(payload.mutationId), wroteCells: false };
            }
            default: throw new Error("Unknown Durable command");
          }
        },
      };
    } catch (error) { await runtime.harness.close(context); throw error; }
  } catch (error) { await storage.close(context); throw error; }
}
