import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { JsonValue } from "@earendil-works/chord";
import type { Models } from "@earendil-works/pi-ai/models";
import { Type } from "typebox";
import { Value } from "typebox/value";
import type { createOfficeStreamFn } from "../auth/stream-proxy.js";
import { excelRun } from "../excel/helpers.js";
import { isExperimentalFeatureEnabled } from "../experiments/flags.js";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { SessionRuntime } from "../taskpane/session-runtime-manager.js";
import type { WorkbookContext } from "../workbook/context.js";
import { openIndexedDbStorage } from "./indexeddb-storage.js";
import { openDurableRuntime, routeModels } from "./runtime.js";
import { importLegacySession } from "./session-import.js";
import { createModelRecoveryLifecycle, type ModelRecoveryLifecycle } from "./model-recovery.js";

const payloadSchema = Type.Object({
  text: Type.Optional(Type.String({ minLength: 1 })),
  requestId: Type.Optional(Type.String({ minLength: 1 })),
  point: Type.Optional(Type.Union([Type.Literal("beforeSync"), Type.Literal("afterSync"), Type.Literal("beforeResult"), Type.Null()])),
  toolCallId: Type.Optional(Type.String({ minLength: 1 })),
  sessionId: Type.Optional(Type.String({ minLength: 1 })),
}, { additionalProperties: false });
const gridSchema = Type.Array(Type.Array(Type.Union([Type.String(), Type.Number(), Type.Boolean(), Type.Null()])));

export interface DurableTaskpane {
  execute(command: string, payloadJson: string): Promise<JsonValue>;
}

/** Evaluation controller, separate from the normal chat Agent and its persistence. */
export async function createDurableTaskpane(options: {
  getActiveRuntime: () => SessionRuntime | null;
  models: Models;
  streamFn: ReturnType<typeof createOfficeStreamFn>;
  getWorkbookContext: () => Promise<WorkbookContext>;
  allowFaultInjection: boolean;
}): Promise<DurableTaskpane> {
  let runtime: Awaited<ReturnType<typeof openDurableRuntime>> | undefined;
  let lifecycle: ModelRecoveryLifecycle | undefined;
  let failure: string | undefined;
  const workbook = await options.getWorkbookContext();
  const active = options.getActiveRuntime();
  try {
    if (!workbook.workbookId) throw new Error("Save this workbook before opening the Durable prototype");
    if (!active?.buildDurableCapabilities) throw new Error("Taskpane capabilities unavailable");
    runtime = await openDurableRuntime({
      storage: await openIndexedDbStorage(`pi-for-excel-durable:${workbook.workbookId}`),
      models: routeModels(options.models, options.streamFn, () => lifecycle),
      onOpened: (harness) => {
        const recovery = createModelRecoveryLifecycle(() => harness.close(context));
        lifecycle = recovery;
        void harness.closed.then(() => recovery.dispose());
      },
      workbookId: workbook.workbookId,
      model: { provider: active.agent.state.model.provider, modelId: active.agent.state.model.id },
      buildCapabilities: async (hooks) => {
        const capabilities = await active.buildDurableCapabilities?.(hooks);
        if (!capabilities) throw new Error("Taskpane capabilities unavailable");
        return { ...capabilities, tools: capabilities.tools.map((tool) => ({ ...tool,
          execute: async (...args) => {
            if (!isExperimentalFeatureEnabled("durable_runtime")
              || (await options.getWorkbookContext()).workbookId !== workbook.workbookId) {
              throw new Error("Durable workbook or flag changed; reload required");
            }
            return tool.execute(...args);
          },
        })) };
      },
      allowFaultInjection: options.allowFaultInjection,
    });
  } catch (error) { failure = error instanceof Error ? error.message : "Durable could not open"; }
  return {
    execute: async (command, payloadJson) => {
      if (!isExperimentalFeatureEnabled("durable_runtime")) throw new Error("Durable flag is off; reload to stop the prototype");
      const payload: unknown = JSON.parse(payloadJson);
      if (!Value.Check(payloadSchema, payload)) throw new Error("Invalid Durable command payload");
      const input = payload;
      if (!runtime) throw new Error(failure ?? "Durable unavailable");
      if ((await options.getWorkbookContext()).workbookId !== workbook.workbookId) {
        throw new Error("Workbook identity changed; reload before using Durable");
      }
      switch (command) {
        case "durableStatus": return runtime.status();
        case "durableFault":
          if (input.point === undefined) throw new Error("durableFault requires point");
          runtime.mutations.configureFault(input.point);
          return runtime.mutations.status();
        case "durableResumeFault": runtime.mutations.resume(); return { resumed: true };
        case "durableSubmit": {
          if (!input.text || !input.requestId) throw new Error("durableSubmit requires text and requestId");
          if (active?.agent.state.isStreaming || active?.actionQueue.isBusy()) throw new Error("Normal chat is busy");
          await runtime.refresh();
          const model = active?.agent.state.model;
          if (model) await runtime.conversation.configure({ model: { provider: model.provider, modelId: model.id },
            thinkingLevel: active?.agent.state.thinkingLevel ?? "off" }, context);
          const submission = await runtime.conversation.submit({ type: "input", content: input.text, requestId: input.requestId }, context);
          return { submissionId: submission.id };
        }
        case "durableImport": {
          if (!input.sessionId) throw new Error("durableImport requires sessionId");
          const saved = await getAppStorage().sessions.get(input.sessionId);
          if (!saved) throw new Error("Legacy session not found");
          const imported = await importLegacySession(runtime.harness, saved, context);
          return { conversationId: imported.id };
        }
        case "durableReconcile": {
          if (!input.toolCallId) throw new Error("durableReconcile requires toolCallId");
          const classification = await runtime.mutations.reconcile(input.toolCallId, (record) => excelRun(async (excel) => {
            const sheet = excel.workbook.worksheets.getItem(record.sheetId);
            const range = sheet.getRange(record.range);
            range.load("values,formulas");
            await excel.sync();
            const values: unknown = range.values;
            const formulas: unknown = range.formulas;
            if (!Value.Check(gridSchema, values) || !Value.Check(gridSchema, formulas)) throw new Error("Invalid workbook read");
            return { values, formulas };
          }));
          return { classification, wroteCells: false };
        }
        default: throw new Error("Unknown Durable command");
      }
    },
  };
}
