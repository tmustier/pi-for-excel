import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream, normalizeContext, type AssistantMessage } from "@earendil-works/pi-ai";
import type { Models } from "@earendil-works/pi-ai/models";
import { Value } from "typebox/value";
import { createRegistry, defineDoc, defineDocFamily, defineEntry, defineTask, defineExtension, GenerationTask, CompactionTask, hook, Harness, section, ToolTask,
  type Conversation, type Storage, type ToolRegistration, type ToolExecutionApi,
  type TaskId, type ConversationId, type ToolTaskResult, type Cursor } from "@earendil-works/pi-durable";
import type { createOfficeStreamFn } from "../auth/stream-proxy.js";
import type { AnyCoreTool, CoreToolName } from "../tools/registry.js";
import type { WorkbookCoordinator } from "../workbook/coordinator.js";
import { createMutationController, MutationExecutionError, parseMutationRecord, type MutationHost, type MutationRecord, type MutationReconciliation } from "./mutations.js";
import { protectModelStream, type ModelRecoveryLifecycle } from "./model-recovery.js";
import { compactionForModel } from "./compaction-policy.js";
import { readConversationHistory } from "./conversation-view.js";
import { createConvertToLlm } from "../messages/convert-to-llm.js";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { WorkbookMutationGate } from "../tools/mutation/gate.js";
import { getToolExecutionMode } from "../tools/execution-policy.js";
import { compactionMemoryNudge } from "./memory-nudge.js";
import { normalizeToolResultImages } from "../messages/image-input-limits.js";

export interface DurableCapabilities {
  tools: AnyCoreTool[];
  systemPrompt: string;
  /** Bind invocation identity without changing the provider's call ID passed to tools. */
  bindMutationGate?: (gate: WorkbookMutationGate, mutationId: string) => AnyCoreTool[];
}

// Workbook-read-only tools only. Mixed-action tools and external/local side
// effects stay unsafe, even when workbook coordination classifies them as reads.
const replaySafe: ReadonlySet<string> = new Set<CoreToolName>([
  "get_workbook_overview", "read_range", "search_workbook", "trace_dependencies", "explain_formula",
]);

export function adaptTools(tools: readonly AnyCoreTool[], bind?: (tool: AnyCoreTool, api: ToolExecutionApi,
  args: Parameters<AnyCoreTool["execute"]>[1], signal: AbortSignal | undefined) => Promise<Awaited<ReturnType<AnyCoreTool["execute"]>>>): ToolRegistration[] {
  return tools.map((tool) => ({
    name: tool.name, description: tool.description, parameters: tool.parameters,
    ...(replaySafe.has(tool.name) ? { replay: "safe" as const } : {}),
    execute: async (args, api, ctx) => {
      const result = bind ? await bind(tool, api, args, ctx.abortSignal) : await tool.execute(api.callId, args, ctx.abortSignal);
      const details = result.details;
      if (details !== undefined && !isJsonValue(details)) throw new Error("Tool details must be JSON serializable");
      return { output: result.content, ...(details !== undefined ? { details } : {}) };
    },
  }));
}

/** Keep credentials in closures, never in Durable's persisted stream settings. */
export function routeModels(
  models: Models,
  streamFn: ReturnType<typeof createOfficeStreamFn>,
  getRecoveryLifecycle?: () => ModelRecoveryLifecycle | undefined,
): Models {
  const streamSimple: Models["streamSimple"] = (model, transcript, options) => {
    const output = createAssistantMessageEventStream();
    const failure = (): AssistantMessage => ({
      role: "assistant", content: [], api: model.api, provider: model.provider, model: model.id,
      stopReason: "error", errorMessage: "Browser model transport failed. Check provider connection and proxy.",
      timestamp: Date.now(), usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
    });
    void (async () => {
      try {
        const auth = await models.getAuth(model);
        const source = await streamFn(model, normalizeContext(transcript), {
          ...options, ...(auth ? { apiKey: auth.auth.apiKey } : {}),
        });
        for await (const event of source) output.push(event);
        output.end(await source.result());
      } catch {
        const error = failure();
        output.push({ type: "error", reason: "error", error });
      }
    })();
    const lifecycle = getRecoveryLifecycle?.();
    return lifecycle ? protectModelStream(output, lifecycle, failure) : output;
  };
  return {
    getProviders: models.getProviders.bind(models), getProvider: models.getProvider.bind(models),
    getModels: models.getModels.bind(models), getModel: models.getModel.bind(models),
    getModelsOfType: models.getModelsOfType.bind(models), getModelOfType: models.getModelOfType.bind(models),
    getAllModels: models.getAllModels.bind(models), refresh: models.refresh.bind(models),
    checkAuth: models.checkAuth.bind(models), getAvailable: models.getAvailable.bind(models),
    getAvailableOfType: models.getAvailableOfType.bind(models), getAllAvailable: models.getAllAvailable.bind(models),
    getAuth: models.getAuth.bind(models), login: models.login.bind(models), logout: models.logout.bind(models),
    stream: models.stream.bind(models), complete: models.complete.bind(models), streamSimple,
    completeSimple: async (model, transcript, options) => streamSimple(model, transcript, options).result(),
    streamDeferred: models.streamDeferred.bind(models), fetchDeferred: models.fetchDeferred.bind(models),
    cancelDeferred: models.cancelDeferred.bind(models), generateImages: models.generateImages.bind(models),
    classify: models.classify.bind(models),
  };
}

const OldMutationsDoc = defineDoc<{ records: Record<string, JsonValue> }>({
  kind: "excel.workbook-mutations", version: 1, scope: "session",
  initial: () => ({ records: {} }),
});

const OldMutationBindingsDoc = defineDoc<{ bindings: Record<string, {
  taskId: TaskId; conversationId: ConversationId; rawCallId: string;
}> }>({ kind: "excel.mutation-bindings", version: 1, scope: "session", initial: () => ({ bindings: {} }) });

const PrototypeMutationsDoc = defineDoc<{ records: Record<string, JsonValue> }>({
  kind: "excel.mutations", version: 1, scope: "conversation", history: "latest", fork: "initial", initial: () => ({ records: {} }),
});
export const MutationDoc = defineDocFamily<{
  record: JsonValue | null;
  binding: { taskId: TaskId; conversationId: ConversationId; rawCallId: string } | null;
  original: JsonValue | null;
  resolution: MutationReconciliation | null;
}, null>({ kind: "excel.mutation", version: 1, scope: "session", family: true,
  initial: () => ({ record: null, binding: null, original: null, resolution: null }) });
export const PendingMutationsDoc = defineDoc<{ ids: Record<string, true>; migrated: boolean }>({
  kind: "excel.pending-mutations", version: 1, scope: "session", initial: () => ({ ids: {}, migrated: false }),
});
export interface MutationReview {
  mutationId: string;
  toolName: string;
  legacy: boolean;
  evidence: JsonValue;
}

export const DirectToolResultEntry = defineEntry<{ name: string; result: JsonValue; status: string }>("excel.direct-tool-result");

export async function openDurableRuntime(options: {
  storage: Storage;
  models: Models;
  workbookId: string;
  mutationHost: MutationHost;
  coordinator: WorkbookCoordinator;
  autoCompaction: boolean;
  autoResizeImages?: boolean;
  allowFaultInjection?: boolean;
  /** Runs before root setup or scheduling, so browser lifecycle owns this exact Harness. */
  onOpened?: (harness: Harness) => void;
}): Promise<{
  harness: Harness;
  mutations: ReturnType<typeof createMutationController>;
  reviews(): Promise<MutationReview[]>;
  install(conversation: Conversation, capabilities: DurableCapabilities,
    transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>): Promise<void>;
  acknowledgeCommittedResults(): Promise<void>;
  executeTool(conversation: Conversation, name: string, args: JsonValue): Promise<JsonValue>;
  status(conversation: Conversation): Promise<JsonValue>;
}> {
  const registry = createRegistry();
  const capabilitiesByConversation = new Map<ConversationId, DurableCapabilities>();
  let invoke: (tool: AnyCoreTool, identity: { taskId: TaskId; conversationId: ConversationId; rawCallId: string }, args: Parameters<AnyCoreTool["execute"]>[1], signal?: AbortSignal, dispatch?: () => Promise<void>) => ReturnType<AnyCoreTool["execute"]>;
  const directTask = defineTask<{ name: string; args: JsonValue }, { phase: "ready" } | { phase: "dispatched" }, JsonValue>({
    name: "excel.direct-tool", version: 1, initial: () => ({ phase: "ready" }),
    phases: {
      ready: async (task, runtime, ctx) => {
        try {
          const tool = capabilitiesByConversation.get(runtime.conversationId)?.tools.find((candidate) => candidate.name === task.input.name);
          if (!tool) throw new Error("Direct tool unavailable for this conversation");
          if (!Value.Check(tool.parameters, task.input.args)) throw new Error("Invalid direct tool arguments");
          const result = await invoke(tool, { taskId: runtime.taskId, conversationId: runtime.conversationId, rawCallId: `direct:${runtime.taskId}` }, task.input.args,
            runtime.signal, () => runtime.commit(async (tx) => {
              if (getToolExecutionMode(tool.name, task.input.args) !== "read") {
                (await tx.doc(MutationDoc, `task:${runtime.taskId}`, null)).binding = {
                  taskId: runtime.taskId, conversationId: runtime.conversationId, rawCallId: `direct:${runtime.taskId}`,
                };
              }
              return { status: "running", checkpoint: { phase: "dispatched" } };
            }, ctx));
          if (!isJsonValue(result)) throw new Error("Tool result must be JSON serializable");
          await runtime.commit(async (tx) => {
            await tx.appendEntry(DirectToolResultEntry, runtime.conversationId, { data: { name: task.input.name, result, status: "completed" } });
            return { status: "terminal", outcome: { status: "completed", result } };
          }, ctx);
        } catch (error) {
          const message = error instanceof MutationExecutionError ? error.message
            : "Direct tool did not complete. Workbook changes may already be applied; inspect pending changes before retrying.";
          if (runtime.signal.aborted) throw error;
          await runtime.commit(async (tx) => {
            await tx.appendEntry(DirectToolResultEntry, runtime.conversationId, { data: { name: task.input.name, result: message, status: "failed" } });
            return { status: "terminal", outcome: { status: "failed", error: { message } } };
          }, ctx);
        }
      },
      dispatched: async (task, runtime, ctx) => {
        await runtime.commit(async (tx) => {
          await tx.appendEntry(DirectToolResultEntry, runtime.conversationId, { data: { name: task.input.name, result: null, status: "interrupted" } });
          return { status: "terminal", outcome: { status: "failed", error: { message: "Direct tool interrupted; never replay" } } };
        }, ctx);
      },
    },
    abort: async (task, runtime, ctx) => {
      await runtime.commit(async (tx) => {
        await tx.appendEntry(DirectToolResultEntry, runtime.conversationId, { data: { name: task.input.name, result: null, status: "aborted" } });
        return { status: "terminal", outcome: { status: "aborted" } };
      }, ctx);
    },
  });
  registry.install(defineExtension({ name: "excel.direct-tools", tasks: [directTask] }));
  const harness = await Harness.open(options.storage, { models: options.models, registry,
    settings: { toolExecution: "sequential", compaction: { enabled: options.autoCompaction }, compactionForModel },
  }, context);
  try {
    options.onOpened?.(harness);
    const pendingIds = async (): Promise<string[]> => Object.keys((await harness.snapshot(PendingMutationsDoc, context))?.ids ?? {});
    if (!(await harness.snapshot(PendingMutationsDoc, context))?.migrated) {
      const copy = async (id: string, original: JsonValue, binding: { taskId: TaskId; conversationId: ConversationId; rawCallId: string } | null): Promise<void> => {
        let record: MutationRecord | null = null;
        try { record = parseMutationRecord(original); } catch { /* Original remains available for explicit reviewed resolution. */ }
        await harness.commit(async (tx) => {
          const target = await tx.doc(MutationDoc, id, null);
          if (target.original !== null || target.record !== null) return;
          target.original = original;
          target.record = record;
          target.binding = binding;
          if (!record || (record.phase !== "resolved" && (record.phase !== "resultCommitted" || record.requiresReview))) {
            (await tx.doc(PendingMutationsDoc)).ids[id] = true;
          }
        }, context);
      };
      const old = await harness.snapshot(OldMutationsDoc, context);
      const bindings = await harness.snapshot(OldMutationBindingsDoc, context);
      for (const [id, raw] of Object.entries(old?.records ?? {})) await copy(id, raw, bindings?.bindings[id] ?? null);
      let cursor: Cursor | undefined;
      do {
        // agent-default (not a user rule): 2026-10-18. Scan 200 conversations per migration page.
        const page = await harness.commit((tx) => tx.scanConversations({}, 200, cursor), context);
        for (const row of page.items) {
          const prototype = await harness.snapshot(PrototypeMutationsDoc, row.id, context);
          for (const [id, raw] of Object.entries(prototype?.records ?? {})) await copy(`legacy:${row.id}:${id}`, raw, null);
        }
        cursor = page.next;
      } while (cursor !== undefined);
      await harness.commit(async (tx) => { (await tx.doc(PendingMutationsDoc)).migrated = true; }, context);
    }
    const controller = createMutationController({
      workbookId: options.workbookId,
      host: options.mutationHost,
      coordinator: options.coordinator,
      allowFaultInjection: options.allowFaultInjection ?? false,
      store: {
        load: async (id) => {
          const value = (await harness.snapshot(MutationDoc, id, context))?.record;
          return value == null ? null : parseMutationRecord(value);
        },
        list: async () => {
          const records: MutationRecord[] = [];
          for (const id of await pendingIds()) {
            const doc = await harness.snapshot(MutationDoc, id, context);
            if (!doc) throw new Error("Pending mutation record is missing");
            if (doc.record !== null) records.push(parseMutationRecord(doc.record));
          }
          return records;
        },
        save: (record) => harness.commit(async (tx) => {
          const doc = await tx.doc(MutationDoc, record.mutationId, null);
          if (doc.record !== null && parseMutationRecord(doc.record).phase === "resolved") return;
          doc.record = record;
          const pending = await tx.doc(PendingMutationsDoc);
          if (record.phase === "resolved" || (record.phase === "resultCommitted" && !record.requiresReview)) delete pending.ids[record.mutationId];
          else pending.ids[record.mutationId] = true;
        }, context),
      },
    });
    const reviews = async (): Promise<MutationReview[]> => {
      const result: MutationReview[] = [];
      for (const id of await pendingIds()) {
        const doc = await harness.snapshot(MutationDoc, id, context);
        if (!doc) throw new Error("Pending mutation record is missing");
        result.push({ mutationId: id, toolName: doc.record ? parseMutationRecord(doc.record).toolName : "Legacy workbook change",
          legacy: doc.record === null, evidence: doc.record ?? doc.original });
      }
      return result;
    };
    const reconcileLegacy = async (id: string): Promise<MutationReconciliation> => {
      const doc = await harness.snapshot(MutationDoc, id, context);
      if (!doc || doc.original === null || doc.record !== null || doc.resolution) throw new Error("Legacy review unavailable");
      const plan = await options.mutationHost.prepare({ mutationId: id, toolCallId: id, toolName: "legacy_mutation_review",
        workbookId: options.workbookId, sessionId: "legacy-review", params: {} });
      const observation = plan.before;
      const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify({ original: doc.original, observation })));
      return { classification: "conflict", reason: "Legacy evidence cannot prove this operation's outcome. Inspect the workbook and acknowledge uncertainty; this never authorizes replay.",
        observation, observationToken: Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("") };
    };
    const mutations: ReturnType<typeof createMutationController> = { ...controller,
      execute: async (call, run) => {
        if ((await reviews()).some((review) => review.legacy)) throw new Error("Review legacy workbook changes before writing; do not retry uncertain operations");
        return controller.execute(call, run);
      },
      reconcile: async (id) => (await harness.snapshot(MutationDoc, id, context))?.record !== null
        ? controller.reconcile(id) : reconcileLegacy(id),
      resolve: async (id, token) => {
        if ((await harness.snapshot(MutationDoc, id, context))?.record !== null) return controller.resolve(id, token);
        return (await options.coordinator.runWrite({ workbookId: options.workbookId, sessionId: "legacy-review", opId: id, journalOnly: true }, async () => {
          const resolution = await reconcileLegacy(id);
          if (resolution.observationToken !== token) throw new Error("Workbook state changed since review; inspect again before acknowledging");
          await harness.commit(async (tx) => {
            (await tx.doc(MutationDoc, id, null)).resolution = resolution;
            delete (await tx.doc(PendingMutationsDoc)).ids[id];
          }, context);
          return resolution;
        })).result;
      },
    };
    // Read the exact task receipt, never search by provider-owned raw call ID.
    const acknowledgeCommittedResults = async (): Promise<void> => {
      const records = await mutations.pending();
      for (const record of records) {
        if (record.phase !== "returned") continue;
        const binding = (await harness.snapshot(MutationDoc, record.mutationId, context))?.binding;
        if (!binding) continue;
        // Both application direct tasks and built-in ToolTasks persist exact bindings.
        const task = await harness.getTask(binding.taskId as TaskId<ToolTaskResult>, context);
        if (task?.state.status !== "terminal" || task.state.outcome.status !== "completed") continue;
        const result = task.state.outcome.result;
        if (task.kind !== "excel.direct-tool" && result.kind === "model") {
          const entry = await harness.commit((tx) => tx.entry(result.entryId), context);
          if (!entry?.model?.some((message) => message.role === "toolResult" && !message.isError)) continue;
        }
        await mutations.resultCommitted(record.mutationId);
      }
    };
    // This admission queue is outside the workbook lock. It holds the next
    // mutation until the previous ToolTask commits and its acknowledgement has
    // acquired/released the coordinator; reads and model requests stay parallel.
    let mutationAdmission = Promise.resolve();
    invoke = async (tool, identity, args, signal, dispatch) => {
      const capabilities = capabilitiesByConversation.get(identity.conversationId);
      const run = async (): ReturnType<AnyCoreTool["execute"]> => {
        if (signal?.aborted) throw new Error("Aborted");
        if (getToolExecutionMode(tool.name, args) === "read") {
          await dispatch?.();
          return tool.execute(identity.rawCallId, args, signal);
        }
        if (!capabilities?.bindMutationGate) throw new Error("Mutation dispatch is not bound to its Durable task");
        const mutationId = `task:${identity.taskId}`;
        if (!dispatch) await harness.commit(async (tx) => {
          (await tx.doc(MutationDoc, mutationId, null)).binding = identity;
        }, context);
        const bound = capabilities.bindMutationGate(mutations, mutationId).find((candidate) => candidate.name === tool.name);
        if (!bound) throw new Error("Mutation tool registration disappeared");
        await dispatch?.();
        return bound.execute(identity.rawCallId, args, signal);
      };
      if (getToolExecutionMode(tool.name, args) === "read") return run();
      const previous = mutationAdmission;
      let release = (): void => {};
      mutationAdmission = new Promise<void>((resolve) => { release = resolve; });
      await previous;
      try { return await run(); }
      finally {
        void harness.waitForTask(identity.taskId, context).then(acknowledgeCommittedResults)
          .catch(() => { console.warn("[durable] Mutation acknowledgement unavailable; workbook changes may already be applied. Review pending changes before further writes."); }).finally(release);
      }
    };
    const install = async (conversation: Conversation, capabilities: DurableCapabilities,
      transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>): Promise<void> => {
      const name = `excel:${conversation.id}`;
      capabilitiesByConversation.set(conversation.id, capabilities);
      const extension = defineExtension({ name, tools: adaptTools(capabilities.tools, async (tool, api, args, signal) => {
        return invoke(tool, { taskId: api.taskId, conversationId: api.conversationId, rawCallId: api.callId }, args, signal);
      }),
        sections: [section("excel", () => capabilities.systemPrompt, { tag: false })],
        hooks: [hook(CompactionTask, { beforeSummary: compactionMemoryNudge }),
          hook(ToolTask, { beforeTool: async () => { await acknowledgeCommittedResults(); } }),
          hook(GenerationTask, {
            beforeRequest: async (request, _api, ctx) => {
              const agent = await conversation.agent(ctx);
              const model = agent.model && options.models.getModel(agent.model.provider, agent.model.modelId);
              const convert = createConvertToLlm({ getContextWindow: () => model?.contextWindow });
              const converted = convert(await transformContext([...request.messages], ctx.abortSignal));
              const messages: typeof converted = [];
              for (const message of converted) {
                messages.push(message.role === "toolResult" && model
                  ? { ...message, content: await normalizeToolResultImages(message.content, model,
                    { autoResizeImages: options.autoResizeImages ?? true, ...(ctx.abortSignal ? { signal: ctx.abortSignal } : {}) }) }
                  : message);
              }
              const unresolved = await reviews();
              if (unresolved.length) messages.push({ role: "user", timestamp: 0, content: `[Workbook mutation uncertainty]\n${unresolved.length} workbook operation(s) require explicit user review. Do not claim these writes succeeded, and do not retry writes under a new call ID. Report read-only observations separately from proof of the intended operation. The pending-review UI is authoritative. Ask the user to inspect and acknowledge the current state before further changes.` });
              return { messages };
            },
            afterTools: acknowledgeCommittedResults,
          })],
      });
      registry.install(extension);
      const selected = (await conversation.agent(context)).extensions;
      if (selected.length !== 1 || selected[0]?.name !== name) await conversation.configure({ extensions: [extension] }, context);
    };
    return { harness, mutations, reviews, install, acknowledgeCommittedResults,
      executeTool: async (conversation, name, args) => {
        const id = await harness.commit((tx) => tx.createTask(directTask, { name, args }, {
          conversationId: conversation.id, ownership: { kind: "conversation" }, abandonOnRestart: true,
        }), context);
        const task = await harness.waitForTask(id, context);
        await acknowledgeCommittedResults();
        if (task.state.status !== "terminal" || task.state.outcome.status !== "completed") {
          throw new Error(task.state.status === "terminal" && "error" in task.state.outcome
            ? task.state.outcome.error?.message ?? "Direct tool did not complete" : "Direct tool did not complete");
        }
        return task.state.outcome.result;
      },
      status: async (conversation) => {
        const status: unknown = JSON.parse(JSON.stringify({
          conversationId: conversation.id,
          fault: mutations.status(),
          records: Object.fromEntries((await reviews()).map((review) => [review.mutationId, review.evidence])),
          entries: await readConversationHistory(conversation),
        }));
        if (!isJsonValue(status)) throw new Error("Invalid durable status");
        return status;
      },
    };
  } catch (error) { await harness.close(context); throw error; }
}
