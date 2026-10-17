import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream, normalizeContext, type AssistantMessage } from "@earendil-works/pi-ai";
import type { Models } from "@earendil-works/pi-ai/models";
import { createRegistry, defineDoc, defineExtension, GenerationTask, hook, Harness, section, ToolTask,
  type Conversation, type Storage, type ToolRegistration, type ToolExecutionApi,
  type TaskId, type ConversationId, type ToolTaskResult } from "@earendil-works/pi-durable";
import type { createOfficeStreamFn } from "../auth/stream-proxy.js";
import type { AnyCoreTool, CoreToolName } from "../tools/registry.js";
import type { WorkbookCoordinator } from "../workbook/coordinator.js";
import { createMutationController, parseMutationRecord, type MutationHost } from "./mutations.js";
import { protectModelStream, type ModelRecoveryLifecycle } from "./model-recovery.js";
import { compactionForModel } from "./compaction-policy.js";
import { readConversationHistory } from "./conversation-view.js";
import { createConvertToLlm } from "../messages/convert-to-llm.js";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { WorkbookMutationGate } from "../tools/mutation/gate.js";
import { getToolExecutionMode } from "../tools/execution-policy.js";

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

export const MutationsDoc = defineDoc<{ records: Record<string, JsonValue> }>({
  kind: "excel.workbook-mutations", version: 1, scope: "session",
  initial: () => ({ records: {} }),
});

export const MutationBindingsDoc = defineDoc<{ bindings: Record<string, {
  taskId: TaskId; conversationId: ConversationId; rawCallId: string;
}> }>({ kind: "excel.mutation-bindings", version: 1, scope: "session", initial: () => ({ bindings: {} }) });

export async function openDurableRuntime(options: {
  storage: Storage;
  models: Models;
  workbookId: string;
  mutationHost: MutationHost;
  coordinator: WorkbookCoordinator;
  autoCompaction: boolean;
  allowFaultInjection?: boolean;
  /** Runs before root setup or scheduling, so browser lifecycle owns this exact Harness. */
  onOpened?: (harness: Harness) => void;
}): Promise<{
  harness: Harness;
  mutations: ReturnType<typeof createMutationController>;
  install(conversation: Conversation, capabilities: DurableCapabilities,
    transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>): Promise<void>;
  acknowledgeCommittedResults(): Promise<void>;
  status(conversation: Conversation): Promise<JsonValue>;
}> {
  const registry = createRegistry();
  const harness = await Harness.open(options.storage, { models: options.models, registry,
    settings: { toolExecution: "sequential", compaction: { enabled: options.autoCompaction }, compactionForModel },
  }, context);
  try {
    options.onOpened?.(harness);
    const mutations = createMutationController({
      workbookId: options.workbookId,
      host: options.mutationHost,
      coordinator: options.coordinator,
      allowFaultInjection: options.allowFaultInjection ?? false,
      store: {
        load: async (id) => {
          const value = (await harness.snapshot(MutationsDoc, context))?.records[id];
          return value === undefined ? null : parseMutationRecord(value);
        },
        list: async () => Object.values((await harness.snapshot(MutationsDoc, context))?.records ?? {}).map(parseMutationRecord),
        save: (record) => harness.commit(async (tx) => {
          const doc = await tx.doc(MutationsDoc);
          const existing = doc.records[record.mutationId];
          if (existing !== undefined && parseMutationRecord(existing).phase === "resolved") return;
          doc.records[record.mutationId] = record;
        }, context),
      },
    });
    // Read the exact task receipt, never search by provider-owned raw call ID.
    const acknowledgeCommittedResults = async (): Promise<void> => {
      const records = await mutations.pending();
      const bindings = (await harness.snapshot(MutationBindingsDoc, context))?.bindings ?? {};
      for (const record of records) {
        if (record.phase !== "returned") continue;
        const binding = bindings[record.mutationId];
        if (!binding) continue;
        // Only adaptTools writes these bindings, from the built-in ToolTask API.
        const task = await harness.getTask(binding.taskId as TaskId<ToolTaskResult>, context);
        if (task?.state.status !== "terminal" || task.state.outcome.status !== "completed") continue;
        const result = task.state.outcome.result;
        if (result.kind === "model") {
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
    const install = async (conversation: Conversation, capabilities: DurableCapabilities,
      transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>): Promise<void> => {
      const name = `excel:${conversation.id}`;
      const extension = defineExtension({ name, tools: adaptTools(capabilities.tools, async (tool, api, args, signal) => {
        if (getToolExecutionMode(tool.name, args) === "read") return tool.execute(api.callId, args, signal);
        const previous = mutationAdmission;
        let release = (): void => {};
        mutationAdmission = new Promise<void>((resolve) => { release = resolve; });
        await previous;
        try {
          if (signal?.aborted) throw new Error("Aborted");
          if (!capabilities.bindMutationGate) throw new Error("Mutation dispatch is not bound to its Durable task");
          const mutationId = `task:${api.taskId}`;
          await harness.commit(async (tx) => {
            (await tx.doc(MutationBindingsDoc)).bindings[mutationId] = {
              taskId: api.taskId, conversationId: api.conversationId, rawCallId: api.callId,
            };
          }, context);
          const bound = capabilities.bindMutationGate(mutations, mutationId).find((candidate) => candidate.name === tool.name);
          if (!bound) throw new Error("Mutation tool registration disappeared");
          return await bound.execute(api.callId, args, signal);
        } finally {
          void harness.waitForTask(api.taskId, context).then(acknowledgeCommittedResults)
            .catch((error: unknown) => { console.warn("[durable] Mutation acknowledgement failed", error); })
            .finally(release);
        }
      }),
        sections: [section("excel", () => capabilities.systemPrompt, { tag: false })],
        hooks: [hook(ToolTask, { beforeTool: async () => { await acknowledgeCommittedResults(); } }),
          hook(GenerationTask, {
            beforeRequest: async (request, _api, ctx) => {
              const agent = await conversation.agent(ctx);
              const model = agent.model && options.models.getModel(agent.model.provider, agent.model.modelId);
              const convert = createConvertToLlm({ getContextWindow: () => model?.contextWindow });
              return { messages: convert(await transformContext([...request.messages], ctx.abortSignal)) };
            },
            afterTools: acknowledgeCommittedResults,
          })],
      });
      registry.install(extension);
      const selected = (await conversation.agent(context)).extensions;
      if (selected.length !== 1 || selected[0]?.name !== name) await conversation.configure({ extensions: [extension] }, context);
    };
    return { harness, mutations, install, acknowledgeCommittedResults,
      status: async (conversation) => {
        const status: unknown = JSON.parse(JSON.stringify({
          conversationId: conversation.id,
          fault: mutations.status(),
          records: (await harness.snapshot(MutationsDoc, context))?.records ?? {},
          mutationBindings: (await harness.snapshot(MutationBindingsDoc, context))?.bindings ?? {},
          entries: await readConversationHistory(conversation),
        }));
        if (!isJsonValue(status)) throw new Error("Invalid durable status");
        return status;
      },
    };
  } catch (error) { await harness.close(context); throw error; }
}
