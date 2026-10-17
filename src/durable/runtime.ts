import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import { createAssistantMessageEventStream, normalizeContext, type AssistantMessage } from "@earendil-works/pi-ai";
import type { Models } from "@earendil-works/pi-ai/models";
import { createRegistry, defineDoc, defineExtension, Harness, section, ToolResultEntry, type Storage, type ToolRegistration } from "@earendil-works/pi-durable";
import type { createOfficeStreamFn } from "../auth/stream-proxy.js";
import type { AnyCoreTool, CoreToolName } from "../tools/registry.js";
import type { WriteCellsHooks } from "../tools/write-cells.js";
import { createMutationController, type MutationRecord } from "./mutations.js";

// Workbook-read-only tools only. Mixed-action tools and external/local side
// effects stay unsafe, even when workbook coordination classifies them as reads.
const replaySafe: ReadonlySet<string> = new Set<CoreToolName>([
  "get_workbook_overview", "read_range", "search_workbook", "trace_dependencies", "explain_formula",
]);

export function adaptTools(tools: readonly AnyCoreTool[]): ToolRegistration[] {
  return tools.map((tool) => ({
    name: tool.name, description: tool.description, parameters: tool.parameters,
    ...(replaySafe.has(tool.name) ? { replay: "safe" as const } : {}),
    execute: async (args, api, ctx) => {
      const result = await tool.execute(api.callId, args, ctx.abortSignal);
      const details = result.details;
      if (details !== undefined && !isJsonValue(details)) throw new Error("Tool details must be JSON serializable");
      return { output: result.content, ...(details !== undefined ? { details } : {}) };
    },
  }));
}

/** Keep credentials in closures, never in Durable's persisted stream settings. */
export function routeModels(models: Models, streamFn: ReturnType<typeof createOfficeStreamFn>): Models {
  const streamSimple: Models["streamSimple"] = (model, transcript, options) => {
    const output = createAssistantMessageEventStream();
    void (async () => {
      try {
        const auth = await models.getAuth(model);
        const source = await streamFn(model, normalizeContext(transcript), {
          ...options, ...(auth ? { apiKey: auth.auth.apiKey } : {}),
        });
        for await (const event of source) output.push(event);
        output.end(await source.result());
      } catch {
        const error: AssistantMessage = {
          role: "assistant", content: [], api: model.api, provider: model.provider, model: model.id,
          stopReason: "error", errorMessage: "Browser model transport failed. Check provider connection and proxy.",
          timestamp: Date.now(), usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
            cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
        };
        output.push({ type: "error", reason: "error", error });
      }
    })();
    return output;
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

export const MutationsDoc = defineDoc<{ records: Record<string, MutationRecord> }>({
  kind: "excel.mutations", version: 1, scope: "conversation", history: "latest", fork: "initial",
  initial: () => ({ records: {} }),
});

export async function openDurableRuntime(options: {
  storage: Storage;
  models: Models;
  workbookId: string;
  model: { provider: string; modelId: string };
  buildCapabilities: (hooks: WriteCellsHooks) => Promise<{ tools: AnyCoreTool[]; systemPrompt: string }>;
  allowFaultInjection?: boolean;
}): Promise<{
  harness: Harness;
  conversation: Awaited<ReturnType<Harness["root"]>>;
  mutations: ReturnType<typeof createMutationController>;
  refresh(): Promise<void>;
  status(): Promise<JsonValue>;
}> {
  const registry = createRegistry();
  const harness = await Harness.open(options.storage, { models: options.models, registry,
    settings: { toolExecution: "sequential", compaction: { enabled: false } },
  }, context);
  try {
    const conversation = await harness.root(context, { agent: { model: options.model } });
    const mutations = createMutationController({
      workbookId: options.workbookId,
      allowFaultInjection: options.allowFaultInjection ?? false,
      store: {
        load: async (id) => (await harness.snapshot(MutationsDoc, conversation.id, context))?.records[id] ?? null,
        save: (record) => conversation.commit(async (tx) => {
          (await tx.doc(MutationsDoc, conversation.id)).records[record.toolCallId] = record;
        }, context),
      },
    });
    const refresh = async (): Promise<void> => {
      const capabilities = await options.buildCapabilities(mutations.hooks);
      const tools = adaptTools(capabilities.tools).map((tool) => ({ ...tool,
        execute: async (...args: Parameters<NonNullable<ToolRegistration["execute"]>>) => {
          if (tool.replay !== "safe") {
            const records = (await harness.snapshot(MutationsDoc, conversation.id, context))?.records ?? {};
            const transcript = await conversation.context(context);
            const completed = new Set(transcript.messages.flatMap((message) =>
              message.role === "toolResult" && !message.isError ? [message.toolCallId] : []));
            const interrupted = transcript.entries.some((entry) => ToolResultEntry.is(entry)
              && entry.data?.diagnostics.some((diagnostic) => diagnostic.code === "interrupted")
              && entry.model?.some((message) => message.role === "toolResult" && !replaySafe.has(message.toolName)));
            if (interrupted || Object.keys(records).some((id) => !completed.has(id))) {
              throw new Error("An earlier write is uncertain. Reconcile it with the verifier; automatic mutation retries are blocked.");
            }
          }
          return tool.execute(...args);
        },
      }));
      registry.install(defineExtension({ name: "excel", tools,
        sections: [section("excel", () => capabilities.systemPrompt, { tag: false })],
      }));
    };
    await refresh();
    harness.resume();
    return { harness, conversation, mutations, refresh,
      status: async () => {
        const status: unknown = JSON.parse(JSON.stringify({
          conversationId: conversation.id,
          fault: mutations.status(),
          records: (await harness.snapshot(MutationsDoc, conversation.id, context))?.records ?? {},
          entries: (await conversation.entries({}, 100, undefined, context)).items,
        }));
        if (!isJsonValue(status)) throw new Error("Invalid durable status");
        return status;
      },
    };
  } catch (error) { await harness.close(context); throw error; }
}
