import type {
  Api,
  Model,
  ModelsStore,
  ModelsStoreEntry,
} from "@earendil-works/pi-ai";
import {
  ModelCostSchema,
  ModelInputLimitsSchema,
  ModelInputModalitySchema,
  ModelPromptCacheSchema,
  ThinkingLevelMapSchema,
} from "@earendil-works/pi-ai/providers/model-schema";
import { ProviderCompatSchema } from "@earendil-works/pi-ai/providers/compat-schema";
import { Type } from "typebox";
import { Value } from "typebox/value";

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

export const MAX_MODEL_ID_LENGTH = 256;

/** Chat model fields shared by persisted catalogues and pi.dev catalogue responses. */
const catalogModelSchema = Type.Object({
  type: Type.Optional(Type.Literal("chat")),
  id: Type.String({ minLength: 1, maxLength: MAX_MODEL_ID_LENGTH }),
  name: Type.String(),
  api: Type.String({ minLength: 1 }),
  provider: Type.String({ minLength: 1 }),
  baseUrl: Type.String(),
  reasoning: Type.Boolean(),
  input: Type.Array(ModelInputModalitySchema),
  cost: ModelCostSchema,
  contextWindow: Type.Number(),
  maxTokens: Type.Number(),
  headers: Type.Optional(Type.Record(Type.String(), Type.String())),
  thinkingLevelMap: Type.Optional(ThinkingLevelMapSchema),
  promptCache: Type.Optional(ModelPromptCacheSchema),
  inputLimits: Type.Optional(ModelInputLimitsSchema),
  compat: Type.Optional(ProviderCompatSchema),
});

const catalogEntrySchema = Type.Object({
  models: Type.Array(Type.Unknown()),
  checkedAt: Type.Optional(Type.Number()),
  lastModified: Type.Optional(Type.Number()),
  etag: Type.Optional(Type.String()),
});

/** Decode one chat model, dropping top-level fields the runtime does not use. Null if invalid. */
export function decodeCatalogModel(value: unknown): Model<Api> | null {
  if (!Value.Check(catalogModelSchema, value)) return null;
  const { headers, thinkingLevelMap, promptCache, inputLimits, compat } = value;
  return {
    id: value.id,
    name: value.name,
    api: value.api,
    provider: value.provider,
    baseUrl: value.baseUrl,
    reasoning: value.reasoning,
    input: value.input,
    cost: value.cost,
    contextWindow: value.contextWindow,
    maxTokens: value.maxTokens,
    ...(headers !== undefined ? { headers } : {}),
    ...(thinkingLevelMap !== undefined ? { thinkingLevelMap } : {}),
    ...(promptCache !== undefined ? { promptCache } : {}),
    ...(inputLimits !== undefined ? { inputLimits } : {}),
    ...(compat !== undefined ? { compat } : {}),
  };
}

function decodeEntry(value: unknown): ModelsStoreEntry | undefined {
  if (!Value.Check(catalogEntrySchema, value)) return undefined;

  const models: Model<Api>[] = [];
  for (const rawModel of value.models) {
    const model = decodeCatalogModel(rawModel);
    if (model) models.push(model);
  }

  const { checkedAt, lastModified, etag } = value;
  return {
    models,
    ...(checkedAt !== undefined ? { checkedAt } : {}),
    ...(lastModified !== undefined ? { lastModified } : {}),
    ...(etag !== undefined ? { etag } : {}),
  };
}

/** IndexedDB-backed dynamic provider catalogues, keyed by runtime provider ID. */
export class ModelCatalogsStore extends Store implements ModelsStore {
  getConfig(): StoreConfig {
    return { name: "model-catalogs" };
  }

  async read(providerId: string): Promise<ModelsStoreEntry | undefined> {
    const raw = await this.getBackend().get<unknown>("model-catalogs", providerId);
    return decodeEntry(raw);
  }

  async write(providerId: string, entry: ModelsStoreEntry): Promise<void> {
    await this.getBackend().set("model-catalogs", providerId, entry);
  }

  async delete(providerId: string): Promise<void> {
    await this.getBackend().delete("model-catalogs", providerId);
  }
}
