import type {
  Api,
  Model,
  ModelsStore,
  ModelsStoreEntry,
} from "@earendil-works/pi-ai";

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

function isModelCatalogPayload(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseFiniteNumber(value: DynamicValue): number | null {
  return typeof value === "number" && Number.isFinite(value) ? value : null;
}

function parseCost(value: DynamicValue): Model<Api>["cost"] | null {
  if (!isModelCatalogPayload(value)) return null;

  const input = parseFiniteNumber(value.input);
  const output = parseFiniteNumber(value.output);
  const cacheRead = parseFiniteNumber(value.cacheRead);
  const cacheWrite = parseFiniteNumber(value.cacheWrite);
  if (input === null || output === null || cacheRead === null || cacheWrite === null) {
    return null;
  }

  return { input, output, cacheRead, cacheWrite };
}

function parseInputKinds(value: DynamicValue): Model<Api>["input"] | null {
  if (!Array.isArray(value)) return null;

  const parsed: Model<Api>["input"] = [];
  for (const item of value) {
    if (item !== "text" && item !== "image") return null;
    parsed.push(item === "text" ? "text" : "image");
  }
  return parsed;
}

function parseHeaders(value: DynamicValue): Record<string, string> | undefined | null {
  if (value === undefined) return undefined;
  if (!isModelCatalogPayload(value)) return null;

  const parsed: Record<string, string> = {};
  for (const [name, headerValue] of Object.entries(value)) {
    if (typeof headerValue !== "string") return null;
    parsed[name] = headerValue;
  }
  return parsed;
}

function parseModel(value: DynamicValue): Model<Api> | null {
  if (!isModelCatalogPayload(value)) return null;

  const { id, name, api, provider, baseUrl, reasoning } = value;
  const contextWindow = parseFiniteNumber(value.contextWindow);
  const maxTokens = parseFiniteNumber(value.maxTokens);
  const input = parseInputKinds(value.input);
  const cost = parseCost(value.cost);
  const headers = parseHeaders(value.headers);

  if (
    typeof id !== "string"
    || typeof name !== "string"
    || typeof api !== "string"
    || typeof provider !== "string"
    || typeof baseUrl !== "string"
    || typeof reasoning !== "boolean"
    || contextWindow === null
    || maxTokens === null
    || !input
    || !cost
    || headers === null
  ) {
    return null;
  }

  return {
    id,
    name,
    api,
    provider,
    baseUrl,
    reasoning,
    input,
    cost,
    contextWindow,
    maxTokens,
    ...(headers !== undefined ? { headers } : {}),
  };
}

function parseEntry(value: DynamicValue): ModelsStoreEntry | undefined {
  if (!isModelCatalogPayload(value) || !Array.isArray(value.models)) {
    return undefined;
  }

  const models: Model<Api>[] = [];
  for (const rawModel of value.models) {
    const model = parseModel(rawModel);
    if (model) models.push(model);
  }

  const checkedAt = parseFiniteNumber(value.checkedAt);
  return {
    models,
    ...(checkedAt !== null ? { checkedAt } : {}),
  };
}

/** IndexedDB-backed dynamic provider catalogues, keyed by runtime provider ID. */
export class ModelCatalogsStore extends Store implements ModelsStore {
  getConfig(): StoreConfig {
    return { name: "model-catalogs" };
  }

  async read(providerId: string): Promise<ModelsStoreEntry | undefined> {
    const raw = await this.getBackend().get<DynamicValue>("model-catalogs", providerId);
    return parseEntry(raw);
  }

  async write(providerId: string, entry: ModelsStoreEntry): Promise<void> {
    await this.getBackend().set("model-catalogs", providerId, entry);
  }

  async delete(providerId: string): Promise<void> {
    await this.getBackend().delete("model-catalogs", providerId);
  }
}
