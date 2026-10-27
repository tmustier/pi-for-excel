import {
  getModelType,
  isModelType,
  type AnyModel,
  type Api,
  type Model,
  type ModelsStoreEntry,
  type Provider,
} from "@earendil-works/pi-ai";
import { PI_AI_VERSION } from "@earendil-works/pi-ai/utils/version";

import { decodeCatalogModel } from "../storage/local/model-catalogs-store.js";

const CATALOG_REQUEST_TIMEOUT_MS = 8_000;
const MINIMUM_VERSION_HEADER = "x-pi-model-catalog-minimum-version";

export interface RemoteCatalogOptions {
  /** Same-origin relay for `https://pi.dev/api/models/providers/<id>?types=chat&pi-version=<pi-ai>`. */
  catalogUrl: (providerId: string) => string;
  fetchFn: typeof globalThis.fetch;
  /** Generation time of the bundled pi-ai model data; older catalogues are ignored. */
  bundledGeneratedAt: number | undefined;
}

/** True when the installed pi-ai is at least `minimum` (major.minor.patch). */
function supportsCatalogVersion(minimum: string): boolean {
  const [required, actual] = [minimum, PI_AI_VERSION]
    .map((version) => version.split(".").map((part) => Number.parseInt(part, 10) || 0));
  for (let index = 0; index < 3; index += 1) {
    const difference = (actual?.[index] ?? 0) - (required?.[index] ?? 0);
    if (difference !== 0) return difference > 0;
  }
  return true;
}

function routeKey(api: string, baseUrl: string): string {
  return `${api}\0${baseUrl}`;
}

function mergeModels<TModel extends AnyModel>(baseline: readonly TModel[], overlay: readonly TModel[]): TModel[] {
  const merged = new Map<string, TModel>();
  for (const model of [...baseline, ...overlay]) merged.set(`${getModelType(model)}\0${model.id}`, model);
  return [...merged.values()];
}

/**
 * Add pi.dev's model catalogue to a static built-in provider, as `pi update --models` does.
 *
 * The catalogue only describes models. Entries must use a transport (API and base URL) the
 * bundled provider already uses, and take its headers, so remote data cannot add a transport
 * this runtime cannot stream or send credentials anywhere new. Failures keep the last
 * published list, and a catalogue older than the bundled data is ignored.
 */
export function withRemoteCatalog(provider: Provider, options: RemoteCatalogOptions): Provider {
  const routes = new Map<string, Model<Api>["headers"]>();
  for (const model of provider.getModels()) {
    routes.set(routeKey(model.api, model.baseUrl), model.headers);
  }

  const supported = (models: readonly AnyModel[]): Model<Api>[] => {
    const result: Model<Api>[] = [];
    for (const model of models) {
      if (!isModelType(model, "chat") || model.provider !== provider.id) continue;
      const key = routeKey(model.api, model.baseUrl);
      if (!routes.has(key)) continue;
      const { headers: _remoteHeaders, ...described } = model;
      const headers = routes.get(key);
      result.push({ ...described, ...(headers !== undefined ? { headers } : {}) });
    }
    return result;
  };

  const overlayFor = (entry: ModelsStoreEntry | undefined): Model<Api>[] => {
    if (!entry) return [];
    const isNewer = options.bundledGeneratedAt === undefined
      || (entry.lastModified !== undefined && entry.lastModified > options.bundledGeneratedAt);
    return isNewer ? supported(entry.models) : [];
  };

  let overlay: readonly Model<Api>[] = [];

  return {
    ...provider,
    getModels: () => mergeModels(provider.getModels(), overlay),
    getAllModels: () => mergeModels<AnyModel>(provider.getAllModels?.() ?? provider.getModels(), overlay),
    refreshModels: async (context) => {
      const { stored } = context;
      const restored = overlayFor(stored);
      if (!(await context.publish({ update: () => { overlay = restored; } }))) return;
      if (!context.allowNetwork || context.signal.aborted) return;

      // Revalidate only when a cached body backs the validator, so a 304 never empties the overlay.
      const validator = stored && stored.models.length > 0 ? stored.etag : undefined;
      const response = await options.fetchFn(options.catalogUrl(provider.id), {
        headers: {
          Accept: "application/json",
          ...(validator !== undefined ? { "If-None-Match": validator } : {}),
        },
        signal: AbortSignal.any([context.signal, AbortSignal.timeout(CATALOG_REQUEST_TIMEOUT_MS)]),
      });
      if (context.signal.aborted) return;

      const checkedAt = Date.now();
      if (response.status === 304 && stored) {
        await context.publish({ persist: { ...stored, checkedAt } });
        return;
      }
      // No relay on this host (self-hosted without the rewrite): keep the bundled models.
      if (response.status === 404) return;
      if (!response.ok) {
        throw new Error(`Model catalogue request failed for ${provider.id}: HTTP ${response.status}`);
      }
      const minimumVersion = response.headers.get(MINIMUM_VERSION_HEADER);
      if (minimumVersion !== null && !supportsCatalogVersion(minimumVersion)) return;

      const body: unknown = await response.json();
      if (!Array.isArray(body)) {
        throw new Error(`Model catalogue for ${provider.id} is not a model list.`);
      }
      const models: Model<Api>[] = [];
      for (const value of body) {
        const model = decodeCatalogModel(value);
        if (model) models.push(model);
      }
      if (context.signal.aborted) return;

      const lastModified = Date.parse(response.headers.get("last-modified") ?? "");
      const etag = response.headers.get("etag");
      const entry: ModelsStoreEntry = {
        models,
        checkedAt,
        lastModified: Number.isNaN(lastModified) ? 0 : lastModified,
        ...(etag !== null ? { etag } : {}),
      };
      const published = overlayFor(entry);
      await context.publish({ persist: entry, update: () => { overlay = published; } });
    },
  };
}
