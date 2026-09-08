import {
  createModels,
  createProvider,
  lazyApi,
  type Api,
  type ApiKeyAuth,
  type Model,
  type ModelsRefreshOptions,
  type ModelsRefreshResult,
  type ModelsStore,
  type ModelsStoreEntry,
  type MutableModels,
  type Provider,
  type ProviderStreams,
} from "@earendil-works/pi-ai";
import { builtinProviders } from "@earendil-works/pi-ai/providers/all";

import { originalFetch } from "../auth/cors-proxy.js";
import { normalizeProxyUrl } from "../auth/proxy-validation.js";
import type { CustomProvider } from "../storage/local/custom-providers-store.js";
import {
  ProviderCredentialsStore,
  type ProviderKeysStoreLike,
} from "../storage/local/provider-credentials-store.js";

const DEFAULT_DISCOVERED_CONTEXT_WINDOW = 32_768;
const DEFAULT_DISCOVERED_MAX_TOKENS = 4_096;
const MODEL_DISCOVERY_TIMEOUT_MS = 8_000;
const MAX_MODEL_DISCOVERY_RESPONSE_BYTES = 2 * 1024 * 1024;
const MAX_DISCOVERED_MODELS = 2_000;
const MAX_DISCOVERED_MODEL_ID_LENGTH = 256;

const openAiCompletionsStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/openai-completions"),
);
const openAiResponsesStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/openai-responses"),
);
const anthropicMessagesStreams: ProviderStreams = lazyApi(
  () => import("@earendil-works/pi-ai/api/anthropic-messages"),
);

export type BrowserProviderApi =
  | "openai-completions"
  | "openai-responses"
  | "anthropic-messages";

export interface BrowserProviderModelDefinition {
  id: string;
  name?: string;
  reasoning?: boolean;
  input?: readonly ("text" | "image")[];
  contextWindow?: number;
  maxTokens?: number;
}

export interface BrowserProviderRegistration {
  id: string;
  name: string;
  api: BrowserProviderApi;
  baseUrl: string;
  models: readonly BrowserProviderModelDefinition[];
  /** Defaults to `${baseUrl}/models` for OpenAI-compatible providers. */
  modelsUrl?: string;
  resolveApiKey: () => Promise<string | undefined>;
  /** Keyless local providers can opt in while still satisfying Pi AI auth semantics. */
  allowKeyless?: boolean;
}

export interface CreateBrowserModelRuntimeOptions {
  providerKeys: ProviderKeysStoreLike;
  modelCatalogs: ModelsStore;
  getProxyUrl: () => Promise<string | undefined>;
  fetchFn?: typeof globalThis.fetch;
}

interface DynamicProviderCatalogSpec {
  api: BrowserProviderApi;
  baseUrl: string;
}

function normalizeNonEmpty(value: string, label: string): string {
  const normalized = value.trim();
  if (normalized.length === 0) {
    throw new Error(`${label} cannot be empty.`);
  }
  return normalized;
}

function normalizeHttpUrl(value: string, label: string): string {
  let parsed: URL;
  try {
    parsed = new URL(normalizeNonEmpty(value, label));
  } catch {
    throw new Error(`${label} must be a valid http:// or https:// URL.`);
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error(`${label} must use http:// or https://.`);
  }

  parsed.hash = "";
  const normalized = parsed.toString();
  return normalized.endsWith("/") ? normalized.slice(0, -1) : normalized;
}

function streamsForApi(api: BrowserProviderApi): ProviderStreams {
  if (api === "openai-completions") return openAiCompletionsStreams;
  if (api === "openai-responses") return openAiResponsesStreams;
  return anthropicMessagesStreams;
}

function createStoredApiKeyAuth(args: {
  name: string;
  resolveApiKey: () => Promise<string | undefined>;
  allowKeyless: boolean;
}): ApiKeyAuth {
  return {
    name: args.name,
    async resolve() {
      const key = (await args.resolveApiKey())?.trim();
      if (!key && !args.allowKeyless) return undefined;

      return {
        auth: key ? { apiKey: key } : {},
        source: key ? "Browser credential store" : "Keyless provider",
      };
    },
  };
}

function createBrowserAdapterProvider(provider: Provider): Provider {
  const authName = provider.auth.oauth?.name ?? provider.name;
  return createProvider({
    id: provider.id,
    name: provider.name,
    ...(provider.baseUrl !== undefined ? { baseUrl: provider.baseUrl } : {}),
    ...(provider.headers !== undefined ? { headers: provider.headers } : {}),
    auth: {
      apiKey: {
        name: authName,
        resolve({ credential }) {
          if (credential?.key) {
            return Promise.resolve({
              auth: { apiKey: credential.key },
              source: "Browser credential store",
            });
          }
          return Promise.resolve(undefined);
        },
      },
    },
    models: provider.getModels(),
    api: {
      stream: (model, context, options) => provider.stream(model, context, options),
      streamSimple: (model, context, options) => provider.streamSimple(model, context, options),
    },
  });
}

function createModel(args: {
  providerId: string;
  api: BrowserProviderApi;
  baseUrl: string;
  definition: BrowserProviderModelDefinition;
}): Model<Api> {
  const id = normalizeNonEmpty(args.definition.id, "Model id");
  const contextWindow = args.definition.contextWindow ?? DEFAULT_DISCOVERED_CONTEXT_WINDOW;
  const maxTokens = args.definition.maxTokens ?? Math.min(DEFAULT_DISCOVERED_MAX_TOKENS, contextWindow);

  if (!Number.isInteger(contextWindow) || contextWindow < 1_024) {
    throw new Error(`Model ${id} contextWindow must be an integer of at least 1024.`);
  }
  if (!Number.isInteger(maxTokens) || maxTokens < 1) {
    throw new Error(`Model ${id} maxTokens must be a positive integer.`);
  }

  return {
    id,
    name: args.definition.name?.trim() || id,
    api: args.api,
    provider: args.providerId,
    baseUrl: args.baseUrl,
    reasoning: args.definition.reasoning ?? false,
    input: args.definition.input ? [...args.definition.input] : ["text"],
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    },
    contextWindow,
    maxTokens,
  };
}

function deriveModelsUrl(baseUrl: string, api: BrowserProviderApi): string | undefined {
  if (api !== "openai-completions" && api !== "openai-responses") {
    return undefined;
  }
  return `${baseUrl}/models`;
}

function sanitizeCatalogEntry(
  providerId: string,
  spec: DynamicProviderCatalogSpec,
  entry: ModelsStoreEntry,
): ModelsStoreEntry {
  const models: Model<Api>[] = [];
  for (const model of entry.models) {
    try {
      models.push(createModel({
        providerId,
        api: spec.api,
        baseUrl: spec.baseUrl,
        definition: {
          id: model.id,
          name: model.name,
          reasoning: model.reasoning,
          input: model.input,
          contextWindow: model.contextWindow,
          maxTokens: model.maxTokens,
        },
      }));
    } catch {
      // Ignore malformed/stale entries rather than restoring unsafe metadata.
    }
  }

  return {
    models,
    ...(entry.checkedAt !== undefined ? { checkedAt: entry.checkedAt } : {}),
  };
}

/** Rebind cached discovery results to the currently registered transport. */
class BrowserModelCatalogsStore implements ModelsStore {
  private readonly persisted: ModelsStore;
  private readonly specs = new Map<string, DynamicProviderCatalogSpec>();

  constructor(persisted: ModelsStore) {
    this.persisted = persisted;
  }

  configure(providerId: string, spec: DynamicProviderCatalogSpec): void {
    this.specs.set(providerId, spec);
  }

  unconfigure(providerId: string): void {
    this.specs.delete(providerId);
  }

  async read(providerId: string): Promise<ModelsStoreEntry | undefined> {
    const entry = await this.persisted.read(providerId);
    const spec = this.specs.get(providerId);
    return entry && spec ? sanitizeCatalogEntry(providerId, spec, entry) : entry;
  }

  write(providerId: string, entry: ModelsStoreEntry): Promise<void> {
    const spec = this.specs.get(providerId);
    return this.persisted.write(
      providerId,
      spec ? sanitizeCatalogEntry(providerId, spec, entry) : entry,
    );
  }

  delete(providerId: string): Promise<void> {
    return this.persisted.delete(providerId);
  }
}

function isModelsResponse(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function parseModelIds(value: DynamicValue): string[] {
  if (!isModelsResponse(value) || !Array.isArray(value.data)) {
    throw new Error("Model discovery response must contain a data array.");
  }
  if (value.data.length > MAX_DISCOVERED_MODELS) {
    throw new Error(`Model discovery returned more than ${MAX_DISCOVERED_MODELS} entries.`);
  }

  const ids: string[] = [];
  for (const entry of value.data) {
    if (!isModelsResponse(entry) || typeof entry.id !== "string") continue;
    const id = entry.id.trim();
    if (id.length === 0) continue;
    if (id.length > MAX_DISCOVERED_MODEL_ID_LENGTH) {
      throw new Error(
        `Discovered model id exceeds ${MAX_DISCOVERED_MODEL_ID_LENGTH} characters.`,
      );
    }
    ids.push(id);
  }

  return Array.from(new Set(ids)).sort((left, right) => left.localeCompare(right));
}

async function readLimitedDiscoveryJson(response: Response): Promise<DynamicValue> {
  const declaredLengthRaw = response.headers.get("content-length");
  const declaredLength = declaredLengthRaw === null
    ? null
    : Number.parseInt(declaredLengthRaw, 10);
  if (
    declaredLength !== null
    && Number.isFinite(declaredLength)
    && declaredLength > MAX_MODEL_DISCOVERY_RESPONSE_BYTES
  ) {
    throw new Error(
      `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
    );
  }

  if (!response.body) {
    const text = await response.text();
    if (new TextEncoder().encode(text).byteLength > MAX_MODEL_DISCOVERY_RESPONSE_BYTES) {
      throw new Error(
        `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
      );
    }
    const parsed: DynamicValue = JSON.parse(text);
    return parsed;
  }

  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let totalBytes = 0;

  while (true) {
    const chunk = await reader.read();
    if (chunk.done) break;
    if (!chunk.value) continue;

    totalBytes += chunk.value.byteLength;
    if (totalBytes > MAX_MODEL_DISCOVERY_RESPONSE_BYTES) {
      await reader.cancel();
      throw new Error(
        `Model discovery response exceeds ${MAX_MODEL_DISCOVERY_RESPONSE_BYTES} bytes.`,
      );
    }
    chunks.push(chunk.value);
  }

  const body = new Uint8Array(totalBytes);
  let offset = 0;
  for (const chunk of chunks) {
    body.set(chunk, offset);
    offset += chunk.byteLength;
  }

  const parsed: DynamicValue = JSON.parse(new TextDecoder().decode(body));
  return parsed;
}

async function resolveDiscoveryRequestUrl(
  targetUrl: string,
  getProxyUrl: () => Promise<string | undefined>,
): Promise<string> {
  const proxyUrl = await getProxyUrl();
  if (!proxyUrl) return targetUrl;
  return `${normalizeProxyUrl(proxyUrl)}/?url=${encodeURIComponent(targetUrl)}`;
}

async function fetchWithDiscoveryTimeout(
  fetchFn: typeof globalThis.fetch,
  requestUrl: string,
  headers: Record<string, string>,
  signal?: AbortSignal,
  lifecycleSignal?: AbortSignal,
): Promise<Response> {
  const controller = new AbortController();
  const handleAbort = (): void => controller.abort();
  const signals: readonly (AbortSignal | undefined)[] = [signal, lifecycleSignal];
  for (const sourceSignal of signals) {
    if (sourceSignal?.aborted) controller.abort();
    sourceSignal?.addEventListener("abort", handleAbort, { once: true });
  }
  const timeoutId = setTimeout(() => controller.abort(), MODEL_DISCOVERY_TIMEOUT_MS);

  try {
    return await fetchFn(requestUrl, {
      headers,
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeoutId);
    for (const sourceSignal of signals) {
      sourceSignal?.removeEventListener("abort", handleAbort);
    }
  }
}

function createRegisteredProvider(
  registration: BrowserProviderRegistration,
  getProxyUrl: () => Promise<string | undefined>,
  fetchFn: typeof globalThis.fetch,
  lifecycleSignal: AbortSignal,
): Provider {
  const id = normalizeNonEmpty(registration.id, "Provider id");
  const name = normalizeNonEmpty(registration.name, "Provider name");
  const baseUrl = normalizeHttpUrl(registration.baseUrl, "Provider baseUrl");
  const baselineModels = registration.models.map((definition) => createModel({
    providerId: id,
    api: registration.api,
    baseUrl,
    definition,
  }));
  const modelsUrlRaw = registration.modelsUrl
    ?? deriveModelsUrl(baseUrl, registration.api);
  const modelsUrl = modelsUrlRaw
    ? normalizeHttpUrl(modelsUrlRaw, "Provider modelsUrl")
    : undefined;

  return createProvider({
    id,
    name,
    baseUrl,
    auth: {
      apiKey: createStoredApiKeyAuth({
        name: `${name} credential`,
        resolveApiKey: registration.resolveApiKey,
        allowKeyless: registration.allowKeyless === true,
      }),
    },
    models: baselineModels,
    ...(modelsUrl !== undefined
      ? {
        fetchModels: async ({ credential, signal }) => {
          const requestUrl = await resolveDiscoveryRequestUrl(modelsUrl, getProxyUrl);
          const headers: Record<string, string> = { Accept: "application/json" };
          if (credential?.type === "api_key" && credential.key) {
            headers.Authorization = `Bearer ${credential.key}`;
          }

          const response = await fetchWithDiscoveryTimeout(
            fetchFn,
            requestUrl,
            headers,
            signal,
            lifecycleSignal,
          );
          if (!response.ok) {
            throw new Error(`Model discovery failed with HTTP ${response.status}.`);
          }

          const ids = parseModelIds(await readLimitedDiscoveryJson(response));
          const template = registration.models[0];
          return ids.map((modelId) => createModel({
            providerId: id,
            api: registration.api,
            baseUrl,
            definition: {
              id: modelId,
              ...(template?.reasoning !== undefined ? { reasoning: template.reasoning } : {}),
              ...(template?.input !== undefined ? { input: template.input } : {}),
              ...(template?.contextWindow !== undefined ? { contextWindow: template.contextWindow } : {}),
              ...(template?.maxTokens !== undefined ? { maxTokens: template.maxTokens } : {}),
            },
          }));
        },
      }
      : {}),
    api: streamsForApi(registration.api),
  });
}

function customProviderRegistrations(provider: CustomProvider): BrowserProviderRegistration[] {
  const storedModels = provider.models ?? [];
  const modelsByProvider = new Map<string, BrowserProviderModelDefinition[]>();

  for (const model of storedModels) {
    const definitions = modelsByProvider.get(model.provider) ?? [];
    definitions.push({
      id: model.id,
      name: model.name,
      reasoning: model.reasoning,
      input: model.input,
      contextWindow: model.contextWindow,
      maxTokens: model.maxTokens,
    });
    modelsByProvider.set(model.provider, definitions);
  }

  if (modelsByProvider.size === 0) return [];
  if (
    provider.type !== "openai-completions"
    && provider.type !== "openai-responses"
    && provider.type !== "anthropic-messages"
  ) {
    return [];
  }

  const providerApi: BrowserProviderApi = provider.type;
  return Array.from(modelsByProvider.entries()).map(([providerId, models]) => ({
    id: providerId,
    name: provider.name,
    api: providerApi,
    baseUrl: provider.baseUrl,
    models,
    resolveApiKey: () => Promise.resolve(provider.apiKey),
    allowKeyless: !provider.apiKey,
  }));
}

/** Browser-native Pi AI provider runtime with IndexedDB credentials/catalogues. */
export class BrowserModelRuntime {
  readonly models: MutableModels;

  private readonly modelCatalogs: BrowserModelCatalogsStore;
  private readonly getProxyUrl: () => Promise<string | undefined>;
  private readonly fetchFn: typeof globalThis.fetch;
  private readonly builtinProviderIds = new Set<string>();
  private readonly customProviderIds = new Set<string>();
  private readonly extensionProviderOwners = new Map<string, string>();
  private readonly dynamicProviderAbortControllers = new Map<string, AbortController>();

  constructor(options: CreateBrowserModelRuntimeOptions) {
    this.modelCatalogs = new BrowserModelCatalogsStore(options.modelCatalogs);
    this.getProxyUrl = options.getProxyUrl;
    this.fetchFn = options.fetchFn ?? originalFetch ?? globalThis.fetch;
    this.models = createModels({
      credentials: new ProviderCredentialsStore(options.providerKeys),
      modelsStore: this.modelCatalogs,
      authContext: {
        env: () => Promise.resolve(undefined),
        fileExists: () => Promise.resolve(false),
      },
    });

    for (const provider of builtinProviders()) {
      const browserProvider = provider.auth.apiKey
        ? provider
        : createBrowserAdapterProvider(provider);
      this.models.setProvider(browserProvider);
      this.builtinProviderIds.add(browserProvider.id);
    }
  }

  private createDynamicProvider(registration: BrowserProviderRegistration): Provider {
    const controller = new AbortController();
    const provider = createRegisteredProvider(
      registration,
      this.getProxyUrl,
      this.fetchFn,
      controller.signal,
    );

    this.dynamicProviderAbortControllers.get(provider.id)?.abort();
    this.dynamicProviderAbortControllers.set(provider.id, controller);
    return provider;
  }

  private stopDynamicProvider(providerId: string): void {
    this.dynamicProviderAbortControllers.get(providerId)?.abort();
    this.dynamicProviderAbortControllers.delete(providerId);
  }

  async syncCustomProviders(customProviders: readonly CustomProvider[]): Promise<void> {
    const nextIds = new Set<string>();

    for (const customProvider of customProviders) {
      for (const registration of customProviderRegistrations(customProvider)) {
        if (this.builtinProviderIds.has(registration.id)) {
          throw new Error(`Custom provider id conflicts with built-in provider: ${registration.id}`);
        }
        if (this.extensionProviderOwners.has(registration.id)) {
          throw new Error(`Custom provider id conflicts with extension provider: ${registration.id}`);
        }

        const provider = this.createDynamicProvider(registration);
        this.modelCatalogs.configure(provider.id, {
          api: registration.api,
          baseUrl: provider.baseUrl ?? normalizeHttpUrl(registration.baseUrl, "Provider baseUrl"),
        });
        this.models.setProvider(provider);
        nextIds.add(registration.id);
      }
    }

    for (const previousId of this.customProviderIds) {
      if (nextIds.has(previousId)) continue;
      this.stopDynamicProvider(previousId);
      this.models.deleteProvider(previousId);
      try {
        await this.modelCatalogs.delete(previousId);
      } finally {
        this.modelCatalogs.unconfigure(previousId);
      }
    }

    this.customProviderIds.clear();
    for (const providerId of nextIds) this.customProviderIds.add(providerId);
  }

  registerExtensionProvider(ownerId: string, registration: BrowserProviderRegistration): void {
    const providerId = normalizeNonEmpty(registration.id, "Provider id");
    if (this.builtinProviderIds.has(providerId) || this.customProviderIds.has(providerId)) {
      throw new Error(`Provider id is reserved: ${providerId}`);
    }

    const existingOwner = this.extensionProviderOwners.get(providerId);
    if (existingOwner && existingOwner !== ownerId) {
      throw new Error(`Provider id is already registered by another extension: ${providerId}`);
    }

    const provider = this.createDynamicProvider(registration);
    this.modelCatalogs.configure(provider.id, {
      api: registration.api,
      baseUrl: provider.baseUrl ?? normalizeHttpUrl(registration.baseUrl, "Provider baseUrl"),
    });
    this.models.setProvider(provider);
    this.extensionProviderOwners.set(providerId, ownerId);
  }

  isExtensionProvider(providerId: string): boolean {
    return this.extensionProviderOwners.has(providerId);
  }

  shouldProxyProvider(providerId: string): boolean {
    return this.customProviderIds.has(providerId) || this.extensionProviderOwners.has(providerId);
  }

  async unregisterExtensionProvider(ownerId: string, providerId: string): Promise<void> {
    if (this.extensionProviderOwners.get(providerId) !== ownerId) {
      throw new Error(`Provider is not owned by this extension: ${providerId}`);
    }

    this.extensionProviderOwners.delete(providerId);
    this.stopDynamicProvider(providerId);
    this.models.deleteProvider(providerId);
    try {
      await this.modelCatalogs.delete(providerId);
    } finally {
      this.modelCatalogs.unconfigure(providerId);
    }
  }

  async unregisterExtensionProviders(ownerId: string): Promise<void> {
    const ownedIds = Array.from(this.extensionProviderOwners.entries())
      .filter(([, owner]) => owner === ownerId)
      .map(([providerId]) => providerId);

    for (const providerId of ownedIds) {
      await this.unregisterExtensionProvider(ownerId, providerId);
    }
  }

  refresh(options?: ModelsRefreshOptions): Promise<ModelsRefreshResult> {
    return this.models.refresh(options);
  }
}
