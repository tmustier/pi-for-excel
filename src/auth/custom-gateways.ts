/**
 * Helpers for custom OpenAI-compatible gateway providers.
 */

import type { Api, Model } from "@mariozechner/pi-ai";
import type { CustomProvider } from "@mariozechner/pi-web-ui/dist/storage/stores/custom-providers-store.js";

const OPENAI_GATEWAY_ID_PREFIX = "pi-openai-gateway:";
export const OPENAI_GATEWAY_PROVIDER_PREFIX = "Gateway · ";
const OPENAI_GATEWAY_TYPE = "openai-completions";

const DEFAULT_CONTEXT_WINDOW = 16_384;
const DEFAULT_MAX_TOKENS = 4_096;

type StoredModel = NonNullable<CustomProvider["models"]>[number];

export interface CustomProvidersStoreLike {
  get(id: string): Promise<CustomProvider | null>;
  set(provider: CustomProvider): Promise<void>;
  delete(id: string): Promise<void>;
  getAll(): Promise<CustomProvider[]>;
}

export interface OpenAiGatewayConfig {
  id: string;
  displayName: string;
  endpointUrl: string;
  modelId: string;
  apiKey: string;
  providerName: string;
}

export interface SaveOpenAiGatewayInput {
  id?: string;
  displayName?: string;
  endpointUrl: string;
  modelId: string;
  apiKey?: string;
}

export interface CustomProviderRuntimeInfo {
  providerNames: Set<string>;
  apiKeys: Map<string, string | undefined>;
  defaultModel: Model<Api> | null;
}

function normalizeOptionalString(value: string | null | undefined): string {
  if (typeof value !== "string") {
    return "";
  }

  return value.trim();
}

function normalizeRequiredString(value: string, label: string): string {
  const trimmed = value.trim();
  if (trimmed.length === 0) {
    throw new Error(`${label} is required.`);
  }

  return trimmed;
}

export function normalizeGatewayEndpointUrl(endpointUrl: string): string {
  const raw = normalizeRequiredString(endpointUrl, "Endpoint URL");

  let parsed: URL;
  try {
    parsed = new URL(raw);
  } catch {
    throw new Error("Endpoint URL must be a valid http:// or https:// URL.");
  }

  if (parsed.protocol !== "http:" && parsed.protocol !== "https:") {
    throw new Error("Endpoint URL must use http:// or https://.");
  }

  parsed.hash = "";
  const normalized = parsed.toString();
  return normalized.endsWith("/") ? normalized.slice(0, -1) : normalized;
}

export function normalizeGatewayModelId(modelId: string): string {
  return normalizeRequiredString(modelId, "Model ID");
}

function deriveDisplayName(rawName: string | undefined, endpointUrl: string): string {
  const explicit = normalizeOptionalString(rawName);
  if (explicit.length > 0) {
    return explicit;
  }

  try {
    const url = new URL(endpointUrl);
    const host = url.port.length > 0 ? `${url.hostname}:${url.port}` : url.hostname;
    if (host.trim().length > 0) {
      return host;
    }
  } catch {
    // noop — endpoint URL is validated before this runs.
  }

  return "Custom gateway";
}

function toGatewayProviderName(displayName: string): string {
  return `${OPENAI_GATEWAY_PROVIDER_PREFIX}${displayName}`;
}

function getFirstModel(provider: CustomProvider): StoredModel | null {
  if (!Array.isArray(provider.models) || provider.models.length === 0) {
    return null;
  }

  return provider.models[0] ?? null;
}

export function isOpenAiGatewayProvider(provider: CustomProvider): boolean {
  if (!provider.id.startsWith(OPENAI_GATEWAY_ID_PREFIX)) {
    return false;
  }

  if (provider.type !== OPENAI_GATEWAY_TYPE) {
    return false;
  }

  const model = getFirstModel(provider);
  if (!model) {
    return false;
  }

  const providerName = normalizeOptionalString(model.provider);
  const modelId = normalizeOptionalString(model.id);
  return providerName.length > 0 && modelId.length > 0;
}

function providerToGatewayConfig(provider: CustomProvider): OpenAiGatewayConfig | null {
  if (!isOpenAiGatewayProvider(provider)) {
    return null;
  }

  const model = getFirstModel(provider);
  if (!model) {
    return null;
  }

  const providerName = normalizeOptionalString(model.provider);
  if (providerName.length === 0) {
    return null;
  }

  const endpointUrl = normalizeOptionalString(provider.baseUrl);
  const modelId = normalizeOptionalString(model.id);
  if (endpointUrl.length === 0 || modelId.length === 0) {
    return null;
  }

  const defaultDisplayName = providerName.startsWith(OPENAI_GATEWAY_PROVIDER_PREFIX)
    ? providerName.slice(OPENAI_GATEWAY_PROVIDER_PREFIX.length)
    : providerName;

  return {
    id: provider.id,
    displayName: normalizeOptionalString(provider.name) || defaultDisplayName,
    endpointUrl,
    modelId,
    apiKey: normalizeOptionalString(provider.apiKey),
    providerName,
  };
}

function createGatewayModel(args: {
  endpointUrl: string;
  modelId: string;
  providerName: string;
}): Model<"openai-completions"> {
  return {
    id: args.modelId,
    name: args.modelId,
    api: "openai-completions",
    provider: args.providerName,
    baseUrl: args.endpointUrl,
    reasoning: false,
    input: ["text"],
    cost: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
    },
    contextWindow: DEFAULT_CONTEXT_WINDOW,
    maxTokens: DEFAULT_MAX_TOKENS,
  };
}

function resolveUniqueProviderName(args: {
  displayName: string;
  existingGateways: OpenAiGatewayConfig[];
  editingId?: string;
}): string {
  const candidateBase = toGatewayProviderName(args.displayName);
  const usedNames = new Set(
    args.existingGateways
      .filter((gateway) => gateway.id !== args.editingId)
      .map((gateway) => gateway.providerName.toLowerCase()),
  );

  if (!usedNames.has(candidateBase.toLowerCase())) {
    return candidateBase;
  }

  let suffix = 2;
  while (suffix < 500) {
    const candidate = `${candidateBase} (${suffix})`;
    if (!usedNames.has(candidate.toLowerCase())) {
      return candidate;
    }

    suffix += 1;
  }

  return `${candidateBase} (${Date.now()})`;
}

export async function listOpenAiGatewayConfigs(
  customProvidersStore: CustomProvidersStoreLike,
): Promise<OpenAiGatewayConfig[]> {
  const providers = await customProvidersStore.getAll();

  return providers
    .map((provider) => providerToGatewayConfig(provider))
    .filter((gateway): gateway is OpenAiGatewayConfig => gateway !== null)
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export async function saveOpenAiGatewayConfig(
  customProvidersStore: CustomProvidersStoreLike,
  input: SaveOpenAiGatewayInput,
): Promise<OpenAiGatewayConfig> {
  const endpointUrl = normalizeGatewayEndpointUrl(input.endpointUrl);
  const modelId = normalizeGatewayModelId(input.modelId);
  const existingGateways = await listOpenAiGatewayConfigs(customProvidersStore);

  if (input.id) {
    const match = existingGateways.find((gateway) => gateway.id === input.id);
    if (!match) {
      throw new Error("Gateway not found.");
    }
  }

  const displayName = deriveDisplayName(input.displayName, endpointUrl);
  const providerName = resolveUniqueProviderName({
    displayName,
    existingGateways,
    editingId: input.id,
  });

  const id = input.id ?? `${OPENAI_GATEWAY_ID_PREFIX}${crypto.randomUUID()}`;
  const apiKey = normalizeOptionalString(input.apiKey);

  const provider: CustomProvider = {
    id,
    name: displayName,
    type: OPENAI_GATEWAY_TYPE,
    baseUrl: endpointUrl,
    apiKey: apiKey.length > 0 ? apiKey : undefined,
    models: [
      createGatewayModel({
        endpointUrl,
        modelId,
        providerName,
      }),
    ],
  };

  await customProvidersStore.set(provider);

  return {
    id,
    displayName,
    endpointUrl,
    modelId,
    apiKey,
    providerName,
  };
}

export async function deleteOpenAiGatewayConfig(
  customProvidersStore: CustomProvidersStoreLike,
  id: string,
): Promise<void> {
  const normalizedId = normalizeOptionalString(id);
  if (!normalizedId.startsWith(OPENAI_GATEWAY_ID_PREFIX)) {
    return;
  }

  const existing = await customProvidersStore.get(normalizedId);
  if (!existing || !isOpenAiGatewayProvider(existing)) {
    return;
  }

  await customProvidersStore.delete(normalizedId);
}

export function collectCustomProviderRuntimeInfo(
  customProviders: CustomProvider[],
): CustomProviderRuntimeInfo {
  const providerNames = new Set<string>();
  const apiKeys = new Map<string, string | undefined>();
  let defaultModel: Model<Api> | null = null;

  for (const provider of customProviders) {
    const apiKey = normalizeOptionalString(provider.apiKey);
    const namesForProvider = new Set<string>();

    const storedModels = Array.isArray(provider.models) ? provider.models : [];

    if (storedModels.length > 0) {
      for (const model of storedModels) {
        const modelProviderName = normalizeOptionalString(model.provider);
        if (modelProviderName.length > 0) {
          namesForProvider.add(modelProviderName);
        }

        if (defaultModel === null) {
          defaultModel = model;
        }
      }
    } else {
      const directName = normalizeOptionalString(provider.name);
      if (directName.length > 0) {
        namesForProvider.add(directName);
      }
    }

    for (const providerName of namesForProvider) {
      providerNames.add(providerName);

      const hasExisting = apiKeys.has(providerName);
      if (apiKey.length > 0) {
        apiKeys.set(providerName, apiKey);
        continue;
      }

      if (!hasExisting) {
        apiKeys.set(providerName, undefined);
      }
    }
  }

  return {
    providerNames,
    apiKeys,
    defaultModel,
  };
}
