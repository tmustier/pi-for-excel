/**
 * Pure model-resolution helpers for the dev-only background verification bridge.
 *
 * The bridge must switch models through the same registry the product uses and
 * must reject unknown models or thinking levels deterministically. Keep this
 * module free of DOM/Office/storage side effects so it can be unit-tested in a
 * plain Node process; the DOM-facing bridge composes these results with the
 * production model-switch seam.
 */

import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import type { Api, Model } from "@earendil-works/pi-ai/compat";
import { getModels, getProviders } from "@earendil-works/pi-ai/compat";

import { getThinkingLevelsForModel } from "../models/thinking-levels.js";

export interface BridgeModelCandidate {
  provider: string;
  id: string;
  model: Model<Api>;
}

export interface BridgeModelResolutionOk {
  ok: true;
  provider: string;
  modelId: string;
  model: Model<Api>;
  thinkingLevel: ThinkingLevel;
  requestedThinkingLevel: ThinkingLevel | null;
  supportedThinkingLevels: ThinkingLevel[];
}

export interface BridgeModelResolutionError {
  ok: false;
  error: string;
  supportedThinkingLevels?: ThinkingLevel[];
  availableModelCount?: number;
}

export type BridgeModelResolution = BridgeModelResolutionOk | BridgeModelResolutionError;

/**
 * Collect the built-in registry models. Custom-provider models are gathered by
 * the DOM-facing bridge (they require async storage access) and concatenated
 * with this list before resolution.
 */
export function collectBuiltInModelCandidates(): BridgeModelCandidate[] {
  const candidates: BridgeModelCandidate[] = [];
  for (const provider of getProviders()) {
    for (const model of getModels(provider)) {
      candidates.push({ provider, id: model.id, model });
    }
  }
  return candidates;
}

/**
 * Faithful production default: reasoning models default to `high`, everything
 * else to `off`. Fall back to the first registry-supported level if that
 * default is somehow unsupported for a given model.
 */
export function productDefaultThinkingLevel(model: Model<Api>): ThinkingLevel {
  const supported = getThinkingLevelsForModel(model);
  const preferred: ThinkingLevel = model.reasoning ? "high" : "off";
  if (supported.includes(preferred)) return preferred;
  const first = supported[0];
  return first ?? "off";
}

function isThinkingLevelString(value: string): value is ThinkingLevel {
  return (
    value === "off"
    || value === "minimal"
    || value === "low"
    || value === "medium"
    || value === "high"
    || value === "xhigh"
    || value === "max"
  );
}

/**
 * Resolve an exact `provider`/`modelId` against the registry candidates and
 * validate the optional thinking level against the registry-supported set.
 * Returns a discriminated result so the bridge can fail closed with a precise
 * error instead of silently mutating runtime state.
 */
export function resolveBridgeModelSelection(args: {
  candidates: readonly BridgeModelCandidate[];
  provider: string;
  modelId: string;
  requestedThinkingLevel?: string | undefined;
}): BridgeModelResolution {
  const provider = args.provider.trim();
  const modelId = args.modelId.trim();
  if (!provider || !modelId) {
    return { ok: false, error: "selectModel requires non-empty provider and modelId" };
  }

  const match = args.candidates.find(
    (candidate) => candidate.provider === provider && candidate.id === modelId,
  );
  if (!match) {
    return {
      ok: false,
      error: `No registered model matches provider "${provider}" and modelId "${modelId}"`,
      availableModelCount: args.candidates.length,
    };
  }

  const supportedThinkingLevels = getThinkingLevelsForModel(match.model);

  const requestedRaw = typeof args.requestedThinkingLevel === "string"
    ? args.requestedThinkingLevel.trim()
    : "";

  if (!requestedRaw) {
    return {
      ok: true,
      provider,
      modelId,
      model: match.model,
      thinkingLevel: productDefaultThinkingLevel(match.model),
      requestedThinkingLevel: null,
      supportedThinkingLevels,
    };
  }

  if (!isThinkingLevelString(requestedRaw) || !supportedThinkingLevels.includes(requestedRaw)) {
    return {
      ok: false,
      error: `Thinking level "${requestedRaw}" is not supported by ${provider}/${modelId}`,
      supportedThinkingLevels,
    };
  }

  return {
    ok: true,
    provider,
    modelId,
    model: match.model,
    thinkingLevel: requestedRaw,
    requestedThinkingLevel: requestedRaw,
    supportedThinkingLevels,
  };
}
