import type { Api, Model, Models } from "@earendil-works/pi-ai";

import {
  compareOpenAiModelIds,
  isOpenAiCodexModelId,
  isOpenAiGeneralGptModelId,
  modelRecencyScore,
  shouldPreferOpenAiGeneralModel,
} from "../models/model-ordering.js";

type DefaultModelRule = {
  provider: "google" | "google-gemini-cli" | "google-antigravity";
  match: RegExp;
};

const DEFAULT_MODEL_RULES: DefaultModelRule[] = [
  { provider: "google", match: /^gemini-.*-pro/i },
  { provider: "google", match: /^gemini-/i },

  { provider: "google-gemini-cli", match: /^gemini-(?!.*preview).*?-pro/i },
  { provider: "google-gemini-cli", match: /^gemini-(?!.*preview)/i },
  { provider: "google-gemini-cli", match: /^gemini-.*-pro/i },
  { provider: "google-gemini-cli", match: /^gemini-/i },

  { provider: "google-antigravity", match: /^gemini-(?!.*preview).*?-pro/i },
  { provider: "google-antigravity", match: /^gemini-(?!.*preview)/i },
  { provider: "google-antigravity", match: /^gemini-.*-pro/i },
  { provider: "google-antigravity", match: /^gemini-/i },
  { provider: "google-antigravity", match: /^.+$/ },
];

function compareByRecency(a: Model<Api>, b: Model<Api>): number {
  return modelRecencyScore(b.id) - modelRecencyScore(a.id) || a.id.localeCompare(b.id);
}

export function pickDefaultModel(
  modelsRuntime: Models,
  availableProviders: string[],
  customDefaultModel?: Model<Api> | null,
): Model<Api> {
  for (const provider of ["openai-codex", "openai"] as const) {
    if (!availableProviders.includes(provider)) continue;
    const models = [...modelsRuntime.getModels(provider)]
      .sort((a, b) => compareOpenAiModelIds(a.id, b.id));
    const bestGpt = models.find((m) => isOpenAiGeneralGptModelId(m.id));
    const bestCodex = models.find((m) => isOpenAiCodexModelId(m.id));
    if (bestGpt && bestCodex) {
      return shouldPreferOpenAiGeneralModel(bestGpt.id, bestCodex.id) ? bestGpt : bestCodex;
    }
    const best = bestGpt ?? bestCodex ?? models[0];
    if (best) return best;
  }

  if (availableProviders.includes("anthropic")) {
    const models = [...modelsRuntime.getModels("anthropic")]
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id));
    for (const prefix of ["claude-opus-", "claude-sonnet-", "claude-fable-"]) {
      const best = models.find((m) => m.id.startsWith(prefix));
      if (best) return best;
    }
  }

  for (const rule of DEFAULT_MODEL_RULES) {
    if (!availableProviders.includes(rule.provider)) continue;
    const best = modelsRuntime.getModels(rule.provider)
      .filter((m) => rule.match.test(m.id))
      .sort(compareByRecency)[0];
    if (best) return best;
  }

  if (customDefaultModel) {
    return customDefaultModel;
  }

  for (const provider of availableProviders) {
    const best = [...modelsRuntime.getModels(provider)].sort(compareByRecency)[0];
    if (best) return best;
  }

  const preferredFallback = modelsRuntime.getModel("openai", "gpt-6.1-sol");
  if (preferredFallback) return preferredFallback;

  const firstKnownModel = modelsRuntime.getModels()[0];
  if (firstKnownModel) return firstKnownModel;
  throw new Error("No models are registered.");
}
