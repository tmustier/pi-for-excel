/**
 * Default model selection for the taskpane.
 */

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
  // Gemini defaults: Pro-ish first, then any Gemini
  { provider: "google", match: /^gemini-.*-pro/i },
  { provider: "google", match: /^gemini-/i },

  // Google Cloud Code Assist (Gemini CLI)
  // Prefer stable Gemini variants before preview models.
  { provider: "google-gemini-cli", match: /^gemini-(?!.*preview).*?-pro/i },
  { provider: "google-gemini-cli", match: /^gemini-(?!.*preview)/i },
  { provider: "google-gemini-cli", match: /^gemini-.*-pro/i },
  { provider: "google-gemini-cli", match: /^gemini-/i },

  // Google Antigravity (Gemini/Claude/GPT-OSS)
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
  // OpenAI special-case:
  // Prefer the newest general GPT model when it is at least as new as Codex,
  // while keeping Codex as fallback. Within one version Sol is the preferred
  // tier; upstream intentionally exposes no bare gpt-5.6 or gpt-6 alias.
  for (const provider of ["openai", "openai-codex"] as const) {
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

  // Anthropic special-case:
  // Default to the latest Opus while Fable is in the registry but unavailable
  // for normal Anthropic use. Keep Sonnet/Fable as fallbacks for resilience if
  // a future registry/provider configuration has no Opus entry.
  if (availableProviders.includes("anthropic")) {
    const models = [...modelsRuntime.getModels("anthropic")]
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id));
    for (const prefix of ["claude-opus-", "claude-sonnet-", "claude-fable-"]) {
      const best = models.find((m) => m.id.startsWith(prefix));
      if (best) return best;
    }
  }

  // Other providers: pattern-based rules
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

  // Any-available-provider fallback (#553): never prefer a model from an
  // unconfigured provider while a configured provider has registry models
  // (e.g. github-copilot- or mistral-only setups). An unusable default just
  // produces an API-key prompt for the wrong provider.
  for (const provider of availableProviders) {
    const best = [...modelsRuntime.getModels(provider)].sort(compareByRecency)[0];
    if (best) return best;
  }

  // Absolute fallback: keep this resilient across pi-ai version bumps.
  // Sol is the maintainer-preferred GPT quality tier.
  const preferredFallback = modelsRuntime.getModel("openai", "gpt-6.1-sol");
  if (preferredFallback) return preferredFallback;

  const firstKnownModel = modelsRuntime.getModels()[0];
  if (firstKnownModel) return firstKnownModel;
  throw new Error("No models are registered.");
}
