/**
 * Default model selection for the taskpane.
 */

import { getModel, getModels, type Api, type Model } from "@earendil-works/pi-ai";

import {
  compareOpenAiModelIds,
  isOpenAiCodexModelId,
  isOpenAiGeneralGptModelId,
  modelRecencyScore,
  parseMajorMinor,
  shouldPreferOpenAiGeneralModel,
} from "../models/model-ordering.js";

type DefaultProvider =
  | "openai-codex"
  | "openai"
  | "google"
  | "google-gemini-cli"
  | "google-antigravity";

type DefaultModelRule = { provider: DefaultProvider; match: RegExp };

function getProviderModels(provider: string): Model<Api>[] {
  return getModels(provider as Parameters<typeof getModels>[0]) as Model<Api>[];
}

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

function pickLatestMatchingModel(provider: DefaultProvider, match: RegExp): Model<Api> | null {
  const models: Model<Api>[] = getProviderModels(provider);
  const candidates = models.filter((m) => match.test(m.id));
  candidates.sort((a, b) => {
    const recency = modelRecencyScore(b.id) - modelRecencyScore(a.id);
    if (recency !== 0) return recency;
    return a.id.localeCompare(b.id);
  });
  return candidates[0] ?? null;
}

function pickPreferredOpenAiModel(provider: "openai-codex" | "openai"): Model<Api> | null {
  const models: Model<Api>[] = getProviderModels(provider);
  const bestGpt = models
    .filter((m) => isOpenAiGeneralGptModelId(m.id))
    .sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0];
  const bestCodex = models
    .filter((m) => isOpenAiCodexModelId(m.id))
    .sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0];

  if (bestGpt && bestCodex) {
    return shouldPreferOpenAiGeneralModel(bestGpt.id, bestCodex.id) ? bestGpt : bestCodex;
  }

  if (bestGpt) return bestGpt;
  if (bestCodex) return bestCodex;

  return models.slice().sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0] ?? null;
}

export function pickDefaultModel(
  availableProviders: string[],
  customDefaultModel?: Model<Api> | null,
): Model<Api> {
  // OpenAI special-case:
  // GPT-5.5 is the preferred default when an OpenAI-backed provider is available.
  // Prefer the newest general GPT-5 model when it is at least as new as Codex,
  // while keeping Codex as fallback.
  for (const provider of ["openai", "openai-codex"] as const) {
    if (!availableProviders.includes(provider)) continue;
    const model = pickPreferredOpenAiModel(provider);
    if (model) return model;
  }

  // Anthropic special-case:
  // Prefer Sonnet when its major/minor version is >= Opus (e.g. Sonnet 4-6 over Opus 4-6).
  // Otherwise prefer Opus.
  if (availableProviders.includes("anthropic")) {
    const models: Model<Api>[] = getProviderModels("anthropic");
    const opus = models
      .filter((m) => m.id.startsWith("claude-opus-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];
    const sonnet = models
      .filter((m) => m.id.startsWith("claude-sonnet-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];

    if (opus && sonnet) {
      return parseMajorMinor(sonnet.id) >= parseMajorMinor(opus.id) ? sonnet : opus;
    }

    if (opus) return opus;
    if (sonnet) return sonnet;
  }

  // Other providers: pattern-based rules
  for (const rule of DEFAULT_MODEL_RULES) {
    if (!availableProviders.includes(rule.provider)) continue;
    const m = pickLatestMatchingModel(rule.provider, rule.match);
    if (m) return m;
  }

  if (customDefaultModel) {
    return customDefaultModel;
  }

  // Absolute fallback: keep this resilient across pi-ai version bumps
  return getModel("openai", "gpt-5.5");
}
