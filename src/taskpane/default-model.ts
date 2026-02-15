/**
 * Default model selection for the taskpane.
 */

import { getModel, getModels, type Api, type Model } from "@mariozechner/pi-ai";

import { modelRecencyScore, parseMajorMinor } from "../models/model-ordering.js";

type DefaultProvider =
  | "openai-codex"
  | "openai"
  | "google"
  | "google-gemini-cli"
  | "google-antigravity";

type DefaultModelRule = { provider: DefaultProvider; match: RegExp };

const DEFAULT_MODEL_RULES: DefaultModelRule[] = [
  // Prefer latest GPT-5.x Codex on ChatGPT subscription (openai-codex)
  { provider: "openai-codex", match: /^gpt-5\.(\d+)-codex$/ },
  { provider: "openai-codex", match: /^gpt-5\./ },

  // API key OpenAI provider (if user connected OpenAI instead of openai-codex)
  { provider: "openai", match: /^gpt-5\.(\d+)-codex$/ },
  { provider: "openai", match: /^gpt-5\./ },

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
  const models: Model<Api>[] = getModels(provider);
  const candidates = models.filter((m) => match.test(m.id));
  candidates.sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id));
  return candidates[0] ?? null;
}

export function pickDefaultModel(availableProviders: string[]): Model<Api> {
  // Anthropic special-case:
  // Prefer Opus, except if there's a *newer-version* Sonnet, use that first.
  if (availableProviders.includes("anthropic")) {
    const models: Model<Api>[] = getModels("anthropic");
    const opus = models
      .filter((m) => m.id.startsWith("claude-opus-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];
    const sonnet = models
      .filter((m) => m.id.startsWith("claude-sonnet-"))
      .sort((a, b) => modelRecencyScore(b.id) - modelRecencyScore(a.id))[0];

    if (opus && sonnet) {
      return parseMajorMinor(sonnet.id) > parseMajorMinor(opus.id) ? sonnet : opus;
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

  // Absolute fallback: keep this resilient across pi-ai version bumps
  return getModel("anthropic", "claude-opus-4-6");
}
