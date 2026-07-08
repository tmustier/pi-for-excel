/**
 * "Featured models" ordering for the model selector.
 *
 * Moved from the retired src/compat/model-selector-patch.ts (which used to
 * monkey-patch pi-web-ui's ModelSelector). Pure function over selector items:
 *
 * - current model at the very top
 * - then "featured" models (latest per provider, pattern-based)
 *   - Anthropic: latest Fable first (post-4.x flagship family), then latest
 *     Sonnet if its version >= latest Opus, then latest Opus
 *   - OpenAI (API + ChatGPT): latest GPT-5 when its version >= latest Codex,
 *     then latest Codex
 *   - Google API-key: latest gemini-*-pro*
 *   - Google OAuth (Gemini CLI / Antigravity): prefer stable Gemini before previews
 * - then the remaining models, sorted deterministically
 */

import type { Api, Model } from "@earendil-works/pi-ai/compat";

import {
  compareModels as compareModelRefs,
  compareOpenAiModelIds,
  familyPriority,
  isOpenAiCodexModelId,
  isOpenAiGeneralGptModelId,
  modelRecencyScore,
  parseMajorMinor,
  providerPriority,
  shouldPreferOpenAiGeneralModel,
} from "./model-ordering.js";

export interface ModelSelectorItem {
  provider: string;
  id: string;
  model: Model<Api>;
}

export function orderModelsForSelector(
  filtered: readonly ModelSelectorItem[],
  currentModel: Model<Api> | null,
): ModelSelectorItem[] {
  const isCurrent = (x: ModelSelectorItem): boolean =>
    Boolean(
      currentModel &&
        x.model.id === currentModel.id &&
        x.model.provider === currentModel.provider,
    );

  const keyOf = (x: { provider: string; id: string }): string => `${x.provider}:${x.id}`;

  const byProvider = new Map<string, ModelSelectorItem[]>();
  for (const m of filtered) {
    const list = byProvider.get(m.provider);
    if (list) list.push(m);
    else byProvider.set(m.provider, [m]);
  }

  const providers = Array.from(byProvider.keys()).sort((a, b) => {
    const aProv = providerPriority(a);
    const bProv = providerPriority(b);
    if (aProv !== bProv) return aProv - bProv;
    return a.localeCompare(b);
  });

  const pickBest = (
    models: ModelSelectorItem[],
    filter?: (m: ModelSelectorItem) => boolean,
  ): ModelSelectorItem | null => {
    const list = filter ? models.filter(filter) : models;
    if (!list.length) return null;
    return (
      list
        .slice()
        .sort((a, b) => {
          const aFam = familyPriority(a.provider, a.id);
          const bFam = familyPriority(b.provider, b.id);
          if (aFam !== bFam) return aFam - bFam;
          const aRec = modelRecencyScore(a.id);
          const bRec = modelRecencyScore(b.id);
          if (aRec !== bRec) return bRec - aRec;
          return a.id.localeCompare(b.id);
        })[0] ?? null
    );
  };

  const pickBestByRecency = (
    models: ModelSelectorItem[],
    filter: (m: ModelSelectorItem) => boolean,
  ): ModelSelectorItem | null => {
    const list = models.filter(filter);
    if (!list.length) return null;
    return (
      list
        .slice()
        .sort((a, b) => {
          const aRec = modelRecencyScore(a.id);
          const bRec = modelRecencyScore(b.id);
          if (aRec !== bRec) return bRec - aRec;
          return a.id.localeCompare(b.id);
        })[0] ?? null
    );
  };

  const pickBestOpenAi = (
    models: ModelSelectorItem[],
    filter: (m: ModelSelectorItem) => boolean,
  ): ModelSelectorItem | null => {
    const list = models.filter(filter);
    if (!list.length) return null;
    return list.slice().sort((a, b) => compareOpenAiModelIds(a.id, b.id))[0] ?? null;
  };

  const featured: ModelSelectorItem[] = [];
  for (const provider of providers) {
    const models = byProvider.get(provider);
    if (!models || models.length === 0) continue;

    if (provider === "anthropic") {
      const bestFable = pickBestByRecency(models, (m) => m.id.startsWith("claude-fable-"));
      if (bestFable) {
        featured.push(bestFable);
      }

      const bestOpus = pickBestByRecency(models, (m) => m.id.startsWith("claude-opus-"));
      const bestSonnet = pickBestByRecency(models, (m) => m.id.startsWith("claude-sonnet-"));

      if (bestOpus && bestSonnet) {
        const opusVer = parseMajorMinor(bestOpus.id);
        const sonnetVer = parseMajorMinor(bestSonnet.id);
        if (sonnetVer >= opusVer) {
          featured.push(bestSonnet, bestOpus);
          continue;
        }
        featured.push(bestOpus);
        continue;
      }

      if (bestOpus) {
        featured.push(bestOpus);
        continue;
      }

      if (bestSonnet) {
        featured.push(bestSonnet);
        continue;
      }

      if (bestFable) {
        // Fable-only provider list — already featured above.
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    if (provider === "openai-codex" || provider === "openai") {
      const bestCodex = pickBestOpenAi(models, (m) => isOpenAiCodexModelId(m.id));
      const bestGpt5 = pickBestOpenAi(models, (m) => isOpenAiGeneralGptModelId(m.id));

      if (bestCodex && bestGpt5) {
        if (shouldPreferOpenAiGeneralModel(bestGpt5.id, bestCodex.id)) {
          featured.push(bestGpt5, bestCodex);
          continue;
        }
        featured.push(bestCodex, bestGpt5);
        continue;
      }

      if (bestGpt5) {
        featured.push(bestGpt5);
        continue;
      }

      if (bestCodex) {
        featured.push(bestCodex);
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    if (provider === "google") {
      const bestPro = pickBestByRecency(models, (m) => /^gemini-.*-pro/i.test(m.id));
      if (bestPro) {
        featured.push(bestPro);
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    if (provider === "google-gemini-cli" || provider === "google-antigravity") {
      const bestStablePro = pickBestByRecency(
        models,
        (m) => /^gemini-(?!.*preview).*?-pro/i.test(m.id),
      );
      if (bestStablePro) {
        featured.push(bestStablePro);
        continue;
      }

      const bestStable = pickBestByRecency(models, (m) => /^gemini-(?!.*preview)/i.test(m.id));
      if (bestStable) {
        featured.push(bestStable);
        continue;
      }

      const bestPro = pickBestByRecency(models, (m) => /^gemini-.*-pro/i.test(m.id));
      if (bestPro) {
        featured.push(bestPro);
        continue;
      }

      const best = pickBest(models);
      if (best) featured.push(best);
      continue;
    }

    // Generic fallback
    const best = pickBest(models);
    if (best) featured.push(best);
  }

  const out: ModelSelectorItem[] = [];
  const used = new Set<string>();

  const push = (m: ModelSelectorItem) => {
    const k = keyOf(m);
    if (used.has(k)) return;
    used.add(k);
    out.push(m);
  };

  // Current model first (if it's in the filtered list)
  for (const m of filtered) {
    if (isCurrent(m)) push(m);
  }

  // Then latest-for-each-provider
  for (const m of featured) {
    push(m);
  }

  // Then the remaining models
  const remaining = filtered.filter((m) => !used.has(keyOf(m)));
  remaining.sort((a, b) => compareModelRefs(a, b));
  for (const m of remaining) push(m);

  return out;
}

/**
 * Score a query against a text using subsequence matching.
 * All query characters must appear in order in the text.
 * Higher score = tighter match (fewer gaps between matched characters).
 * Returns 0 if no match. Callers must lowercase both sides.
 */
export function subsequenceScore(query: string, text: string): number {
  let qi = 0;
  let ti = 0;
  let gaps = 0;
  let lastMatchIndex = -1;

  while (qi < query.length && ti < text.length) {
    if (query[qi] === text[ti]) {
      if (lastMatchIndex >= 0) {
        gaps += ti - lastMatchIndex - 1;
      }
      lastMatchIndex = ti;
      qi += 1;
    }
    ti += 1;
  }

  if (qi < query.length) return 0;

  return query.length / (query.length + gaps);
}
