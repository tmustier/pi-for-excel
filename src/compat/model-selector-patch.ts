/**
 * Model selector patch.
 *
 * Pi-web-ui's `ModelSelector` dialog currently does not expose hooks for filtering
 * providers by “configured API keys” and for presenting "featured" models.
 *
 * We patch its private `getFilteredModels()` method at runtime.
 * Keep all monkey-patch logic isolated here.
 */

import type { Api, Model } from "@mariozechner/pi-ai";
import { ModelSelector } from "@mariozechner/pi-web-ui/dist/dialogs/ModelSelector.js";

import {
  compareModels as compareModelRefs,
  familyPriority,
  modelRecencyScore,
  parseMajorMinor,
  providerPriority,
} from "../models/model-ordering.js";

let _activeProviders: Set<string> | null = null;

export function setActiveProviders(providers: Set<string>) {
  _activeProviders = providers;
}

type ModelSelectorItem = {
  provider: string;
  id: string;
  model: Model<Api>;
};

type ModelSelectorPrivate = {
  // Private method on ModelSelector — patched at runtime.
  getFilteredModels: (this: ModelSelector) => ModelSelectorItem[];
};

let _installed = false;

export function installModelSelectorPatch(): void {
  if (_installed) return;
  _installed = true;

  const modelSelectorProto = ModelSelector.prototype as unknown as Partial<ModelSelectorPrivate>;
  const orig = modelSelectorProto.getFilteredModels;

  if (typeof orig !== "function") {
    console.warn(
      "[pi] ModelSelector.getFilteredModels() is missing; provider filtering is disabled.",
    );
    return;
  }

  modelSelectorProto.getFilteredModels = function (this: ModelSelector): ModelSelectorItem[] {
    // Upstream bug: search tokens aren't lowercased while the target text is,
    // so matching is accidentally case-sensitive.  Temporarily lowercase the
    // query so the original filter compares apples-to-apples.
    const savedQuery = this.searchQuery;
    this.searchQuery = savedQuery.toLowerCase();
    const all = orig.call(this);
    this.searchQuery = savedQuery;

    let filtered = all;

    const active = _activeProviders;
    if (active && active.size > 0) {
      filtered = all.filter((m) => active.has(m.provider));
    }

    const currentModel = this.currentModel;

    const isCurrent = (x: ModelSelectorItem): boolean =>
      Boolean(
        currentModel &&
          x.model.id === currentModel.id &&
          x.model.provider === currentModel.provider,
      );

    const keyOf = (x: { provider: string; id: string }): string => `${x.provider}:${x.id}`;

    const compareModels = (a: ModelSelectorItem, b: ModelSelectorItem): number =>
      compareModelRefs(a, b);

    // "Latest for each" behavior:
    // - keep current model at the very top
    // - then show "featured" models (latest per provider, pattern-based)
    //   - Anthropic: latest Sonnet if its version > latest Opus, then latest Opus
    //   - OpenAI Codex: latest gpt-5.x-codex, then latest gpt-5.x
    //   - Google: latest gemini-*-pro*
    // - then show the remaining models, sorted deterministically

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

    const featured: ModelSelectorItem[] = [];
    for (const provider of providers) {
      const models = byProvider.get(provider);
      if (!models || models.length === 0) continue;

      // Provider-specific "latest" rules
      if (provider === "anthropic") {
        const bestOpus = pickBestByRecency(models, (m) => m.id.startsWith("claude-opus-"));
        const bestSonnet = pickBestByRecency(models, (m) => m.id.startsWith("claude-sonnet-"));

        if (bestOpus && bestSonnet) {
          const opusVer = parseMajorMinor(bestOpus.id);
          const sonnetVer = parseMajorMinor(bestSonnet.id);
          if (sonnetVer > opusVer) {
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

        const best = pickBest(models);
        if (best) featured.push(best);
        continue;
      }

      if (provider === "openai-codex") {
        const bestCodex = pickBestByRecency(models, (m) => /^gpt-5\.(\d+)-codex$/.test(m.id));
        const bestGpt5 = pickBestByRecency(
          models,
          (m) => /^gpt-5\./.test(m.id) && !/codex/.test(m.id),
        );

        if (bestCodex) featured.push(bestCodex);
        if (bestGpt5) featured.push(bestGpt5);
        if (bestCodex || bestGpt5) continue;

        const best = pickBest(models);
        if (best) featured.push(best);
        continue;
      }

      if (provider === "google" || provider === "google-gemini-cli" || provider === "google-antigravity") {
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
    remaining.sort(compareModels);
    for (const m of remaining) push(m);

    return out;
  };
}
