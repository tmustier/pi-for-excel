/**
 * Runtime model reconciliation after provider configuration changes.
 *
 * Fixes #553: on a fresh install the first runtime is created before any
 * provider is connected, so it holds the absolute-fallback model
 * (`openai/gpt-5.6-sol` — the API provider). When the user then logs in with
 * ChatGPT (`openai-codex`), the runtime kept pointing at the unusable
 * `openai` provider and pi popped an "enter API key" prompt for the wrong
 * provider — charging the pasted API key instead of the subscription.
 *
 * A runtime whose model belongs to a provider with no configured credentials
 * cannot complete any request, so swapping it to the (recomputed) default
 * model is strictly an improvement. Runtimes whose provider is still
 * configured are never touched.
 */

import type { Api, Model } from "@earendil-works/pi-ai";

export interface RuntimeModelSwap {
  model: Model<Api>;
  thinkingLevel: "high" | "off";
}

export function resolveRuntimeModelSwap(opts: {
  /** Only the provider of the runtime's current model matters here. */
  currentModel: { provider: string };
  availableProviders: readonly string[];
  defaultModel: Model<Api>;
  /**
   * True when the runtime is doing any work — streaming OR processing queued
   * actions (`agent.state.isStreaming || actionQueue.isBusy()`). Working
   * sessions must never have their model swapped underneath them; callers
   * skip them and reconcile on a later providers-changed pass.
   */
  isBusy: boolean;
}): RuntimeModelSwap | null {
  const { currentModel, availableProviders, defaultModel, isBusy } = opts;

  // Never yank the model out from under a working session (streaming or
  // queue-busy — e.g. /compact, auto-compaction, queued prompts).
  if (isBusy) return null;

  // No providers configured — nothing usable to swap to.
  if (availableProviders.length === 0) return null;

  // Current provider still has credentials — leave the session alone.
  if (availableProviders.includes(currentModel.provider)) return null;

  // Only swap onto a model whose provider is actually usable, otherwise we
  // would just trade one wrong API-key prompt for another.
  if (!availableProviders.includes(defaultModel.provider)) return null;

  return {
    model: defaultModel,
    // Mirror runtime-creation semantics (init.ts): reasoning models default
    // to "high", non-reasoning models must be "off".
    thinkingLevel: defaultModel.reasoning ? "high" : "off",
  };
}
