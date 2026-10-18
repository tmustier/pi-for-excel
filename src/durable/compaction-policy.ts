import type { CompactionPolicy } from "@earendil-works/pi-durable";
import {
  effectiveKeepRecentTokens,
  effectiveReserveTokens,
  getCompactionThresholds,
} from "../compaction/defaults.js";

/** Pure policy for the actual request model, not the selected UI tab. */
export function compactionForModel(model: { readonly contextWindow: number }): Partial<CompactionPolicy> {
  const reserveTokens = effectiveReserveTokens(model.contextWindow);
  return {
    reserveTokens,
    keepRecentTokens: effectiveKeepRecentTokens(model.contextWindow, reserveTokens),
    hardTriggerTokens: getCompactionThresholds(model.contextWindow).hardTriggerTokens,
    backgroundTokens: 0,
  };
}
