export const DEFAULT_COMPACTION_RESERVE_TOKENS = 16_384;
export const DEFAULT_COMPACTION_KEEP_RECENT_TOKENS = 20_000;

const FALLBACK_CONTEXT_WINDOW = 200_000;

const LARGE_CONTEXT_WINDOW_TOKENS = 128_000;
const XL_CONTEXT_WINDOW_TOKENS = 200_000;
const LARGE_CONTEXT_HARD_RATIO = 0.88;
const XL_CONTEXT_HARD_RATIO = 0.85;

const SOFT_MARGIN_RATIO = 0.05;
const SOFT_FLOOR_RATIO = 0.7;
const MIN_SOFT_MARGIN_TOKENS = 2_048;

function normalizeContextWindow(contextWindow: number): number {
  if (!Number.isFinite(contextWindow)) return FALLBACK_CONTEXT_WINDOW;
  const rounded = Math.floor(contextWindow);
  return rounded > 0 ? rounded : FALLBACK_CONTEXT_WINDOW;
}

export function effectiveReserveTokens(contextWindow: number): number {
  const normalizedWindow = normalizeContextWindow(contextWindow);

  return Math.min(
    DEFAULT_COMPACTION_RESERVE_TOKENS,
    Math.max(256, Math.floor(normalizedWindow / 2)),
  );
}

export function effectiveKeepRecentTokens(contextWindow: number, reserveTokens: number): number {
  const normalizedWindow = normalizeContextWindow(contextWindow);
  const normalizedReserve = Math.max(0, Math.floor(reserveTokens));

  return Math.min(
    DEFAULT_COMPACTION_KEEP_RECENT_TOKENS,
    Math.max(0, normalizedWindow - normalizedReserve),
  );
}

export interface CompactionThresholds {
  contextWindow: number;
  reserveTokens: number;
  hardTriggerTokens: number;
  softWarningTokens: number;
}

function getQualityHardCap(contextWindow: number): number {
  if (contextWindow >= XL_CONTEXT_WINDOW_TOKENS) {
    return Math.floor(contextWindow * XL_CONTEXT_HARD_RATIO);
  }

  if (contextWindow >= LARGE_CONTEXT_WINDOW_TOKENS) {
    return Math.floor(contextWindow * LARGE_CONTEXT_HARD_RATIO);
  }

  return contextWindow;
}

export function getCompactionThresholds(contextWindow: number): CompactionThresholds {
  const normalizedWindow = normalizeContextWindow(contextWindow);
  const reserveTokens = effectiveReserveTokens(normalizedWindow);

  const piHardTrigger = Math.max(0, normalizedWindow - reserveTokens);
  const qualityHardCap = getQualityHardCap(normalizedWindow);
  const hardTriggerTokens = Math.max(0, Math.min(piHardTrigger, qualityHardCap));

  const softMargin = Math.max(
    MIN_SOFT_MARGIN_TOKENS,
    Math.floor(normalizedWindow * SOFT_MARGIN_RATIO),
  );
  const softFloor = Math.floor(hardTriggerTokens * SOFT_FLOOR_RATIO);
  const softWarningTokens = Math.max(0, Math.max(softFloor, hardTriggerTokens - softMargin));

  return {
    contextWindow: normalizedWindow,
    reserveTokens,
    hardTriggerTokens,
    softWarningTokens,
  };
}
