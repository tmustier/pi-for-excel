/**
 * Shared compaction defaults.
 *
 * Base values mirror pi-coding-agent defaults (see
 * /opt/homebrew/lib/node_modules/@earendil-works/pi-coding-agent/docs/compaction.md).
 *
 * Slice 5 adds a quality-first cap for large context windows and a soft warning
 * threshold so we compact earlier in long/noisy sessions.
 */

export const DEFAULT_COMPACTION_RESERVE_TOKENS = 16_384;
export const DEFAULT_COMPACTION_KEEP_RECENT_TOKENS = 20_000;

const FALLBACK_CONTEXT_WINDOW = 200_000;

/**
 * Quality cap breakpoints.
 *
 * Rationale:
 * - windows >=128k are large enough that waiting until (window - reserve) can
 *   delay compaction too long for response quality.
 * - we cap hard compaction earlier for these models while preserving Pi's
 *   reserve-based behavior for smaller windows.
 */
const LARGE_CONTEXT_WINDOW_TOKENS = 128_000;
const XL_CONTEXT_WINDOW_TOKENS = 200_000;
const LARGE_CONTEXT_HARD_RATIO = 0.88;
const XL_CONTEXT_HARD_RATIO = 0.85;

/**
 * Soft warning budget before hard trigger.
 *
 * - margin: ~5% of context (minimum 2k tokens)
 * - floor: never below 70% of the hard threshold
 */
const SOFT_MARGIN_RATIO = 0.05;
const SOFT_FLOOR_RATIO = 0.7;
const MIN_SOFT_MARGIN_TOKENS = 2_048;

function normalizeContextWindow(contextWindow: number): number {
  if (!Number.isFinite(contextWindow)) return FALLBACK_CONTEXT_WINDOW;
  const rounded = Math.floor(contextWindow);
  return rounded > 0 ? rounded : FALLBACK_CONTEXT_WINDOW;
}

/**
 * Reserve token budget used to ensure we always have room for the model's response.
 *
 * Pi uses a fixed default (16,384). For smaller context windows, we clamp to
 * half the context window (and a small minimum) to avoid pathological behavior.
 */
export function effectiveReserveTokens(contextWindow: number): number {
  const normalizedWindow = normalizeContextWindow(contextWindow);

  return Math.min(
    DEFAULT_COMPACTION_RESERVE_TOKENS,
    Math.max(256, Math.floor(normalizedWindow / 2)),
  );
}

/**
 * How many tokens of the recent conversation to keep verbatim after compaction.
 *
 * Pi defaults to 20k. For smaller context windows, clamp so that the kept tail
 * fits into the prompt budget (contextWindow - reserveTokens).
 */
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
  /** Auto-compaction trigger budget. */
  hardTriggerTokens: number;
  /** Early warning budget shown in UI (before hard trigger). */
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

/**
 * Compaction budgets for the active model context window.
 *
 * hardTriggerTokens:
 * - baseline: contextWindow - reserveTokens (Pi behavior)
 * - plus a quality cap for >=128k windows (88% / 85%)
 *
 * softWarningTokens:
 * - budget where we start nudging the user to compact before quality drifts.
 */
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
