/**
 * Context-window-scaled budgets (#566).
 *
 * The fixed defaults (50KB / 2,000 lines per tool output, 6 verbatim recent
 * tool results) are tuned for 128k-200k context windows. On small-window
 * models (e.g. 65k custom gateways) a handful of full-size tool results can
 * fill the entire window, so budgets scale linearly below a 128k baseline.
 *
 * Windows >= 128k keep the existing defaults unchanged.
 */

import { DEFAULT_TOOL_RESULT_SHAPING } from "../messages/tool-result-shaping.js";
import {
  DEFAULT_TOOL_OUTPUT_MAX_BYTES,
  DEFAULT_TOOL_OUTPUT_MAX_LINES,
  type ToolOutputTruncationLimits,
} from "../tools/output-truncation.js";

/** Window size at (and above) which the full default budgets apply. */
const FULL_BUDGET_CONTEXT_WINDOW = 128_000;

const MIN_TOOL_OUTPUT_MAX_BYTES = 8 * 1024;
const MIN_TOOL_OUTPUT_MAX_LINES = 200;
const MIN_RECENT_TOOL_RESULTS_TO_KEEP = 2;

/**
 * Scale factor in [0, 1]. Unknown/invalid windows are treated as large so
 * behavior matches the pre-scaling defaults.
 */
function windowRatio(contextWindow: number | undefined): number {
  if (
    typeof contextWindow !== "number" ||
    !Number.isFinite(contextWindow) ||
    contextWindow <= 0
  ) {
    return 1;
  }

  return Math.min(1, contextWindow / FULL_BUDGET_CONTEXT_WINDOW);
}

/**
 * Execution-time tool output truncation caps for the active model.
 *
 * At 65k this yields ~25KB (~6.5k tokens, ~10% of the window) per tool result
 * instead of the fixed 50KB (~20% of a 65k window).
 */
export function effectiveToolOutputLimits(contextWindow?: number): ToolOutputTruncationLimits {
  const ratio = windowRatio(contextWindow);

  return {
    maxBytes: Math.max(
      MIN_TOOL_OUTPUT_MAX_BYTES,
      Math.floor(DEFAULT_TOOL_OUTPUT_MAX_BYTES * ratio),
    ),
    maxLines: Math.max(
      MIN_TOOL_OUTPUT_MAX_LINES,
      Math.floor(DEFAULT_TOOL_OUTPUT_MAX_LINES * ratio),
    ),
  };
}

/**
 * How many recent tool results history shaping keeps verbatim
 * (see `shapeToolResultsForLlm`). 6 at >=128k, 3 at 65k, floor of 2.
 */
export function effectiveRecentToolResultsToKeep(contextWindow?: number): number {
  const ratio = windowRatio(contextWindow);

  return Math.max(
    MIN_RECENT_TOOL_RESULTS_TO_KEEP,
    Math.floor(DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep * ratio),
  );
}
