import { DEFAULT_TOOL_RESULT_SHAPING } from "../messages/tool-result-shaping.js";
import {
  DEFAULT_TOOL_OUTPUT_MAX_BYTES,
  DEFAULT_TOOL_OUTPUT_MAX_LINES,
  type ToolOutputTruncationLimits,
} from "../tools/output-truncation.js";

const FULL_BUDGET_CONTEXT_WINDOW = 128_000;

const MIN_TOOL_OUTPUT_MAX_BYTES = 8 * 1024;
const MIN_TOOL_OUTPUT_MAX_LINES = 200;
const MIN_RECENT_TOOL_RESULTS_TO_KEEP = 2;

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

export function effectiveRecentToolResultsToKeep(contextWindow?: number): number {
  const ratio = windowRatio(contextWindow);

  return Math.max(
    MIN_RECENT_TOOL_RESULTS_TO_KEEP,
    Math.floor(DEFAULT_TOOL_RESULT_SHAPING.recentToolResultsToKeep * ratio),
  );
}
