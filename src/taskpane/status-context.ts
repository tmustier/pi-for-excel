import { getCompactionThresholds } from "../compaction/defaults.js";
import { t } from "../language/index.js";

export const STATUS_CONTEXT_DESC_ATTR = "data-ctx-desc";
export const STATUS_CONTEXT_TOKENS_ATTR = "data-ctx-tokens";
export const STATUS_CONTEXT_WARNING_ATTR = "data-ctx-warn";
export const STATUS_CONTEXT_WARNING_SEVERITY_ATTR = "data-ctx-severity";

export function getStatusContextTooltipDescription(): string {
  return t("status.context.tooltip");
}
export function getStatusContextPopoverFallbackDescription(): string {
  return t("status.context.popoverFallback");
}

export type StatusContextWarningSeverity = "yellow" | "red";

type StatusContextColorClass = "" | "pi-status-ctx--yellow" | "pi-status-ctx--red";

export interface StatusContextWarning {
  text: string;
  severity: StatusContextWarningSeverity;
  actionText: string;
}

export interface StatusContextHealth {
  colorClass: StatusContextColorClass;
  warning: StatusContextWarning | null;
}

function getStrongActionText(): string {
  return t("status.context.strongAction");
}
function getSoftActionText(): string {
  return t("status.context.softAction");
}

/** Yellow starts at this share of the soft warning budget. */
const YELLOW_SHARE_OF_SOFT_WARNING = 0.75;

/**
 * Colours follow the compaction budgets for this context window, so red means
 * close to compaction (past the soft warning budget) on every model, and
 * yellow means well on the way there.
 */
export function getStatusContextHealth(totalTokens: number, contextWindow: number): StatusContextHealth {
  const pct = contextWindow > 0 ? Math.round((totalTokens / contextWindow) * 100) : 0;
  const { softWarningTokens } = getCompactionThresholds(contextWindow);
  if (totalTokens > contextWindow) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: t("status.context.full"),
        severity: "red",
        actionText: getStrongActionText(),
      },
    };
  }

  if (totalTokens > softWarningTokens) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: t("status.context.severe", { pct }),
        severity: "red",
        actionText: getStrongActionText(),
      },
    };
  }

  if (totalTokens > softWarningTokens * YELLOW_SHARE_OF_SOFT_WARNING) {
    return {
      colorClass: "pi-status-ctx--yellow",
      warning: {
        text: t("status.context.warning", { pct }),
        severity: "yellow",
        actionText: getSoftActionText(),
      },
    };
  }

  return {
    colorClass: "",
    warning: null,
  };
}

export function parseStatusContextWarningSeverity(
  value: string | null,
): StatusContextWarningSeverity {
  return value === "red" ? "red" : "yellow";
}
