export const STATUS_CONTEXT_DESC_ATTR = "data-ctx-desc";
export const STATUS_CONTEXT_TOKENS_ATTR = "data-ctx-tokens";
export const STATUS_CONTEXT_WARNING_ATTR = "data-ctx-warn";
export const STATUS_CONTEXT_WARNING_SEVERITY_ATTR = "data-ctx-severity";

export const STATUS_CONTEXT_TOOLTIP_DESCRIPTION = "How much of Pi's memory (context window) this conversation is using.";
export const STATUS_CONTEXT_POPOVER_FALLBACK_DESCRIPTION = "How much of Pi's memory this conversation is using.";

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

const STRONG_ACTION_TEXT = "Use /compact to free space or /new to start fresh.";
const SOFT_ACTION_TEXT = "Consider using /compact to free space or /new to start fresh.";

export function getStatusContextHealth(pct: number): StatusContextHealth {
  if (pct > 100) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: "Context is full — the next message will fail.",
        severity: "red",
        actionText: STRONG_ACTION_TEXT,
      },
    };
  }

  if (pct > 60) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: `Context ${pct}% full — responses may become less reliable.`,
        severity: "red",
        actionText: STRONG_ACTION_TEXT,
      },
    };
  }

  if (pct > 40) {
    return {
      colorClass: "pi-status-ctx--yellow",
      warning: {
        text: `Context ${pct}% full.`,
        severity: "yellow",
        actionText: SOFT_ACTION_TEXT,
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
