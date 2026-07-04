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

export function getStatusContextHealth(pct: number): StatusContextHealth {
  if (pct > 100) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: t("status.context.full"),
        severity: "red",
        actionText: getStrongActionText(),
      },
    };
  }

  if (pct > 60) {
    return {
      colorClass: "pi-status-ctx--red",
      warning: {
        text: t("status.context.severe", { pct }),
        severity: "red",
        actionText: getStrongActionText(),
      },
    };
  }

  if (pct > 40) {
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
