import type { ThinkingLevel } from "@earendil-works/pi-agent-core";

import { t } from "../language/index.js";

export const THINKING_LEVEL_COLORS: Record<ThinkingLevel, string> = {
  off: "#a0a0a0",
  minimal: "#767676",
  low: "#4488cc",
  medium: "#22998a",
  high: "#875f87",
  xhigh: "#8b008b",
  max: "#af005f",
};

export function getThinkingLevelLabel(level: ThinkingLevel): string {
  switch (level) {
    case "off":
      return t("status.thinking.off");
    case "minimal":
      return t("status.thinking.min");
    case "low":
      return t("status.thinking.low");
    case "medium":
      return t("status.thinking.medium");
    case "high":
      return t("status.thinking.high");
    case "xhigh":
      return t("status.thinking.xhigh");
    case "max":
      return t("status.thinking.max");
  }
}

export function getThinkingLevelHint(level: ThinkingLevel): string {
  switch (level) {
    case "off":
      return t("status.thinking.offHint");
    case "minimal":
      return t("status.thinking.minimalHint");
    case "low":
      return t("status.thinking.lowHint");
    case "medium":
      return t("status.thinking.mediumHint");
    case "high":
      return t("status.thinking.highHint");
    case "xhigh":
      return t("status.thinking.xhighHint");
    case "max":
      return t("status.thinking.maxHint");
  }
}
