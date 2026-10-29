import type { ExperimentalFeatureId } from "./flags.js";

export const PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT = "pi:experimental-feature-changed";
export const PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT = "pi:experimental-tool-config-changed";

export interface ExperimentalFeatureChangedDetail {
  featureId: ExperimentalFeatureId;
  enabled: boolean;
}

export interface ExperimentalToolConfigChangedDetail {
  configKey: string;
}

export function dispatchExperimentalFeatureChanged(
  detail: ExperimentalFeatureChangedDetail,
): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(
    new CustomEvent<ExperimentalFeatureChangedDetail>(PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT, {
      detail,
    }),
  );
}

export function dispatchExperimentalToolConfigChanged(
  detail: ExperimentalToolConfigChangedDetail,
): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(
    new CustomEvent<ExperimentalToolConfigChangedDetail>(
      PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT,
      {
        detail,
      },
    ),
  );
}
