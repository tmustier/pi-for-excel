import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import {
  getSupportedThinkingLevels,
  type Api,
  type Model,
} from "@earendil-works/pi-ai/compat";

/**
 * Return the exact thinking levels exposed by the bundled Pi model registry.
 *
 * Keep this provider-agnostic: model-level maps are the canonical source for
 * exceptions such as GPT-5.6's separate `xhigh` and `max` effort levels.
 */
export function getThinkingLevelsForModel(model: Model<Api> | null): ThinkingLevel[] {
  if (!model || !model.reasoning) return ["off"];
  return getSupportedThinkingLevels(model);
}
