import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import {
  getSupportedThinkingLevels,
  type Api,
  type Model,
} from "@earendil-works/pi-ai";

export function getThinkingLevelsForModel(model: Model<Api> | null): ThinkingLevel[] {
  if (!model || !model.reasoning) return ["off"];
  return getSupportedThinkingLevels(model);
}
