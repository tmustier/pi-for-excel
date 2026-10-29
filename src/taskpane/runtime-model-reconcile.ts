import type { Api, Model } from "@earendil-works/pi-ai";

export interface RuntimeModelSwap {
  model: Model<Api>;
  thinkingLevel: "high" | "off";
}

export function resolveRuntimeModelSwap(opts: {
  currentModel: { provider: string };
  availableProviders: readonly string[];
  defaultModel: Model<Api>;
  isBusy: boolean;
}): RuntimeModelSwap | null {
  const { currentModel, availableProviders, defaultModel, isBusy } = opts;

  if (isBusy) return null;

  if (availableProviders.length === 0) return null;

  if (availableProviders.includes(currentModel.provider)) return null;

  if (!availableProviders.includes(defaultModel.provider)) return null;

  return {
    model: defaultModel,
    thinkingLevel: defaultModel.reasoning ? "high" : "off",
  };
}
