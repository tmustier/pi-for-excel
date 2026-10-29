import type { Agent } from "@earendil-works/pi-agent-core";
import {
  normalizeContext,
  type Api,
  type AssistantMessage,
  type Context,
  type Model,
  type SimpleStreamOptions,
} from "@earendil-works/pi-ai";

import { t } from "../language/index.js";

export class MissingApiKeyError extends Error {
  readonly provider: string;

  constructor(provider: string) {
    super(t("runtime.error.noApiKey", { provider }));
    this.name = "MissingApiKeyError";
    this.provider = provider;
  }
}

export async function completeWithAgentConfig(
  agent: Agent,
  model: Model<Api>,
  context: Context,
  options: Omit<SimpleStreamOptions, "apiKey"> = {},
): Promise<AssistantMessage> {
  const apiKey = await agent.getApiKey?.(model.provider);
  if (!apiKey) throw new MissingApiKeyError(model.provider);

  const stream = await agent.streamFunction(model, normalizeContext(context), {
    ...(agent.sessionId !== undefined ? { sessionId: agent.sessionId } : {}),
    ...options,
    apiKey,
  });
  return stream.result();
}
