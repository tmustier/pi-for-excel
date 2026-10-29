import type { Agent, AgentLoopTurnUpdate } from "@earendil-works/pi-agent-core";

import { estimateContextTokens, estimateTextTokens } from "../utils/context-tokens.js";

import { getCompactionThresholds } from "./defaults.js";

export function shouldAutoCompactForProjectedTokens(args: {
  projectedTokens: number;
  contextWindow: number;
}): boolean {
  const { projectedTokens, contextWindow } = args;
  const { hardTriggerTokens } = getCompactionThresholds(contextWindow);
  return projectedTokens > hardTriggerTokens;
}

export async function maybeAutoCompactBeforePrompt(args: {
  agent: Agent;
  nextUserText: string;
  enabled: boolean;
  runCompact: () => Promise<void>;
}): Promise<boolean> {
  const { agent, nextUserText, enabled, runCompact } = args;

  if (!enabled) return false;
  if (agent.state.isStreaming) return false;

  const model = agent.state.model;
  if (!model) return false;

  const contextWindow = model.contextWindow || 200000;

  const { totalTokens } = estimateContextTokens(agent.state);
  const projectedTokens = totalTokens + estimateTextTokens(nextUserText);

  if (!shouldAutoCompactForProjectedTokens({ projectedTokens, contextWindow })) {
    return false;
  }

  if (agent.state.messages.length < 4) return false;

  await runCompact();
  return true;
}

export async function maybeAutoCompactBeforeContinuation(args: {
  agent: Agent;
  enabled: boolean;
  runCompact: () => Promise<void>;
}): Promise<AgentLoopTurnUpdate | undefined> {
  const { agent, enabled, runCompact } = args;

  if (!enabled) return undefined;

  const messages = agent.state.messages;
  const last = messages[messages.length - 1];
  if (!last || last.role !== "toolResult") return undefined;

  const model = agent.state.model;
  if (!model) return undefined;

  const contextWindow = model.contextWindow || 200000;

  const { totalTokens } = estimateContextTokens(agent.state);
  if (!shouldAutoCompactForProjectedTokens({ projectedTokens: totalTokens, contextWindow })) {
    return undefined;
  }

  if (messages.length < 4) return undefined;

  const before = agent.state.messages;
  await runCompact();
  if (agent.state.messages === before) return undefined;

  return {
    context: {
      messages: agent.state.messages.slice(),
      tools: agent.state.tools.slice(),
    },
  };
}
