/**
 * Auto-compaction.
 *
 * Hard trigger:
 *   projectedContextTokens > hardTriggerTokens
 *
 * where hardTriggerTokens is derived from model context window and compaction
 * defaults (see `getCompactionThresholds`).
 */

import type { Agent } from "@earendil-works/pi-agent-core";

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

  // Nothing to summarize / no room to improve.
  if (agent.state.messages.length < 4) return false;

  // Delegate compaction execution to the caller so UI can show indicators and
  // to ensure we respect any ordered action queue.
  await runCompact();
  return true;
}
