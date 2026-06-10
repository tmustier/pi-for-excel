/**
 * Auto-compaction.
 *
 * Hard trigger:
 *   projectedContextTokens > hardTriggerTokens
 *
 * where hardTriggerTokens is derived from model context window and compaction
 * defaults (see `getCompactionThresholds`).
 *
 * Two trigger points share the same budgets:
 * - before a queued user prompt (`maybeAutoCompactBeforePrompt`)
 * - mid-turn, between tool-loop continuations (`maybeAutoCompactBeforeContinuation`)
 */

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

  // Nothing to summarize / no room to improve.
  if (agent.state.messages.length < 4) return false;

  // Delegate compaction execution to the caller so UI can show indicators and
  // to ensure we respect any ordered action queue.
  await runCompact();
  return true;
}

/**
 * Mid-turn compaction check, run from `Agent.prepareNextTurn` after each
 * completed tool batch. A single tool-heavy turn can overflow a small context
 * window before the next between-prompt check would ever fire (#566).
 *
 * Returns a replacement loop context when compaction rewrote the transcript,
 * so the in-flight run continues from the compacted history.
 */
export async function maybeAutoCompactBeforeContinuation(args: {
  agent: Agent;
  enabled: boolean;
  runCompact: () => Promise<void>;
}): Promise<AgentLoopTurnUpdate | undefined> {
  const { agent, enabled, runCompact } = args;

  if (!enabled) return undefined;

  const messages = agent.state.messages;
  const last = messages[messages.length - 1];
  // Only act when another continuation request is coming (tool loop). When the
  // turn ended with a plain assistant message, the run is about to stop.
  if (!last || last.role !== "toolResult") return undefined;

  const model = agent.state.model;
  if (!model) return undefined;

  const contextWindow = model.contextWindow || 200000;

  const { totalTokens } = estimateContextTokens(agent.state);
  if (!shouldAutoCompactForProjectedTokens({ projectedTokens: totalTokens, contextWindow })) {
    return undefined;
  }

  // Nothing to summarize / no room to improve.
  if (messages.length < 4) return undefined;

  // `state.messages` is replaced (new array identity) when compaction rewrites
  // history; an unchanged reference means compaction failed or was a no-op.
  const before = agent.state.messages;
  await runCompact();
  if (agent.state.messages === before) return undefined;

  return {
    context: {
      systemPrompt: agent.state.systemPrompt,
      messages: agent.state.messages.slice(),
      tools: agent.state.tools.slice(),
    },
  };
}
