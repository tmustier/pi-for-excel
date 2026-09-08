/**
 * Context-overflow recovery (#566).
 *
 * When a provider rejects a request because the prompt exceeds the model's
 * context window, the run ends with a trailing assistant error message.
 * Mirroring pi-mono's agent-session behavior, we recover once per prompt:
 * drop the failed assistant message from context, compact, and retry via
 * `agent.continue()`.
 */

import type { Agent, AgentState } from "@earendil-works/pi-agent-core";
import { isContextOverflow, type AssistantMessage } from "@earendil-works/pi-ai";

/**
 * Returns the trailing assistant context-overflow error message, or null.
 *
 * Only failures from the currently active model count: after switching to a
 * larger-context model, a stale overflow error must not trigger recovery.
 */
export function findTrailingContextOverflowError(
  state: Pick<AgentState, "messages" | "model">,
): AssistantMessage | null {
  const messages = state.messages;
  const last = messages[messages.length - 1];
  if (!last || last.role !== "assistant") return null;
  if (last.stopReason !== "error") return null;

  const model = state.model;
  if (model && (last.provider !== model.provider || last.model !== model.id)) {
    return null;
  }

  return isContextOverflow(last, model?.contextWindow || undefined) ? last : null;
}

/**
 * Compact-and-retry recovery for a run that ended in a context-overflow error.
 *
 * Call once after a prompt run settles. Returns true when a retry was started
 * (and has settled — `agent.continue()` is awaited). The caller must not loop:
 * if the retry overflows again, the error stays in the transcript and the user
 * is pointed at `/compact` instead.
 */
export async function recoverFromContextOverflow(args: {
  agent: Agent;
  runCompact: () => Promise<void>;
}): Promise<boolean> {
  const { agent, runCompact } = args;

  const failure = findTrailingContextOverflowError(agent.state);
  if (!failure) return false;

  // Drop the failed assistant message so the retry context is clean.
  agent.state.messages = agent.state.messages.slice(0, -1);

  // `state.messages` is replaced (new array identity) when compaction rewrites
  // history; an unchanged reference means compaction failed or was a no-op.
  const beforeCompact = agent.state.messages;
  let compacted = false;
  try {
    await runCompact();
    compacted = agent.state.messages !== beforeCompact;
  } catch (err) {
    // Compaction errors are normally surfaced via toasts inside runCompact;
    // don't let an unexpected throw break the caller's queue processing.
    console.warn("[pi] Overflow recovery compaction failed:", err);
  }
  if (!compacted) {
    // Compaction freed nothing (no-op or failure) — restore the failure so it
    // stays visible and persisted, and let the user act on the banner hint.
    agent.state.messages = [...agent.state.messages, failure];
    return false;
  }

  const messages = agent.state.messages;
  const last = messages[messages.length - 1];
  // Continuations require a non-assistant tail (compaction keeps the original
  // tail, so this only fails in pathological cases).
  if (!last || last.role === "assistant") return false;

  await agent.continue();
  return true;
}
