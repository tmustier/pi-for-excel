import type { Agent, AgentMessage, AgentState } from "@earendil-works/pi-agent-core";
import { isContextOverflow, type AssistantMessage } from "@earendil-works/pi-ai";

import { t } from "../language/index.js";
import { announce } from "../ui/announcer.js";

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

function latestPrompt(messages: readonly AgentMessage[]): AgentMessage | null {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message?.role === "user") return message;
  }
  return null;
}

const recoveredPrompts = new WeakMap<Agent, AgentMessage | null>();

export async function recoverFromContextOverflow(args: {
  agent: Agent;
  runCompact: () => Promise<void>;
}): Promise<boolean> {
  const { agent, runCompact } = args;

  const failure = findTrailingContextOverflowError(agent.state);
  if (!failure) return false;
  if (recoveredPrompts.get(agent) === latestPrompt(agent.state.messages)) return false;

  announce(t("announce.retrying"));

  agent.state.messages = agent.state.messages.slice(0, -1);

  const beforeCompact = agent.state.messages;
  let compacted = false;
  try {
    await runCompact();
    compacted = agent.state.messages !== beforeCompact;
  } catch (err) {
    console.warn("[pi] Overflow recovery compaction failed:", err);
  }
  recoveredPrompts.set(agent, latestPrompt(agent.state.messages));
  if (!compacted) {
    agent.state.messages = [...agent.state.messages, failure];
    return false;
  }

  const messages = agent.state.messages;
  const last = messages[messages.length - 1];
  if (!last || last.role === "assistant") return false;

  await agent.continue();
  return true;
}
