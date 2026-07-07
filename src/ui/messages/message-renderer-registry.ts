/**
 * Registry of custom message renderers by role.
 *
 * First-party replacement for pi-web-ui's message-renderer-registry
 * (docs/ui-ownership.md). <message-list> consults this registry before
 * falling back to the built-in user/assistant renderers.
 */

import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { TemplateResult } from "lit";

export interface MessageRenderer<Message extends AgentMessage> {
  render(message: Message): TemplateResult;
}

type MessageRole = AgentMessage["role"];
type MessageForRole<Role extends MessageRole> = Extract<AgentMessage, { role: Role }>;

const messageRenderers = new Map<MessageRole, (message: AgentMessage) => TemplateResult>();

export function registerMessageRenderer<Role extends MessageRole>(
  role: Role,
  renderer: MessageRenderer<MessageForRole<Role>>,
): void {
  messageRenderers.set(role, (message) =>
    // Safe: the registry dispatches by `message.role`, so a renderer stored
    // under `role` only ever receives messages whose role matches it.
    renderer.render(message as MessageForRole<Role>),
  );
}

/** Render a message via its registered custom renderer, if any. */
export function renderMessage(message: AgentMessage): TemplateResult | undefined {
  return messageRenderers.get(message.role)?.(message);
}
