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
    renderer.render(message as MessageForRole<Role>),
  );
}

export function renderMessage(message: AgentMessage): TemplateResult | undefined {
  return messageRenderers.get(message.role)?.(message);
}
