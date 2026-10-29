import type { AssistantMessage, Message } from "@earendil-works/pi-ai";

export interface InterruptedResponse {
  role: "interruptedResponse";
  message: AssistantMessage;
  timestamp: number;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    interruptedResponse: InterruptedResponse;
  }
}

export function interruptedResponseToLlm(partial: InterruptedResponse): Message[] {
  const { errorMessage: _error, ...message } = partial.message;
  return [
    { ...message, content: message.content.filter((part) => part.type === "text"), stopReason: "stop" },
    { role: "user", content: "The previous response was interrupted by a provider failure. Continue from the partial answer above without repeating it. Do not repeat completed tool actions.", timestamp: partial.timestamp },
  ];
}
