import type { AgentEvent as LegacyEvent, AgentState } from "@earendil-works/pi-agent-core";
import type { AssistantMessage, AssistantMessageEvent, Tool } from "@earendil-works/pi-ai";
import type { AgentEvent } from "@earendil-works/pi-durable";
import type { ConversationSnapshot, ConversationPresenter } from "../durable/conversation-view.js";

export type ExtensionReadState = Readonly<Pick<AgentState,
  "messages" | "model" | "thinkingLevel" | "systemPrompt" | "isStreaming" | "streamingMessage" | "pendingToolCalls" | "errorMessage">> & {
    readonly tools: readonly Tool[];
  };

export function extensionReadState(snapshot: ConversationSnapshot): ExtensionReadState {
  const latestAssistant = snapshot.streamingMessage ?? [...snapshot.messages].reverse().find((message) => message.role === "assistant");
  const errorMessage = latestAssistant?.errorMessage;
  return Object.freeze(structuredClone({
    messages: snapshot.messages, model: snapshot.model, thinkingLevel: snapshot.thinkingLevel,
    tools: snapshot.tools ?? [],
    systemPrompt: snapshot.systemPrompt ?? "", isStreaming: snapshot.busy,
    ...(snapshot.streamingMessage ? { streamingMessage: snapshot.streamingMessage } : {}),
    pendingToolCalls: snapshot.pendingToolCalls,
    ...(errorMessage ? { errorMessage } : {}),
  }));
}

/** Read projection only: never invokes a model or executes a tool. */
export function legacyEventProjection(initial?: AssistantMessage): (event: AgentEvent) => LegacyEvent[] {
  let partial = initial ? structuredClone(initial) : undefined;
  let messages: AgentState["messages"] = [];
  let turnMessages: AgentState["messages"] = [];
  return (event) => {
    switch (event.type) {
      case "snapshot": partial = event.generation?.message ? structuredClone(event.generation.message) : undefined; return [];
      case "run_start": messages = []; turnMessages = []; partial = undefined; return [{ type: "agent_start" }];
      case "run_end": return [{ type: "agent_end", messages }];
      case "turn_start": turnMessages = []; return [{ type: "turn_start" }];
      case "message_start":
        if (event.message.role === "assistant") partial = structuredClone(event.message);
        return [{ type: "message_start", message: event.message }];
      case "message_end": {
        const committed = [...(event.entry.model ?? [])];
        messages.push(...committed);
        turnMessages.push(...committed);
        const result: LegacyEvent[] = [];
        for (const message of committed) {
          if (message.role === "assistant" && message.stopReason !== "pending") {
            for (const [contentIndex, block] of message.content.entries()) {
              const end: AssistantMessageEvent = block.type === "text"
                ? { type: "text_end", contentIndex, content: block.text, partial: message }
                : block.type === "thinking"
                  ? { type: "thinking_end", contentIndex, content: block.thinking, partial: message }
                  : { type: "toolcall_end", contentIndex, toolCall: block, partial: message };
              result.push({ type: "message_update", message, assistantMessageEvent: end });
            }
            const assistantMessageEvent: AssistantMessageEvent = message.stopReason === "error" || message.stopReason === "aborted"
              ? { type: "error", reason: message.stopReason, error: message }
              : { type: "done", reason: message.stopReason, message };
            result.push({ type: "message_update", message, assistantMessageEvent });
          }
          result.push({ type: "message_end", message });
        }
        return result;
      }
      case "turn_end": {
        const message = [...turnMessages].reverse().find((item) => item.role === "assistant");
        return message ? [{ type: "turn_end", message, toolResults: turnMessages.filter((item) => item.role === "toolResult") }] : [];
      }
      case "message_update": {
        const result: LegacyEvent[] = [];
        for (const change of event.changes) {
          if (change.type === "message") {
            partial = structuredClone(change.message);
            result.push({ type: "message_update", message: structuredClone(partial), assistantMessageEvent: { type: "start", partial: structuredClone(partial) } });
            continue;
          }
          if (!partial) continue;
          let update: AssistantMessageEvent | undefined;
          if (change.type === "block" || "block" in change) {
            partial.content[change.contentIndex] = structuredClone(change.block);
            if (change.type !== "block") update = { type: change.type, contentIndex: change.contentIndex, partial };
            else update = { type: "start", partial };
          } else {
            const block = partial.content[change.contentIndex];
            if (change.type === "text_delta" && block?.type === "text") block.text += change.delta;
            if (change.type === "thinking_delta" && block?.type === "thinking") block.thinking += change.delta;
            update = { type: change.type, contentIndex: change.contentIndex, delta: change.delta, partial };
          }
          partial.usage = event.usage;
          if (update) result.push(structuredClone({ type: "message_update", message: partial, assistantMessageEvent: update }));
        }
        return result;
      }
      case "tool_execution_start": return [event];
      case "tool_execution_end": {
        const result = event.entry?.model?.find((message) => message.role === "toolResult");
        return result ? [{ type: "tool_execution_end", toolCallId: event.toolCallId, toolName: event.toolName, result, isError: result.isError }] : [];
      }
      default: return [];
    }
  };
}

/** Hosts should supply a synchronous tab-change notifier. The fallback also
 * filters old-tab events immediately and reattaches on the next tick. */
export function subscribeActiveConversation(
  getActive: () => ConversationPresenter | null,
  handler: (event: AgentEvent) => void,
  onChange?: (listener: () => void) => () => void,
): () => void {
  let active: ConversationPresenter | null = null;
  let detach = (): void => {};
  const attach = (): void => {
    const next = getActive();
    if (next === active) return;
    detach(); active = next;
    detach = next?.onEvent((event) => { if (getActive() === next) handler(event); }) ?? (() => {});
  };
  attach();
  // Agent default (2026-10-18): 25 ms fallback only for third-party hosts without a tab notifier.
  const timer = onChange ? undefined : setInterval(attach, 25);
  const stop = onChange?.(attach);
  return () => { if (timer !== undefined) clearInterval(timer); stop?.(); detach(); };
}
