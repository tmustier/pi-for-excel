/**
 * <message-list> — the transcript.
 *
 * First-party replacement for pi-web-ui's MessageList (docs/ui-ownership.md).
 * Custom role renderers (message-renderer-registry) take precedence; then
 * user and assistant messages render via the built-in elements. Assistant
 * responses are drawn by turn (turn-model.ts): the prose before the first
 * tool call, one <turn-group> for every tool call of the turn, then the
 * answer. The response being streamed is part of the list, so a turn's
 * group gains rows as the model asks for them.
 */

import type { AgentMessage, AgentTool } from "@earendil-works/pi-agent-core";
import type { AssistantMessage, ToolResultMessage } from "@earendil-works/pi-ai";
import { html, LitElement, type TemplateResult } from "lit";
import { property } from "lit/decorators.js";
import { repeat } from "lit/directives/repeat.js";

import { renderMessage } from "./message-renderer-registry.js";
import { buildTranscriptBlocks, type TranscriptBlock } from "./turn-model.js";
import "./messages.js";
import "./turn-group.js";

interface RenderItem {
  key: string;
  template: TemplateResult;
}

const EMPTY_PENDING = new Set<string>();

export class MessageList extends LitElement {
  @property({ type: Array }) messages: AgentMessage[] = [];
  @property({ type: Array }) tools: AgentTool[] = [];
  @property({ type: Object }) pendingToolCalls?: ReadonlySet<string>;
  /** The turn is running: tool calls without a result are still to come. */
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: String }) errorSummary: string | null = null;
  /**
   * Results pi-agent-core reported at tool_execution_end. In a parallel batch
   * each tool finishes on its own, but its toolResult message is only added
   * once the whole batch is done; these let each row finish as it happens.
   */
  @property({ type: Object }) liveToolResults?: ReadonlyMap<string, ToolResultMessage>;
  /** The response being streamed, which pi-agent-core adds to `messages` only when it ends. */
  @property({ type: Object }) streamingMessage: AssistantMessage | null = null;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  /** Tool call ids the list draws, for questions asked in their rows. */
  showsToolCall(toolCallId: string): boolean {
    return this._blocks().some((block) => block.type === "turnGroup"
      && block.group.steps.some((step) => step.type === "tool" && step.toolCall.id === toolCallId));
  }

  private _blocks(): TranscriptBlock[] {
    const results = new Map<string, ToolResultMessage>(this.liveToolResults ?? []);
    for (const message of this.messages) {
      if (message.role === "toolResult") results.set(message.toolCallId, message);
    }
    return buildTranscriptBlocks({
      messages: this.messages,
      streamingMessage: this.streamingMessage,
      tools: this.tools,
      results,
      pendingToolCalls: this.pendingToolCalls ?? EMPTY_PENDING,
      running: this.isStreaming,
    });
  }

  private _buildRenderItems(): RenderItem[] {
    const items: RenderItem[] = [];
    const lastMessage = this.streamingMessage ?? this.messages.at(-1);

    for (const block of this._blocks()) {
      if (block.type === "turnGroup") {
        items.push({ key: block.key, template: html`<turn-group .model=${block.group}></turn-group>` });
        continue;
      }

      const { message } = block;
      const customTemplate = renderMessage(message);
      if (customTemplate) {
        items.push({ key: block.key, template: customTemplate });
        continue;
      }

      if (message.role === "user" || message.role === "user-with-attachments") {
        items.push({ key: block.key, template: html`<user-message .message=${message}></user-message>` });
      } else if (message.role === "assistant") {
        const isLast = message === lastMessage;
        const tail = block.contentTo === undefined || block.contentTo >= message.content.length;
        items.push({
          key: block.key,
          template: html`
            <assistant-message
              .message=${message}
              .contentFrom=${block.contentFrom ?? 0}
              .contentTo=${block.contentTo ?? message.content.length}
              .isStreaming=${block.streaming}
              .copyable=${tail}
              .hideError=${!tail || (this.isStreaming && isLast)}
              .canRetry=${tail && !this.isStreaming && isLast}
              .errorSummary=${tail && isLast ? this.errorSummary : null}
            ></assistant-message>
          `,
        });
      }
    }

    return items;
  }

  override render() {
    const items = this._buildRenderItems();
    return html`
      <div class="pi-message-list">
        ${repeat(items, (item) => item.key, (item) => item.template)}
      </div>
    `;
  }
}

if (!customElements.get("message-list")) {
  customElements.define("message-list", MessageList);
}

declare global {
  interface HTMLElementTagNameMap {
    "message-list": MessageList;
  }
}
