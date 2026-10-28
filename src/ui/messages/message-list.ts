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
import { html, LitElement, nothing, type TemplateResult } from "lit";
import { property } from "lit/decorators.js";
import { repeat } from "lit/directives/repeat.js";

import { registerInlineApprovalHost } from "../approval-dialog.js";
import { renderMessage } from "./message-renderer-registry.js";
import { buildTranscriptBlocks, type TranscriptBlock } from "./turn-model.js";
import "./messages.js";
import "./turn-group.js";

export class MessageList extends LitElement {
  @property({ type: Array }) messages: AgentMessage[] = [];
  @property({ type: Array }) tools: AgentTool[] = [];
  @property({ type: Object }) pendingToolCalls: ReadonlySet<string> = new Set();
  /** The turn is running: tool calls without a result are still to come. */
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: String }) errorSummary: string | null = null;
  /**
   * Results pi-agent-core reported at tool_execution_end. In a parallel batch
   * each tool finishes on its own, but its toolResult message is only added
   * once the whole batch is done; these let each row finish as it happens.
   */
  @property({ type: Object }) liveToolResults: ReadonlyMap<string, ToolResultMessage> = new Map();
  /** The response being streamed, which pi-agent-core adds to `messages` only when it ends. */
  @property({ type: Object }) streamingMessage: AssistantMessage | null = null;

  private _approvalHost: ReturnType<typeof registerInlineApprovalHost> | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
    this._approvalHost = registerInlineApprovalHost({
      claims: (id) => this.isConnected && this._blocks().some((block) => block.type === "turnGroup"
        && block.group.steps.some((step) => step.type === "tool" && step.toolCall.id === id)),
      onChange: () => this.requestUpdate(),
    });
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._approvalHost?.unregister();
    this._approvalHost = undefined;
  }

  override updated(): void {
    this._approvalHost?.refresh();
  }

  private _blocks(): TranscriptBlock[] {
    const results = new Map(this.liveToolResults);
    for (const message of this.messages) {
      if (message.role === "toolResult") results.set(message.toolCallId, message);
    }
    return buildTranscriptBlocks({
      messages: this.messages,
      streamingMessage: this.streamingMessage,
      tools: this.tools,
      results,
      pendingToolCalls: this.pendingToolCalls,
      running: this.isStreaming,
    });
  }

  override render() {
    const lastMessage = this.streamingMessage ?? this.messages.at(-1);

    const render = (block: TranscriptBlock): TemplateResult | typeof nothing => {
      if (block.type === "turnGroup") return html`<turn-group .model=${block.group}></turn-group>`;
      const custom = renderMessage(block.message);
      if (custom) return custom;
      if (block.type === "message") {
        const { message } = block;
        return message.role === "user" || message.role === "user-with-attachments"
          ? html`<user-message .message=${message}></user-message>`
          : nothing;
      }
      const { message } = block;
      const isLast = message === lastMessage;
      const tail = block.contentTo >= message.content.length;
      return html`
        <assistant-message
          .message=${message}
          .contentFrom=${block.contentFrom}
          .contentTo=${block.contentTo}
          .isStreaming=${block.streaming}
          .copyable=${tail}
          .hideError=${!tail || (this.isStreaming && isLast)}
          .canRetry=${tail && !this.isStreaming && isLast}
          .errorSummary=${tail && isLast ? this.errorSummary : null}
        ></assistant-message>
      `;
    };

    return html`
      <div class="pi-message-list">
        ${repeat(this._blocks(), (block) => block.type === "turnGroup" ? block.group.key : block.key, render)}
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
