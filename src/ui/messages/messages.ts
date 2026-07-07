/**
 * <user-message>, <assistant-message>, <tool-message> — chat message elements.
 *
 * First-party replacements for pi-web-ui's Messages.js components
 * (docs/ui-ownership.md). Tag names, light DOM, and the CSS-facing structure
 * (`.user-message-container`, `.pi-assistant-body`, `.pi-assistant-aborted`,
 * `.pi-tool-card-fallback`) are preserved; Tailwind utility classes are
 * replaced with semantic `pi-*` classes styled in src/ui/theme/content/.
 *
 * Intentional divergences from upstream:
 * - No usage/cost row (theme CSS always hid it; the status bar shows context).
 * - Non-custom tool renders get the `.pi-tool-card-fallback` wrapper class
 *   directly instead of via applyMessageStyleHooks() DOM post-processing.
 */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type {
  AssistantMessage as AssistantMessageType,
  ToolCall,
  ToolResultMessage,
  UserMessage as UserMessageType,
} from "@earendil-works/pi-ai";
import { html, LitElement, nothing, type TemplateResult } from "lit";
import { customElement, property } from "lit/decorators.js";

import { t } from "../../language/index.js";
import type { UserMessageWithAttachments } from "../../messages/attachments.js";
import { renderTool } from "./tool-renderer-registry.js";
import "./attachment-tile.js";
import "./markdown-block.js";
import "./thinking-block.js";

@customElement("user-message")
export class UserMessage extends LitElement {
  @property({ type: Object }) message?: UserMessageType | UserMessageWithAttachments;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override render() {
    const message = this.message;
    if (!message) return html``;

    const content =
      typeof message.content === "string"
        ? message.content
        : message.content.find((chunk) => chunk.type === "text")?.text || "";

    const attachments =
      message.role === "user-with-attachments" && message.attachments?.length
        ? message.attachments
        : null;

    return html`
      <div class="pi-user-row">
        <div class="user-message-container">
          <markdown-block .content=${content}></markdown-block>
          ${attachments
            ? html`
                <div class="pi-user-attachments">
                  ${attachments.map(
                    (attachment) => html`<attachment-tile .attachment=${attachment}></attachment-tile>`,
                  )}
                </div>
              `
            : nothing}
        </div>
      </div>
    `;
  }
}

@customElement("assistant-message")
export class AssistantMessage extends LitElement {
  @property({ type: Object }) message?: AssistantMessageType;
  @property({ type: Array }) tools?: AgentTool[];
  @property({ type: Object }) pendingToolCalls?: ReadonlySet<string>;
  @property({ type: Boolean }) hideToolCalls = false;
  @property({ type: Object }) toolResultsById?: Map<string, ToolResultMessage>;
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: Boolean }) hidePendingToolCalls = false;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override render() {
    const message = this.message;
    if (!message) return html``;

    // Render content in the order it appears.
    const orderedParts: TemplateResult[] = [];
    for (const chunk of message.content) {
      if (chunk.type === "text" && chunk.text.trim() !== "") {
        orderedParts.push(html`<markdown-block .content=${chunk.text}></markdown-block>`);
      } else if (chunk.type === "thinking" && chunk.thinking.trim() !== "") {
        orderedParts.push(
          html`<thinking-block .content=${chunk.thinking} .isStreaming=${this.isStreaming}></thinking-block>`,
        );
      } else if (chunk.type === "toolCall") {
        if (this.hideToolCalls) continue;

        const tool = this.tools?.find((candidate) => candidate.name === chunk.name);
        const pending = this.pendingToolCalls?.has(chunk.id) ?? false;
        const result = this.toolResultsById?.get(chunk.id);

        // Skip pending tool calls when hidePendingToolCalls is set (prevents
        // duplication while <streaming-message-container> is showing them).
        if (this.hidePendingToolCalls && pending && !result) {
          continue;
        }

        // Aborted turn with no result for this call → render as aborted.
        const aborted = message.stopReason === "aborted" && !result;

        orderedParts.push(html`
          <tool-message
            .tool=${tool}
            .toolCall=${chunk}
            .result=${result}
            .pending=${pending}
            .aborted=${aborted}
            .isStreaming=${this.isStreaming}
          ></tool-message>
        `);
      }
    }

    return html`
      <div>
        ${orderedParts.length ? html`<div class="pi-assistant-body">${orderedParts}</div>` : nothing}
        ${message.stopReason === "error" && message.errorMessage
          ? html`
              <div class="pi-assistant-error">
                <strong>${t("messages.errorLabel")}</strong> ${message.errorMessage}
              </div>
            `
          : nothing}
        ${message.stopReason === "aborted"
          ? html`<span class="pi-assistant-aborted">${t("messages.aborted")}</span>`
          : nothing}
      </div>
    `;
  }
}

@customElement("tool-message")
export class ToolMessage extends LitElement {
  @property({ type: Object }) toolCall?: ToolCall;
  @property({ type: Object }) tool?: AgentTool;
  @property({ type: Object }) result?: ToolResultMessage;
  @property({ type: Boolean }) pending = false;
  @property({ type: Boolean }) aborted = false;
  @property({ type: Boolean }) isStreaming = false;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override render() {
    const toolCall = this.toolCall;
    if (!toolCall) return html``;

    const toolName = this.tool?.name || toolCall.name;

    // Aborted calls render like an errored result with no content.
    const result: ToolResultMessage | undefined = this.aborted
      ? {
          role: "toolResult",
          isError: true,
          content: [],
          toolCallId: toolCall.id,
          toolName: toolCall.name,
          timestamp: Date.now(),
        }
      : this.result;

    const renderResult = renderTool(
      toolName,
      toolCall.arguments,
      result,
      !this.aborted && (this.isStreaming || this.pending),
    );

    // Custom renderers own their full card; fallback renders get a wrapper.
    if (renderResult.isCustom) {
      return renderResult.content;
    }

    return html`<div class="pi-tool-card-fallback">${renderResult.content}</div>`;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "user-message": UserMessage;
    "assistant-message": AssistantMessage;
    "tool-message": ToolMessage;
  }
}
