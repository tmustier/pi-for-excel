/**
 * <user-message>, <assistant-message> — chat message elements.
 *
 * First-party replacements for pi-web-ui's Messages.js components
 * (docs/ui-ownership.md). Tag names, light DOM, and the CSS-facing structure
 * (`.user-message-container`, `.pi-assistant-body`, `.pi-assistant-aborted`)
 * are preserved; Tailwind utility classes are replaced with semantic `pi-*`
 * classes styled in src/ui/theme/content/.
 *
 * Intentional divergences from upstream:
 * - No usage/cost row (theme CSS always hid it; the status bar shows context).
 * - Tool calls are not drawn here: the transcript draws them in the turn's
 *   <turn-group> (turn-model.ts). `contentFrom`/`contentTo` let the list
 *   draw the prose before and after the calls separately.
 */

import type {
  AssistantMessage as AssistantMessageType,
  UserMessage as UserMessageType,
} from "@earendil-works/pi-ai";
import { html, LitElement, nothing, type TemplateResult } from "lit";
import { customElement, property } from "lit/decorators.js";

import { t } from "../../language/index.js";
import { turnFailureFix } from "../../taskpane/turn-error.js";
import { copyText } from "../copy-text.js";
import { showToast } from "../toast.js";
import { markdownText } from "./markdown-block.js";
import type { UserMessageWithAttachments } from "../../messages/attachments.js";
import { getThinkingDurationMs, thinkingKey } from "./transcript-state.js";
import "./attachment-tile.js";
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
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: String }) errorSummary: string | null = null;
  @property({ type: Boolean }) canRetry = false;
  @property({ type: Boolean }) hideError = false;
  /** First content block to draw (the list splits a response around its tool calls). */
  @property({ type: Number }) contentFrom = 0;
  /** One past the last content block to draw; the end of the message by default. */
  @property({ type: Number }) contentTo: number | undefined;
  /** Offer "Copy answer": the answer, not the prose before a turn's tool calls. */
  @property({ type: Boolean }) copyable = true;

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

    const from = this.contentFrom;
    const to = this.contentTo ?? message.content.length;
    const tail = to >= message.content.length;

    // Render content in the order it appears. Only the last block of a
    // streaming response can still be growing, so a thinking block reads
    // "Thinking…" only until the model moves on to text or a tool call.
    const lastIndex = message.content.length - 1;
    const parts: TemplateResult[] = [];
    for (let index = from; index < to; index++) {
      const chunk = message.content[index];
      if (!chunk) break;
      if (chunk.type === "text" && chunk.text.trim() !== "") {
        parts.push(html`<markdown-block .content=${chunk.text} .answer=${!this.isStreaming}></markdown-block>`);
      } else if (chunk.type === "thinking" && chunk.thinking.trim() !== "") {
        const key = thinkingKey(message, index);
        parts.push(html`
          <thinking-block
            .content=${chunk.thinking}
            .isStreaming=${this.isStreaming && index === lastIndex}
            .disclosureKey=${key}
            .durationMs=${getThinkingDurationMs(key)}
          ></thinking-block>
        `);
      }
    }

    const hasText = message.content.slice(from, to).some((chunk) => chunk.type === "text" && chunk.text.trim());
    return html`
      <div>
        ${parts.length ? html`<div class="pi-assistant-body">${parts}</div>` : nothing}
        ${this.copyable && !this.isStreaming && hasText
          ? html`<button type="button" class="pi-answer-copy pi-answer-copy--message" @click=${() => {
              const text = Array.from(this.querySelectorAll(".pi-assistant-body > markdown-block"),
                (block) => markdownText(block).trim()).join("\n\n");
              void copyText(text).then((outcome) => {
                if (outcome === "copied") showToast(t("code-block.copied"));
              });
            }}>${t("messages.copyAnswer")}</button>`
          : nothing}
        ${tail && message.stopReason === "error" && !this.hideError
          ? html`
              <div class="pi-assistant-error" role="alert">
                <p>${this.errorSummary || t("turn.error.generic")}</p>
                ${this.canRetry ? html`
                  <button type="button" class="pi-overlay-btn" @click=${() => this.dispatchEvent(new CustomEvent("pi-retry", { bubbles: true }))}>${t("turn.retry")}</button>
                  <button type="button" class="pi-overlay-btn" @click=${() => this.dispatchEvent(new CustomEvent("pi-fix-error", { bubbles: true, detail: turnFailureFix(message.errorMessage ?? "").command }))}>${turnFailureFix(message.errorMessage ?? "").label}</button>
                ` : nothing}
                <details><summary>${t("turn.error.details")}</summary><pre>${message.errorMessage}</pre></details>
              </div>
            `
          : nothing}
        ${tail && message.stopReason === "aborted"
          ? html`<span class="pi-assistant-aborted">${t("messages.aborted")}</span>`
          : nothing}
        ${tail && this.isStreaming ? html`<span class="pi-streaming-cursor"></span>` : nothing}
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "user-message": UserMessage;
    "assistant-message": AssistantMessage;
  }
}
