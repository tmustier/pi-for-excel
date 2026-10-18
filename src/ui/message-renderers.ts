/**
 * Custom message renderers.
 *
 * We render compaction as a tool-style collapsible card instead of an assistant
 * text blob.
 */

import { html, LitElement, nothing } from "lit";
import { customElement, property, state } from "lit/decorators.js";
import { createRef, ref } from "lit/directives/ref.js";
import { registerMessageRenderer } from "./messages/message-renderer-registry.js";
import { renderCollapsibleToolCardHeader } from "./tool-card-header.js";
import { t } from "../language/index.js";

import { formatCompactionSummaryExtent, type CompactionSummaryMessage } from "../messages/compaction.js";
import type { ArchivedMessagesMessage } from "../messages/archived-history.js";

// Ensure <markdown-block> and <message-list> are registered.
import "./messages/markdown-block.js";
import "./messages/message-list.js";

const EMPTY_PENDING_TOOL_CALLS = new Set<string>();

@customElement("pi-archived-history")
export class ArchivedHistoryElement extends LitElement {
  @property({ attribute: false }) message?: ArchivedMessagesMessage;
  @state() private expanded = false;
  @state() private visible = 100;

  protected override createRenderRoot(): HTMLElement { return this; }

  override render() {
    const message = this.message;
    if (!message) return nothing;
    const contentRef = createRef<HTMLDivElement>();
    const chevronRef = createRef<HTMLElement>();
    const toggleMessages = (): void => {
      this.expanded = !this.expanded;
      if (!this.expanded) this.visible = 100;
    };

    const title = html`
      <span class="pi-tool-card__title">
        <strong>${t("message-renderers.showEarlier")}</strong>
        <span class="pi-tool-card__detail-text">${message.archivedChatMessageCount} chat message${message.archivedChatMessageCount === 1 ? "" : "s"}</span>
      </span>
    `;

    return html`
      <div class="pi-message-gutter">
        <div class="pi-tool-card" data-state="complete" data-tool-name="archive_history">
          <div class="pi-tool-card__header" @click=${toggleMessages}>
            ${renderCollapsibleToolCardHeader("complete", title, contentRef, chevronRef, this.expanded)}
          </div>

          <div
            ${ref(contentRef)}
            class="pi-tool-card__body ${this.expanded ? "" : "pi-tool-card__body--collapsed"}"
          >
            <div class="pi-tool-card__inner">
              <div class="pi-tool-card__detail">
                <span class="pi-tool-card__tool-id">archive</span>
              </div>

              <div class="pi-tool-card__section">
                <div class="pi-tool-card__section-label">${t("message-renderers.archived")}</div>
                ${!this.expanded ? nothing : message.archivedMessages.length === 0
                  ? html`<div class="pi-tool-card__plain-text">${t("message-renderers.noArchived")}</div>`
                  : html`
                    <message-list .messages=${message.archivedMessages.slice(0, this.visible)}
                      .tools=${[]} .pendingToolCalls=${EMPTY_PENDING_TOOL_CALLS} .isStreaming=${false}></message-list>
                    ${this.visible < message.archivedMessages.length ? html`
                      <button type="button" class="pi-tool-card__toggle" @click=${() => { this.visible += 100; }}>
                        ${t("message-renderers.showEarlier")}
                      </button>` : nothing}
                  `}
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  }
}

registerMessageRenderer("archivedMessages", {
  render: (message: ArchivedMessagesMessage) => html`<pi-archived-history .message=${message}></pi-archived-history>`,
});

registerMessageRenderer("compactionSummary", {
  render(message: CompactionSummaryMessage) {
    const contentRef = createRef<HTMLDivElement>();
    const chevronRef = createRef<HTMLElement>();

    const title = html`
      <span class="pi-tool-card__title">
        <strong>${t("message-renderers.summarized", { extent: formatCompactionSummaryExtent(message) })}</strong>
      </span>
    `;

    return html`
      <div class="pi-message-gutter">
        <div class="pi-tool-card" data-state="complete" data-tool-name="compact">
          <div class="pi-tool-card__header">
            ${renderCollapsibleToolCardHeader("complete", title, contentRef, chevronRef, false)}
          </div>

          <div
            ${ref(contentRef)}
            class="pi-tool-card__body pi-tool-card__body--collapsed"
          >
            <div class="pi-tool-card__inner">
              <div class="pi-tool-card__detail">
                <span class="pi-tool-card__tool-id">compact</span>
              </div>

              <div class="pi-tool-card__section">
                <div class="pi-tool-card__section-label">${t("message-renderers.summaryLabel")}</div>
                <div class="pi-tool-card__markdown">
                  <markdown-block .content=${message.summary || t("message-renderers.noSummary")}></markdown-block>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  },
});
