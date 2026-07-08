/**
 * Custom message renderers.
 *
 * We render compaction as a tool-style collapsible card instead of an assistant
 * text blob.
 */

import { html } from "lit";
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

registerMessageRenderer("archivedMessages", {
  render(message: ArchivedMessagesMessage) {
    const contentRef = createRef<HTMLDivElement>();
    const chevronRef = createRef<HTMLElement>();

    const title = html`
      <span class="pi-tool-card__title">
        <strong>${t("message-renderers.showEarlier")}</strong>
        <span class="pi-tool-card__detail-text">${message.archivedChatMessageCount} chat message${message.archivedChatMessageCount === 1 ? "" : "s"}</span>
      </span>
    `;

    return html`
      <div class="pi-message-gutter">
        <div class="pi-tool-card" data-state="complete" data-tool-name="archive_history">
          <div class="pi-tool-card__header">
            ${renderCollapsibleToolCardHeader("complete", title, contentRef, chevronRef, false)}
          </div>

          <div
            ${ref(contentRef)}
            class="pi-tool-card__body pi-tool-card__body--collapsed"
          >
            <div class="pi-tool-card__inner">
              <div class="pi-tool-card__detail">
                <span class="pi-tool-card__tool-id">archive</span>
              </div>

              <div class="pi-tool-card__section">
                <div class="pi-tool-card__section-label">${t("message-renderers.archived")}</div>
                ${message.archivedMessages.length === 0
                  ? html`<div class="pi-tool-card__plain-text">${t("message-renderers.noArchived")}</div>`
                  : html`
                    <message-list
                      .messages=${message.archivedMessages}
                      .tools=${[]}
                      .pendingToolCalls=${EMPTY_PENDING_TOOL_CALLS}
                      .isStreaming=${false}
                    ></message-list>
                  `}
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  },
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
