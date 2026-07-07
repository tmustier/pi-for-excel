/**
 * <message-list> — stable (non-streaming) message history.
 *
 * First-party replacement for pi-web-ui's MessageList (docs/ui-ownership.md).
 * Custom role renderers (message-renderer-registry) take precedence; then
 * user/assistant messages render via the built-in elements. Standalone
 * toolResult messages are skipped — they render inside the paired
 * <tool-message> of their assistant message.
 */

import type { AgentMessage, AgentTool } from "@earendil-works/pi-agent-core";
import type { ToolResultMessage } from "@earendil-works/pi-ai";
import { html, LitElement, type TemplateResult } from "lit";
import { property } from "lit/decorators.js";
import { repeat } from "lit/directives/repeat.js";

import { renderMessage } from "./message-renderer-registry.js";
import "./messages.js";

interface RenderItem {
  key: string;
  template: TemplateResult;
}

export class MessageList extends LitElement {
  @property({ type: Array }) messages: AgentMessage[] = [];
  @property({ type: Array }) tools: AgentTool[] = [];
  @property({ type: Object }) pendingToolCalls?: ReadonlySet<string>;
  @property({ type: Boolean }) isStreaming = false;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  private _buildRenderItems(): RenderItem[] {
    // Map tool results by call id for quick lookup.
    const resultByCallId = new Map<string, ToolResultMessage>();
    for (const message of this.messages) {
      if (message.role === "toolResult") {
        resultByCallId.set(message.toolCallId, message);
      }
    }

    const items: RenderItem[] = [];
    let index = 0;

    for (const message of this.messages) {
      // Artifact messages are session-persistence-only.
      if (message.role === "artifact") {
        continue;
      }

      const customTemplate = renderMessage(message);
      if (customTemplate) {
        items.push({ key: `msg:${index}`, template: customTemplate });
        index++;
        continue;
      }

      if (message.role === "user" || message.role === "user-with-attachments") {
        items.push({
          key: `msg:${index}`,
          template: html`<user-message .message=${message}></user-message>`,
        });
        index++;
      } else if (message.role === "assistant") {
        items.push({
          key: `msg:${index}`,
          template: html`
            <assistant-message
              .message=${message}
              .tools=${this.tools}
              .isStreaming=${false}
              .pendingToolCalls=${this.pendingToolCalls}
              .toolResultsById=${resultByCallId}
              .hideToolCalls=${false}
              .hidePendingToolCalls=${this.isStreaming}
            ></assistant-message>
          `,
        });
        index++;
      }
      // Standalone toolResult messages and unknown roles are skipped.
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
