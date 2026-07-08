/**
 * <streaming-message-container> — renders the actively streaming message.
 *
 * First-party replacement for pi-web-ui's StreamingMessageContainer
 * (docs/ui-ownership.md). Keeps the public contract used by <pi-sidebar>:
 * - setMessage(message, immediate) with rAF batching during streaming
 * - deep-clones batched messages so Lit detects nested mutations
 *   (e.g. toolCall.arguments mutated in place while streaming)
 */

import type { AgentMessage, AgentTool } from "@earendil-works/pi-agent-core";
import type { ToolResultMessage } from "@earendil-works/pi-ai";
import { html, LitElement } from "lit";
import { property, state } from "lit/decorators.js";

import "./messages.js";

export class StreamingMessageContainer extends LitElement {
  @property({ type: Array }) tools: AgentTool[] = [];
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: Object }) pendingToolCalls?: ReadonlySet<string>;
  @property({ type: Object }) toolResultsById?: Map<string, ToolResultMessage>;

  @state() private _message: AgentMessage | null = null;

  private _pendingMessage: AgentMessage | null = null;
  private _updateScheduled = false;
  private _immediateUpdate = false;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  /** Update the displayed message, batching non-immediate updates per frame. */
  setMessage(message: AgentMessage | null, immediate = false): void {
    this._pendingMessage = message;

    if (immediate || message === null) {
      this._immediateUpdate = true;
      this._message = message;
      this.requestUpdate();
      this._pendingMessage = null;
      this._updateScheduled = false;
      return;
    }

    if (!this._updateScheduled) {
      this._updateScheduled = true;
      requestAnimationFrame(() => {
        if (!this._immediateUpdate && this._pendingMessage !== null) {
          // Deep-clone so Lit sees new object identities for nested state
          // (e.g. toolCall.arguments mutated in place while streaming).
          this._message = structuredClone(this._pendingMessage);
          this.requestUpdate();
        }
        this._pendingMessage = null;
        this._updateScheduled = false;
        this._immediateUpdate = false;
      });
    }
  }

  override render() {
    const message = this._message;

    if (!message) {
      if (this.isStreaming) {
        return html`
          <div class="pi-streaming-body">
            <span class="pi-streaming-cursor"></span>
          </div>
        `;
      }
      return html``;
    }

    // Only assistant messages stream; user/toolResult messages render in the
    // stable <message-list> immediately.
    if (message.role !== "assistant") {
      return html``;
    }

    return html`
      <div class="pi-streaming-body">
        <assistant-message
          .message=${message}
          .tools=${this.tools}
          .isStreaming=${this.isStreaming}
          .pendingToolCalls=${this.pendingToolCalls}
          .toolResultsById=${this.toolResultsById}
          .hideToolCalls=${false}
        ></assistant-message>
        ${this.isStreaming ? html`<span class="pi-streaming-cursor"></span>` : ""}
      </div>
    `;
  }
}

if (!customElements.get("streaming-message-container")) {
  customElements.define("streaming-message-container", StreamingMessageContainer);
}

declare global {
  interface HTMLElementTagNameMap {
    "streaming-message-container": StreamingMessageContainer;
  }
}
