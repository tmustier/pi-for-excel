/**
 * Pi for Excel — Sidebar layout component.
 *
 * Replaces pi-web-ui's ChatPanel + AgentInterface with a purpose-built
 * layout for the ~350px Excel sidebar. Reuses pi-web-ui's content
 * components (message-list, streaming-message-container) for rendering.
 */

import { html, LitElement } from "lit";
import { customElement, property, query, state } from "lit/decorators.js";
import type { Agent, AgentEvent } from "@mariozechner/pi-agent-core";
import type { ToolResultMessage } from "@mariozechner/pi-ai";
import type { StreamingMessageContainer } from "@mariozechner/pi-web-ui";
import "@mariozechner/pi-web-ui";
import "./pi-input.js";
import "./working-indicator.js";
import type { PiInput } from "./pi-input.js";

@customElement("pi-sidebar")
export class PiSidebar extends LitElement {
  @property({ attribute: false }) agent?: Agent;
  @property({ attribute: false }) emptyHints: string[] = [];
  @property({ attribute: false }) onSend?: (text: string) => void;
  @property({ attribute: false }) onAbort?: () => void;

  @state() private _hasMessages = false;
  @state() private _isStreaming = false;

  @query(".pi-messages") private _scrollContainer?: HTMLElement;
  @query("streaming-message-container") private _streamingContainer?: StreamingMessageContainer;
  @query("pi-input") private _input?: PiInput;

  private _unsubscribe?: () => void;
  private _autoScroll = true;
  private _lastScrollTop = 0;
  private _resizeObserver?: ResizeObserver;

  getInput(): PiInput | undefined { return this._input ?? undefined; }
  getTextarea(): HTMLTextAreaElement | undefined { return this._input?.getTextarea(); }

  /** Force re-sync from agent state (e.g. after replaceMessages). */
  syncFromAgent(): void {
    if (!this.agent) return;
    this._hasMessages = this.agent.state.messages.length > 0;
    this._isStreaming = this.agent.state.isStreaming;
    this.requestUpdate();
  }

  sendMessage(text: string): void {
    if (this.onSend) {
      this.onSend(text);
      this._input?.clear();
    }
  }

  protected override createRenderRoot() { return this; }

  override connectedCallback() {
    super.connectedCallback();
    this.style.display = "flex";
    this.style.flexDirection = "column";
    this.style.height = "100%";
    this.style.minHeight = "0";
    this.style.position = "relative";
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
    this._resizeObserver?.disconnect();
  }

  override willUpdate(changed: Map<string, any>) {
    if (changed.has("agent")) this._setupSubscription();
  }

  override async firstUpdated() {
    await this.updateComplete;
    this._setupAutoScroll();
  }

  private _setupSubscription() {
    this._unsubscribe?.();
    if (!this.agent) return;

    this._hasMessages = this.agent.state.messages.length > 0;
    this._isStreaming = this.agent.state.isStreaming;

    this._unsubscribe = this.agent.subscribe((ev: AgentEvent) => {
      switch (ev.type) {
        case "message_start":
        case "message_end":
          this._hasMessages = this.agent!.state.messages.length > 0;
          this._isStreaming = this.agent!.state.isStreaming;
          this.requestUpdate();
          break;
        case "turn_start":
        case "turn_end":
        case "agent_start":
          this._isStreaming = this.agent!.state.isStreaming;
          this.requestUpdate();
          break;
        case "agent_end":
          this._isStreaming = false;
          if (this._streamingContainer) {
            this._streamingContainer.isStreaming = false;
            this._streamingContainer.setMessage(null, true);
          }
          this.requestUpdate();
          break;
        case "message_update":
          if (this._streamingContainer) {
            const streaming = this.agent!.state.isStreaming;
            this._streamingContainer.isStreaming = streaming;
            this._streamingContainer.setMessage(ev.message, !streaming);
          }
          break;
      }
    });
  }

  private _setupAutoScroll() {
    const container = this._scrollContainer;
    if (!container) return;
    const content = container.querySelector(".pi-messages__inner");
    if (content) {
      this._resizeObserver = new ResizeObserver(() => {
        if (this._autoScroll && this._scrollContainer) {
          this._scrollContainer.scrollTop = this._scrollContainer.scrollHeight;
        }
      });
      this._resizeObserver.observe(content);
    }
    container.addEventListener("scroll", () => {
      const top = container.scrollTop;
      const distFromBottom = container.scrollHeight - top - container.clientHeight;
      if (top < this._lastScrollTop && distFromBottom > 50) this._autoScroll = false;
      else if (distFromBottom < 10) this._autoScroll = true;
      this._lastScrollTop = top;
    });
  }

  private _onSend = (e: CustomEvent) => {
    this._autoScroll = true;
    this.onSend?.(e.detail.text);
    this._input?.clear();
  };

  private _onAbort = () => { this.onAbort?.(); };

  private _buildToolResultsMap(): Map<string, ToolResultMessage<any>> {
    const map = new Map<string, ToolResultMessage<any>>();
    if (!this.agent) return map;
    for (const msg of this.agent.state.messages) {
      if (msg.role === "toolResult") map.set(msg.toolCallId, msg);
    }
    return map;
  }

  override render() {
    const agent = this.agent;
    if (!agent) return html``;
    const state = agent.state;
    const toolResultsById = this._buildToolResultsMap();

    // Derive from agent state directly — _hasMessages may lag behind after
    // batch operations like replaceMessages() that don't fire per-message events.
    const hasMessages = this._hasMessages || state.messages.length > 0;

    return html`
      <div class="pi-messages">
        <div class="pi-messages__inner">
          ${hasMessages ? html`
            <message-list
              .messages=${state.messages}
              .tools=${state.tools}
              .pendingToolCalls=${state.pendingToolCalls}
              .isStreaming=${state.isStreaming}
            ></message-list>
            <streaming-message-container
              class="${state.isStreaming ? "" : "hidden"}"
              .tools=${state.tools}
              .isStreaming=${state.isStreaming}
              .pendingToolCalls=${state.pendingToolCalls}
              .toolResultsById=${toolResultsById}
            ></streaming-message-container>
          ` : ""}
        </div>
        ${!hasMessages ? this._renderEmptyState() : ""}
      </div>
      <pi-working-indicator .active=${this._isStreaming}></pi-working-indicator>
      <div id="pi-widget-slot" class="pi-widget-slot" style="display:none"></div>
      <div class="pi-input-area">
        <pi-input
          .isStreaming=${this._isStreaming}
          @pi-send=${this._onSend}
          @pi-abort=${this._onAbort}
        ></pi-input>
        <div id="pi-status-bar" class="pi-status-bar"></div>
      </div>
    `;
  }

  private _renderEmptyState() {
    return html`
      <div class="pi-empty">
        <div class="pi-empty__logo">π</div>
        <p class="pi-empty__tagline">
          Your AI assistant for Excel.<br/>Ask anything about your spreadsheet.
        </p>
        <div class="pi-empty__hints">
          ${this.emptyHints.map(hint => html`
            <button class="pi-empty__hint" @click=${() => this.sendMessage(hint)}>
              ${hint}
            </button>
          `)}
        </div>
      </div>
    `;
  }
}
