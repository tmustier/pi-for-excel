/**
 * Pi for Excel — Sidebar layout component.
 *
 * Replaces pi-web-ui's ChatPanel + AgentInterface with a purpose-built
 * layout for the ~350px Excel sidebar. Reuses pi-web-ui's content
 * components (message-list, streaming-message-container) for rendering.
 */

import { html, LitElement, nothing, type PropertyValues } from "lit";
import { icon } from "@mariozechner/mini-lit";
import { customElement, property, query, state } from "lit/decorators.js";
import type { Agent, AgentEvent } from "@mariozechner/pi-agent-core";
import type { ToolResultMessage } from "@mariozechner/pi-ai";
import type { StreamingMessageContainer } from "@mariozechner/pi-web-ui/dist/components/StreamingMessageContainer.js";
import { ChevronRight, Cog } from "lucide";
import "./pi-input.js";
import "./working-indicator.js";
import { initToolGrouping } from "./tool-grouping.js";
import { applyMessageStyleHooks } from "./message-style-hooks.js";
import type { PiInput, PiInputAction } from "./pi-input.js";
import { isDebugEnabled, formatK } from "../debug/debug.js";
import { INTEGRATIONS_MANAGER_LABEL } from "../integrations/naming.js";
import {
  getPayloadStats,
  getLastContext,
  getPayloadSnapshots,
  type PayloadSnapshot,
  type PayloadShapeSummary,
  type PayloadStats,
} from "../auth/stream-proxy.js";

export interface EmptyHint {
  /** Short text shown on the button. */
  label: string;
  /** Full prompt sent when the button is clicked. */
  prompt: string;
}

export type SessionTabLockState = "idle" | "waiting_for_lock" | "holding_lock";

export interface SessionTabView {
  runtimeId: string;
  title: string;
  isActive: boolean;
  isBusy: boolean;
  lockState: SessionTabLockState;
}

function getHorizontalArrowDirection(event: KeyboardEvent): -1 | 1 | null {
  if (event.metaKey || event.ctrlKey || event.altKey || event.shiftKey) {
    return null;
  }

  const key = event.key;
  const code = event.code;
  const keyCode = event.keyCode;

  if (key === "ArrowLeft" || key === "Left" || code === "ArrowLeft" || keyCode === 37) {
    return -1;
  }

  if (key === "ArrowRight" || key === "Right" || code === "ArrowRight" || keyCode === 39) {
    return 1;
  }

  return null;
}

function formatPayloadShape(shape: PayloadShapeSummary | undefined): string {
  if (!shape) return "—";

  if (shape.rootType === "array") {
    const length = shape.rootArrayLength ?? 0;
    return `array[len=${length}]`;
  }

  if (shape.rootType === "null") return "null";
  if (shape.rootType === "primitive") return "primitive";

  const keyPreview = shape.topLevelKeys.slice(0, 4).join(",");
  const keySuffix = shape.topLevelKeys.length > 4 ? ",…" : "";
  const keysLabel = keyPreview.length > 0 ? `${keyPreview}${keySuffix}` : "(none)";

  if (shape.arrayFields.length === 0) {
    return `keys:${keysLabel}`;
  }

  const arrayPreview = shape.arrayFields
    .slice(0, 3)
    .map((field) => `${field.key}:${field.length}`)
    .join(",");
  const arraySuffix = shape.arrayFields.length > 3 ? ",…" : "";
  return `keys:${keysLabel}; arrays:${arrayPreview}${arraySuffix}`;
}

@customElement("pi-sidebar")
export class PiSidebar extends LitElement {
  @property({ attribute: false }) agent?: Agent;
  @property({ attribute: false }) emptyHints: EmptyHint[] = [];
  @property({ attribute: false }) onSend?: (text: string) => void;
  @property({ attribute: false }) onAbort?: () => void;
  @property({ attribute: false }) sessionTabs: SessionTabView[] = [];
  @property({ attribute: false }) onCreateTab?: () => void;
  @property({ attribute: false }) onSelectTab?: (runtimeId: string) => void;
  @property({ attribute: false }) onCloseTab?: (runtimeId: string) => void;
  @property({ attribute: false }) onRenameTab?: (runtimeId: string) => void;
  @property({ attribute: false }) onDuplicateTab?: (runtimeId: string) => void;
  @property({ attribute: false }) onMoveTabLeft?: (runtimeId: string) => void;
  @property({ attribute: false }) onMoveTabRight?: (runtimeId: string) => void;
  @property({ attribute: false }) onCloseOtherTabs?: (runtimeId: string) => void;
  @property({ attribute: false }) onOpenRules?: () => void;
  @property({ attribute: false }) onOpenIntegrations?: () => void;
  @property({ attribute: false }) onOpenSkills?: () => void;
  @property({ attribute: false }) onOpenExtensions?: () => void;
  @property({ attribute: false }) onOpenSettings?: () => void;
  @property({ attribute: false }) onOpenFiles?: () => void;
  @property({ attribute: false }) onFilesDrop?: (files: File[]) => void;
  @property({ attribute: false }) onOpenResumePicker?: () => void;
  @property({ attribute: false }) onOpenRecovery?: () => void;
  @property({ attribute: false }) onReopenLastClosed?: () => void;
  @property({ attribute: false }) onOpenShortcuts?: () => void;
  @property({ type: Boolean }) hasRecoveryCheckpoints = false;

  @state() private _hasMessages = false;
  @state() private _isStreaming = false;
  @state() private _busyLabel: string | null = null;
  @state() private _busyHint: string | null = null;
  @state() private _payloadStats: PayloadStats | null = null;
  @state() private _payloadSnapshots: PayloadSnapshot[] = [];
  @state() private _contextPillExpanded = false;
  @state() private _utilitiesMenuOpen = false;
  @state() private _tabCanScrollLeft = false;
  @state() private _tabCanScrollRight = false;
  @state() private _tabContextMenuRuntimeId: string | null = null;
  @state() private _tabContextMenuPosition: { x: number; y: number } | null = null;

  @query(".pi-messages") private _scrollContainer?: HTMLElement;
  @query("streaming-message-container") private _streamingContainer?: StreamingMessageContainer;
  @query("pi-input") private _input?: PiInput;
  @query(".pi-session-tabs__scroller") private _tabsScroller?: HTMLElement;

  private _unsubscribe?: () => void;
  private _cleanupGrouping?: () => void;
  private _autoScroll = true;
  private _lastScrollTop = 0;
  private _resizeObserver?: ResizeObserver;
  private _scrollContainerEl?: HTMLElement;
  private _scrollListener?: () => void;
  private _groupingRoot?: HTMLElement;
  private _utilitiesMenuClickHandler?: (event: MouseEvent) => void;
  private _tabContextMenuClickHandler?: (event: MouseEvent) => void;
  private readonly _utilitiesMenuId = "pi-utilities-menu";
  private readonly _tabContextMenuId = "pi-tab-context-menu";
  private readonly _contextPillBodyId = "pi-context-pill-body";
  private _onEscapeKey = (event: KeyboardEvent) => {
    if (event.key !== "Escape") {
      return;
    }

    if (this._tabContextMenuRuntimeId) {
      this._closeTabContextMenu();
      return;
    }

    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu();
    }
  };
  private _onWindowResize = () => {
    this._updateSessionTabOverflow();
  };
  private _onPayloadUpdate = () => {
    if (isDebugEnabled()) {
      const s = getPayloadStats();
      this._payloadStats = s.calls > 0 ? { ...s } : null;
      this._payloadSnapshots = [...getPayloadSnapshots()];
    } else {
      this._payloadStats = null;
      this._payloadSnapshots = [];
    }
  };

  getInput(): PiInput | undefined { return this._input ?? undefined; }
  getTextarea(): HTMLTextAreaElement | undefined { return this._input?.getTextarea(); }

  focusTabNavigationAnchor(): boolean {
    const activeTab = this.querySelector<HTMLButtonElement>(".pi-session-tab.is-active .pi-session-tab__main");
    if (activeTab) {
      activeTab.focus();
      return true;
    }

    const firstTab = this.querySelector<HTMLButtonElement>(".pi-session-tab__main");
    if (firstTab) {
      firstTab.focus();
      return true;
    }

    const utilitiesButton = this.querySelector<HTMLButtonElement>(".pi-utilities-btn");
    if (utilitiesButton) {
      utilitiesButton.focus();
      return true;
    }

    return false;
  }

  /** Force re-sync from agent state (e.g. after replaceMessages). */
  syncFromAgent(): void {
    if (!this.agent) return;
    this._hasMessages = this.agent.state.messages.length > 0;
    this._isStreaming = this.agent.state.isStreaming;
    this.requestUpdate();
  }

  /**
   * Show a non-streaming busy indicator (e.g. while `/compact` runs).
   * Pass `null` to clear.
   */
  setBusyIndicator(label: string | null, hint?: string | null): void {
    this._busyLabel = label;
    this._busyHint = hint ?? null;
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
    document.addEventListener("pi:status-update", this._onPayloadUpdate);
    document.addEventListener("pi:debug-changed", this._onPayloadUpdate);
    document.addEventListener("keydown", this._onEscapeKey);
    window.addEventListener("resize", this._onWindowResize);
    this._onPayloadUpdate();
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
    this._cleanupGrouping?.();
    this._cleanupGrouping = undefined;
    this._groupingRoot = undefined;
    this._resizeObserver?.disconnect();
    this._resizeObserver = undefined;

    if (this._scrollContainerEl && this._scrollListener) {
      this._scrollContainerEl.removeEventListener("scroll", this._scrollListener);
    }
    this._scrollContainerEl = undefined;
    this._scrollListener = undefined;

    document.removeEventListener("pi:status-update", this._onPayloadUpdate);
    document.removeEventListener("pi:debug-changed", this._onPayloadUpdate);
    document.removeEventListener("keydown", this._onEscapeKey);
    window.removeEventListener("resize", this._onWindowResize);
    this._detachUtilitiesMenuDocumentListener();
    this._detachTabContextMenuDocumentListener();
  }

  override willUpdate(changed: PropertyValues<this>) {
    if (changed.has("agent")) this._setupSubscription();
  }

  override firstUpdated() {
    this._ensureMessageEnhancements();
  }

  override updated(_changed: PropertyValues<this>) {
    this._ensureMessageEnhancements();
    this._updateSessionTabOverflow();
  }

  private _setupSubscription() {
    this._unsubscribe?.();
    const agent = this.agent;
    if (!agent) return;

    this._hasMessages = agent.state.messages.length > 0;
    this._isStreaming = agent.state.isStreaming;

    this._unsubscribe = agent.subscribe((ev: AgentEvent) => {
      switch (ev.type) {
        case "message_start":
        case "message_end":
          this._hasMessages = agent.state.messages.length > 0;
          this._isStreaming = agent.state.isStreaming;
          this.requestUpdate();
          break;
        case "turn_start":
        case "turn_end":
        case "agent_start":
          this._isStreaming = agent.state.isStreaming;
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
            const streaming = agent.state.isStreaming;
            this._streamingContainer.isStreaming = streaming;
            this._streamingContainer.setMessage(ev.message, !streaming);
          }
          break;
      }
    });
  }

  private _ensureMessageEnhancements() {
    this._setupAutoScroll();

    const inner = this.querySelector<HTMLElement>(".pi-messages__inner");
    if (!inner) return;

    applyMessageStyleHooks(inner);

    if (this._groupingRoot === inner) return;

    this._cleanupGrouping?.();
    this._cleanupGrouping = initToolGrouping(inner);
    this._groupingRoot = inner;
  }

  private _setupAutoScroll() {
    const container = this._scrollContainer;
    if (!container || this._scrollContainerEl === container) return;

    this._resizeObserver?.disconnect();

    if (this._scrollContainerEl && this._scrollListener) {
      this._scrollContainerEl.removeEventListener("scroll", this._scrollListener);
    }

    this._scrollContainerEl = container;

    const content = container.querySelector(".pi-messages__inner");
    if (content) {
      this._resizeObserver = new ResizeObserver(() => {
        if (this._autoScroll && this._scrollContainerEl) {
          this._scrollContainerEl.scrollTop = this._scrollContainerEl.scrollHeight;
        }
      });
      this._resizeObserver.observe(content);
    }

    this._scrollListener = () => {
      const top = container.scrollTop;
      const distFromBottom = container.scrollHeight - top - container.clientHeight;
      if (top < this._lastScrollTop && distFromBottom > 50) this._autoScroll = false;
      else if (distFromBottom < 10) this._autoScroll = true;
      this._lastScrollTop = top;
    };
    container.addEventListener("scroll", this._scrollListener);
  }

  private _onSend = (e: CustomEvent<{ text: string }>) => {
    this._autoScroll = true;
    this.onSend?.(e.detail.text);
    this._input?.clear();
  };

  private _onAbort = () => { this.onAbort?.(); };

  private _onSessionTabKeyDown = (runtimeId: string, event: KeyboardEvent) => {
    const direction = getHorizontalArrowDirection(event);
    if (!direction) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    const tabs = this.sessionTabs;
    if (tabs.length <= 1) {
      return;
    }

    const currentIndex = tabs.findIndex((tab) => tab.runtimeId === runtimeId);
    if (currentIndex < 0) {
      return;
    }

    const nextIndex = (currentIndex + direction + tabs.length) % tabs.length;
    const nextTab = tabs[nextIndex];
    if (!nextTab) {
      return;
    }

    this._closeTabContextMenu();
    this.onSelectTab?.(nextTab.runtimeId);

    requestAnimationFrame(() => {
      this.focusTabNavigationAnchor();
    });
  };

  private _onFilesDrop = (event: CustomEvent<{ files: File[] }>) => {
    this.onFilesDrop?.(event.detail.files);
  };

  private _onInputAction = (event: CustomEvent<{ action: PiInputAction }>) => {
    const { action } = event.detail;

    if (action === "open-files") {
      this.onOpenFiles?.();
      return;
    }

    if (action === "open-rules") {
      this.onOpenRules?.();
      return;
    }

    if (action === "open-resume") {
      this.onOpenResumePicker?.();
      return;
    }

    if (action === "open-backups") {
      this.onOpenRecovery?.();
    }
  };

  private _updateSessionTabOverflow() {
    const scroller = this._tabsScroller;
    if (!scroller) {
      if (this._tabCanScrollLeft) this._tabCanScrollLeft = false;
      if (this._tabCanScrollRight) this._tabCanScrollRight = false;
      return;
    }

    const maxScrollLeft = Math.max(0, scroller.scrollWidth - scroller.clientWidth);
    const canScrollLeft = maxScrollLeft > 1 && scroller.scrollLeft > 1;
    const canScrollRight = maxScrollLeft > 1 && scroller.scrollLeft < maxScrollLeft - 1;

    if (canScrollLeft !== this._tabCanScrollLeft) {
      this._tabCanScrollLeft = canScrollLeft;
    }

    if (canScrollRight !== this._tabCanScrollRight) {
      this._tabCanScrollRight = canScrollRight;
    }
  }

  private _scrollTabs(direction: -1 | 1): void {
    const scroller = this._tabsScroller;
    if (!scroller) return;

    scroller.scrollBy({
      left: 140 * direction,
      behavior: "smooth",
    });

    requestAnimationFrame(() => {
      this._updateSessionTabOverflow();
    });
  }

  private _resolveTabContextMenuPosition(event: MouseEvent): { x: number; y: number } {
    const offset = 6;
    const estimatedMenuWidth = 190;
    const estimatedMenuHeight = 220;
    const margin = 8;

    const x = Math.min(
      event.clientX + offset,
      Math.max(margin, window.innerWidth - estimatedMenuWidth - margin),
    );
    const y = Math.min(
      event.clientY + offset,
      Math.max(margin, window.innerHeight - estimatedMenuHeight - margin),
    );

    return { x, y };
  }

  private _openTabContextMenu(runtimeId: string, event: MouseEvent): void {
    event.preventDefault();
    event.stopPropagation();

    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu();
    }

    this._tabContextMenuRuntimeId = runtimeId;
    this._tabContextMenuPosition = this._resolveTabContextMenuPosition(event);
    this._attachTabContextMenuDocumentListener();
  }

  private _attachTabContextMenuDocumentListener(): void {
    if (this._tabContextMenuClickHandler) return;

    this._tabContextMenuClickHandler = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Node)) {
        this._closeTabContextMenu();
        return;
      }

      const menu = this.querySelector(".pi-session-tab-context-menu");
      if (menu && menu.contains(target)) {
        return;
      }

      this._closeTabContextMenu();
    };

    document.addEventListener("click", this._tabContextMenuClickHandler, true);
  }

  private _detachTabContextMenuDocumentListener(): void {
    if (!this._tabContextMenuClickHandler) return;

    document.removeEventListener("click", this._tabContextMenuClickHandler, true);
    this._tabContextMenuClickHandler = undefined;
  }

  private _closeTabContextMenu(): void {
    this._tabContextMenuRuntimeId = null;
    this._tabContextMenuPosition = null;
    this._detachTabContextMenuDocumentListener();
  }

  private _buildToolResultsMap(): Map<string, ToolResultMessage<unknown>> {
    const map = new Map<string, ToolResultMessage<unknown>>();
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
      ${this._renderSessionTabs()}
      ${this._renderTabContextMenuOverlay()}
      <div class="pi-messages">
        <div class="pi-messages__inner">
          ${hasMessages ? html`
            <message-list
              .messages=${state.messages}
              .tools=${state.tools}
              .pendingToolCalls=${state.pendingToolCalls}
              .isStreaming=${state.isStreaming}
            ></message-list>
            ${this._renderContextPill()}
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
      <pi-working-indicator
        .active=${this._isStreaming || this._busyLabel !== null}
        .primaryText=${this._isStreaming ? undefined : (this._busyLabel ?? undefined)}
        .hintText=${this._isStreaming ? undefined : (this._busyHint ?? undefined)}
      ></pi-working-indicator>
      <div id="pi-widget-slot" class="pi-widget-slot" style="display:none"></div>
      <div class="pi-input-area">
        <pi-input
          .isStreaming=${this._isStreaming}
          .hasRecoveryCheckpoints=${this.hasRecoveryCheckpoints}
          @pi-send=${this._onSend}
          @pi-abort=${this._onAbort}
          @pi-files-drop=${this._onFilesDrop}
          @pi-input-action=${this._onInputAction}
        ></pi-input>
        <div id="pi-status-bar" class="pi-status-bar"></div>
      </div>
    `;
  }

  private _renderSessionTabs() {
    if (this.sessionTabs.length === 0) return nothing;

    const canCloseTabs = this.sessionTabs.length > 1;

    return html`
      <div class="pi-session-tabs">
        <div class="pi-session-tabs__scroller-wrap">
          <button
            class="pi-session-tabs__scroll pi-session-tabs__scroll--left"
            type="button"
            ?hidden=${!this._tabCanScrollLeft}
            @click=${() => this._scrollTabs(-1)}
            aria-label="Scroll tabs left"
          >
            ‹
          </button>
          <div class="pi-session-tabs__scroller" @scroll=${() => this._updateSessionTabOverflow()}>
            ${this.sessionTabs.map((tab) => {
              const isContextOpen = this._tabContextMenuRuntimeId === tab.runtimeId;
              const canCloseThisTab = canCloseTabs && tab.lockState !== "holding_lock";

              return html`
                <div
                  class="pi-session-tab ${tab.isActive ? "is-active" : ""} ${isContextOpen ? "is-menu-open" : ""}"
                  @contextmenu=${(event: MouseEvent) => this._openTabContextMenu(tab.runtimeId, event)}
                >
                  <button
                    class="pi-session-tab__main"
                    @click=${() => {
                      this._closeTabContextMenu();
                      this.onSelectTab?.(tab.runtimeId);
                    }}
                    @dblclick=${() => {
                      this._closeTabContextMenu();
                      this.onRenameTab?.(tab.runtimeId);
                    }}
                    @keydown=${(event: KeyboardEvent) => this._onSessionTabKeyDown(tab.runtimeId, event)}
                    title=${tab.title}
                    aria-label=${`Open tab ${tab.title}`}
                  >
                    <span class="pi-session-tab__title">${tab.title}</span>
                    ${tab.lockState === "waiting_for_lock"
                      ? html`<span class="pi-session-tab__lock">lock…</span>`
                      : nothing}
                    ${tab.isBusy
                      ? html`<span class="pi-session-tab__busy" aria-hidden="true"></span>`
                      : nothing}
                  </button>
                  ${canCloseTabs
                    ? html`
                      <button
                        class="pi-session-tab__close"
                        @click=${(event: Event) => {
                          event.stopPropagation();
                          this._closeTabContextMenu();
                          this.onCloseTab?.(tab.runtimeId);
                        }}
                        ?disabled=${!canCloseThisTab}
                        title=${tab.lockState === "holding_lock"
                          ? "Wait for workbook changes to finish"
                          : "Close tab"}
                        aria-label="Close tab"
                      >
                        ×
                      </button>
                    `
                    : nothing}
                </div>
              `;
            })}
            <button class="pi-session-tabs__new" @click=${() => this.onCreateTab?.()} aria-label="New tab">+</button>
          </div>
          <button
            class="pi-session-tabs__scroll pi-session-tabs__scroll--right"
            type="button"
            ?hidden=${!this._tabCanScrollRight}
            @click=${() => this._scrollTabs(1)}
            aria-label="Scroll tabs right"
          >
            ›
          </button>
        </div>
        <div class="pi-utilities-anchor">
          <button
            class="pi-utilities-btn"
            @click=${() => this._toggleUtilitiesMenu()}
            aria-label="Settings and tools"
            title="Settings and tools"
            aria-haspopup="menu"
            aria-controls=${this._utilitiesMenuId}
            aria-expanded=${this._utilitiesMenuOpen ? "true" : "false"}
          >
            <span class="pi-utilities-btn__icon" aria-hidden="true">${icon(Cog, "sm")}</span>
          </button>
          ${this._utilitiesMenuOpen ? this._renderUtilitiesMenu() : nothing}
        </div>
      </div>
    `;
  }

  private _renderTabContextMenu(tab: SessionTabView) {
    const canCloseTabs = this.sessionTabs.length > 1;
    const closeDisabled = !canCloseTabs || tab.lockState === "holding_lock";
    const closeOthersDisabled = this.sessionTabs.length <= 1 || !this.onCloseOtherTabs;
    const tabIndex = this.sessionTabs.findIndex((entry) => entry.runtimeId === tab.runtimeId);
    const moveLeftDisabled = tabIndex <= 0 || !this.onMoveTabLeft;
    const moveRightDisabled = tabIndex < 0 || tabIndex >= this.sessionTabs.length - 1 || !this.onMoveTabRight;

    return html`
      <div
        class="pi-session-tab-context-menu pi-session-tab-context-menu--floating"
        id=${this._tabContextMenuId}
        role="menu"
        aria-label=${`Tab actions for ${tab.title}`}
        style=${this._tabContextMenuPosition
          ? `left:${this._tabContextMenuPosition.x}px;top:${this._tabContextMenuPosition.y}px;`
          : ""}
      >
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item"
          ?disabled=${!this.onRenameTab}
          @click=${() => {
            this._closeTabContextMenu();
            this.onRenameTab?.(tab.runtimeId);
          }}
        >
          Rename tab…
        </button>
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item"
          ?disabled=${!this.onDuplicateTab}
          @click=${() => {
            this._closeTabContextMenu();
            this.onDuplicateTab?.(tab.runtimeId);
          }}
        >
          Duplicate tab
        </button>
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item"
          ?disabled=${moveLeftDisabled}
          @click=${() => {
            this._closeTabContextMenu();
            this.onMoveTabLeft?.(tab.runtimeId);
          }}
        >
          Move left
        </button>
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item"
          ?disabled=${moveRightDisabled}
          @click=${() => {
            this._closeTabContextMenu();
            this.onMoveTabRight?.(tab.runtimeId);
          }}
        >
          Move right
        </button>
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item"
          ?disabled=${closeOthersDisabled}
          @click=${() => {
            this._closeTabContextMenu();
            this.onCloseOtherTabs?.(tab.runtimeId);
          }}
        >
          Close other tabs
        </button>
        <div class="pi-session-tab-context-menu__divider" role="separator"></div>
        <button
          role="menuitem"
          type="button"
          class="pi-session-tab-context-menu__item pi-session-tab-context-menu__item--danger"
          ?disabled=${closeDisabled}
          @click=${() => {
            this._closeTabContextMenu();
            this.onCloseTab?.(tab.runtimeId);
          }}
        >
          Close tab
        </button>
      </div>
    `;
  }

  private _renderTabContextMenuOverlay() {
    if (!this._tabContextMenuRuntimeId) {
      return nothing;
    }

    const tab = this.sessionTabs.find((entry) => entry.runtimeId === this._tabContextMenuRuntimeId);
    if (!tab) {
      return nothing;
    }

    return this._renderTabContextMenu(tab);
  }

  private _toggleUtilitiesMenu() {
    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu();
      return;
    }

    this._closeTabContextMenu();
    this._utilitiesMenuOpen = true;
    requestAnimationFrame(() => {
      if (!this.isConnected || !this._utilitiesMenuOpen) return;
      this._attachUtilitiesMenuDocumentListener();
    });
  }

  private _attachUtilitiesMenuDocumentListener() {
    if (this._utilitiesMenuClickHandler) return;

    this._utilitiesMenuClickHandler = (event: MouseEvent) => {
      const anchor = this.querySelector(".pi-utilities-anchor");
      if (anchor && !anchor.contains(event.target as Node)) {
        this._closeUtilitiesMenu();
      }
    };

    document.addEventListener("click", this._utilitiesMenuClickHandler, true);
  }

  private _detachUtilitiesMenuDocumentListener() {
    if (!this._utilitiesMenuClickHandler) return;
    document.removeEventListener("click", this._utilitiesMenuClickHandler, true);
    this._utilitiesMenuClickHandler = undefined;
  }

  private _closeUtilitiesMenu() {
    this._utilitiesMenuOpen = false;
    this._detachUtilitiesMenuDocumentListener();
  }

  private _renderUtilitiesMenu() {
    return html`
      <div class="pi-utilities-menu" id=${this._utilitiesMenuId} role="menu" aria-label="Settings and tools">
        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenSettings?.(); }}>
          Settings…
        </button>
        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenRules?.(); }}>
          Rules & conventions…
        </button>
        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenIntegrations?.(); }}>
          ${INTEGRATIONS_MANAGER_LABEL}…
        </button>
        ${this.onOpenSkills
          ? html`
            <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenSkills?.(); }}>
              Skills…
            </button>
          `
          : nothing}
        ${this.onOpenExtensions
          ? html`
            <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenExtensions?.(); }}>
              Extensions…
            </button>
          `
          : nothing}
        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenFiles?.(); }}>
          Files…
        </button>

        <div class="pi-utilities-menu__divider" role="separator"></div>

        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenResumePicker?.(); }}>
          Resume session…
        </button>
        ${this.onReopenLastClosed
          ? html`
            <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onReopenLastClosed?.(); }}>
              Reopen last closed
            </button>
          `
          : nothing}
        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenRecovery?.(); }}>
          Backups…
        </button>

        <div class="pi-utilities-menu__divider" role="separator"></div>

        <button role="menuitem" class="pi-utilities-menu__item" @click=${() => { this._closeUtilitiesMenu(); this.onOpenShortcuts?.(); }}>
          Keyboard shortcuts…
        </button>
      </div>
    `;
  }

  private _toggleContextPill() {
    this._contextPillExpanded = !this._contextPillExpanded;
  }

  private _copyToolsJson() {
    const sessionId = this.agent?.sessionId;
    const ctx = getLastContext(sessionId);
    if (!ctx?.tools) return;
    const json = JSON.stringify(ctx.tools, null, 2);
    navigator.clipboard.writeText(json).catch(() => { /* ignore */ });
  }

  private _renderContextPill() {
    if (!this._payloadStats) return nothing;

    const sessionId = this.agent?.sessionId;
    const sessionSnapshots = sessionId
      ? this._payloadSnapshots.filter((snapshot) => snapshot.sessionId === sessionId)
      : this._payloadSnapshots;

    const latestSnapshot = sessionSnapshots.length > 0
      ? sessionSnapshots[sessionSnapshots.length - 1]
      : null;

    const expanded = this._contextPillExpanded;

    if (!latestSnapshot) {
      const hintMd = [
        "No payload snapshots for this session yet.",
        "",
        "Send a prompt in this tab to capture call-level context details.",
      ].join("\n");

      return html`
        <div class="px-4">
          <div class="pi-context-pill">
            <button
              type="button"
              class="pi-context-pill__header"
              aria-controls=${this._contextPillBodyId}
              aria-expanded=${expanded ? "true" : "false"}
              @click=${this._toggleContextPill}
            >
              <span>Context · no calls yet for this session</span>
              <span class="pi-context-pill__chevron ${expanded ? "pi-context-pill__chevron--open" : ""}">${icon(ChevronRight, "sm")}</span>
            </button>
            ${expanded ? html`
              <div class="pi-context-pill__body" id=${this._contextPillBodyId}>
                <div class="pi-context-pill__section">
                  <markdown-block .content=${hintMd}></markdown-block>
                </div>
              </div>
            ` : nothing}
          </div>
        </div>
      `;
    }

    const call = latestSnapshot.call;
    const systemChars = latestSnapshot.systemChars;
    const toolSchemaChars = latestSnapshot.toolSchemaChars;
    const toolCount = latestSnapshot.toolCount;
    const messageCount = latestSnapshot.messageCount;
    const messageChars = latestSnapshot.messageChars;
    const total = latestSnapshot.totalChars;

    const ctx = expanded ? getLastContext(sessionId) : undefined;

    const summaryRows = [
      `| | value |`,
      `|---|---|`,
      `| Call | #${call}${latestSnapshot.isToolContinuation ? " (continuation)" : " (first)"} |`,
      `| System prompt | ${systemChars.toLocaleString()} chars |`,
      `| Tool schemas (${toolCount}) | ${toolSchemaChars.toLocaleString()} chars |`,
      `| Tool bundle | \`${latestSnapshot.toolBundle}\` |`,
      `| Messages (${messageCount}) | ${messageChars.toLocaleString()} chars |`,
      `| **Total** | **${total.toLocaleString()} chars** |`,
      `| Provider/model | \`${latestSnapshot.provider}/${latestSnapshot.modelId}\` |`,
    ];

    const summaryMd = summaryRows.join("\n");

    const recentMd = [
      `| call | phase | bundle | tools | total chars | payload shape |`,
      `|---|---|---|---|---|---|`,
      ...sessionSnapshots.slice(-8).reverse().map((snapshot) => {
        const phase = snapshot.isToolContinuation ? "continuation" : "first";
        const tools = snapshot.toolsIncluded ? String(snapshot.toolCount) : "stripped";
        const payloadShape = formatPayloadShape(snapshot.payloadShape);
        return `| #${snapshot.call} | ${phase} | \`${snapshot.toolBundle}\` | ${tools} | ${snapshot.totalChars.toLocaleString()} | ${payloadShape} |`;
      }),
    ].join("\n");

    // Tools table
    const toolsTableMd = ctx?.tools
      ? [
          `| tool | description | schema |`,
          `|---|---|---|`,
          ...ctx.tools.map((t) => {
            const schemaSize = JSON.stringify(t.parameters).length;
            const desc = t.description.split("\n")[0].slice(0, 80);
            return `| \`${t.name}\` | ${desc} | ${formatK(schemaSize)} |`;
          }),
        ].join("\n")
      : "*(tools were stripped on this call or context snapshot is unavailable)*";

    // System prompt rendered as markdown (not in a code fence)
    const systemMd = ctx?.systemPrompt ?? "*(none captured for this call)*";

    const phaseLabel = latestSnapshot.isToolContinuation ? " · continuation" : " · first";

    return html`
      <div class="px-4">
        <div class="pi-context-pill">
          <button
            type="button"
            class="pi-context-pill__header"
            aria-controls=${this._contextPillBodyId}
            aria-expanded=${expanded ? "true" : "false"}
            @click=${this._toggleContextPill}
          >
            <span>Context · call #${call}${phaseLabel} · ${formatK(total)} chars</span>
            <span class="pi-context-pill__chevron ${expanded ? "pi-context-pill__chevron--open" : ""}">${icon(ChevronRight, "sm")}</span>
          </button>
          ${expanded ? html`
            <div class="pi-context-pill__body" id=${this._contextPillBodyId}>
              <div class="pi-context-pill__section">
                <markdown-block .content=${summaryMd}></markdown-block>
              </div>
              <div class="pi-context-pill__section">
                <span class="pi-context-pill__section-label">Recent calls</span>
                <markdown-block .content=${recentMd}></markdown-block>
              </div>
              <div class="pi-context-pill__section">
                <div class="pi-context-pill__section-header">
                  <span class="pi-context-pill__section-label">Tools</span>
                  ${ctx?.tools ? html`<button class="pi-context-pill__copy" @click=${this._copyToolsJson}>Copy JSON</button>` : nothing}
                </div>
                <markdown-block .content=${toolsTableMd}></markdown-block>
              </div>
              <div class="pi-context-pill__section">
                <span class="pi-context-pill__section-label">System prompt</span>
                <markdown-block .content=${systemMd}></markdown-block>
              </div>
            </div>
          ` : nothing}
        </div>
      </div>
    `;
  }

  private _summarizeHintPrompt(prompt: string): string {
    const collapsed = prompt.replace(/\s+/gu, " ").trim();
    if (collapsed.length <= 120) {
      return collapsed;
    }

    return `${collapsed.slice(0, 117)}…`;
  }

  private _applyHintPrompt(prompt: string): void {
    const input = this._input;
    if (!input) {
      this.sendMessage(prompt);
      return;
    }

    input.value = prompt;

    const textarea = input.getTextarea();
    textarea.focus();
    const cursor = textarea.value.length;
    textarea.setSelectionRange(cursor, cursor);
  }

  private _renderEmptyState() {
    return html`
      <div class="pi-empty">
        <div class="pi-empty__content">
          <div class="pi-empty__logo">π</div>
          <p class="pi-empty__tagline">
            Reads your workbook, writes formulas, formats cells, and builds models.
          </p>
          <div class="pi-empty__hints">
            ${this.emptyHints.map((hint) => html`
              <button
                class="pi-empty__hint"
                title="Insert prompt into the input so you can edit it before sending"
                @click=${() => this._applyHintPrompt(hint.prompt)}
              >
                <span class="pi-empty__hint-label">${hint.label}</span>
                <span class="pi-empty__hint-preview">${this._summarizeHintPrompt(hint.prompt)}</span>
              </button>
            `)}
          </div>
        </div>
      </div>
    `;
  }
}
