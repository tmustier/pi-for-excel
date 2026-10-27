/**
 * Pi for Excel — Sidebar layout component.
 *
 * Replaces pi-web-ui's ChatPanel + AgentInterface with a purpose-built
 * layout for the ~350px Excel sidebar. Reuses pi-web-ui's content
 * components (message-list, streaming-message-container) for rendering.
 */

import { html, LitElement, nothing, type PropertyValues } from "lit";
import { icon } from "./icons.js";
import { customElement, property, query, state } from "lit/decorators.js";
import type { Agent, AgentEvent, AgentMessage } from "@earendil-works/pi-agent-core";
import type { ToolResultMessage } from "@earendil-works/pi-ai";
import type { StreamingMessageContainer } from "./messages/streaming-message-container.js";
import { Archive, ChevronRight, FileText, Keyboard, Puzzle, RotateCcw, Ruler, Settings, Wrench } from "lucide";
import "./pi-input.js";
import "./working-indicator.js";
import { describePendingToolCall } from "./tool-renderers.js";
import {
  endOpenThinking,
  forgetOpenThinking,
  recordThinkingEnd,
  recordThinkingStart,
  thinkingKey,
} from "./messages/transcript-state.js";
import { focusInitialMenuItem, handleMenuKeydown } from "./menu-keyboard.js";
import type { PiInput } from "./pi-input.js";
import { isDebugEnabled, formatK } from "../debug/debug.js";
import {
  getPayloadStats,
  getLastContext,
  getPayloadSnapshots,
  type PayloadSnapshot,
  type PayloadShapeSummary,
  type PayloadStats,
} from "../auth/stream-proxy.js";
import { t } from "../language/index.js";
import { Type } from "typebox";
import { Value } from "typebox/value";

const liveJsonValueSchema = Type.Cyclic({
  JsonValue: Type.Union([
    Type.String(), Type.Number(), Type.Boolean(), Type.Null(),
    Type.Array(Type.Ref("JsonValue")),
    Type.Record(Type.String(), Type.Ref("JsonValue")),
  ]),
}, "JsonValue");

/** The `result` of a tool_execution_end event, as far as a tool card reads it. */
const liveToolResultSchema = Type.Object({
  content: Type.Optional(Type.Array(Type.Union([
    Type.Object({ type: Type.Literal("text"), text: Type.String() }),
    Type.Object({ type: Type.Literal("image"), data: Type.String(), mimeType: Type.String() }),
  ]))),
  details: Type.Optional(liveJsonValueSchema),
});

function hasRenderableMessages(messages: readonly AgentMessage[]): boolean {
  return messages.some((message) => message.role !== "system" && message.role !== "artifact");
}

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
  /** Run ended in this background tab while the user was elsewhere. */
  attention: "none" | "finished" | "failed";
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

function formatPrefixChange(reasons: readonly string[]): string {
  if (reasons.length === 0) {
    return "stable";
  }

  return reasons.join("+");
}

const AUTO_SCROLL_DISENGAGE_PX = 32;
const AUTO_SCROLL_REENGAGE_PX = 20;

/**
 * Inner-scroller hysteresis — same idea as the outer constants above.
 * A per-element scroll listener marks the pane as "detached" when the user
 * scrolls up past the disengage threshold, and clears it near the bottom.
 */
const INNER_SCROLL_DISENGAGE_PX = AUTO_SCROLL_DISENGAGE_PX;
const INNER_SCROLL_REENGAGE_PX = AUTO_SCROLL_REENGAGE_PX;

/**
 * CSS selector for inner scrollable areas inside the streaming message.
 *
 * - `thinking-block markdown-block` — thinking body (max-height 300 px)
 * - `.pi-tool-card__body--open`     — expanded tool card (max-height 2000 px)
 */
const INNER_SCROLLABLE_SELECTOR =
  "thinking-block markdown-block, .pi-tool-card__body--open";

@customElement("pi-sidebar")
export class PiSidebar extends LitElement {
  @property({ type: Boolean }) subscriptionWarning = false;
  @property({ attribute: false }) agent?: Agent;
  @property({ attribute: false }) emptyHints: EmptyHint[] = [];
  /** Handles a submitted draft; returns true when the composer should clear. */
  @property({ attribute: false }) onSend?: (text: string, opts: { followUp: boolean }) => boolean;
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
  @property({ attribute: false }) onOpenExtensions?: () => void;
  @property({ attribute: false }) onOpenSettings?: () => void;
  @property({ attribute: false }) onOpenFilesWorkspace?: () => void;
  @property({ attribute: false }) onFilesDrop?: (files: File[]) => void;
  @property({ attribute: false }) onOpenResumePicker?: () => void;
  @property({ attribute: false }) onOpenRecovery?: () => void;
  @property({ attribute: false }) onOpenShortcuts?: () => void;
  /** The active tab is busy (responding, compacting or waiting to retry). */
  @property({ attribute: false }) tabBusy = false;
  @property({ attribute: false }) turnError: string | null = null;

  @state() private _hasMessages = false;
  @state() private _isStreaming = false;
  @state() private _thinking = false;
  @state() private _toolLabel: string | null = null;
  @state() private _busyLabel: string | null = null;
  @state() private _busyHint: string | null = null;
  @state() private _hostWaitLabel: string | null = null;
  @state() private _hostWaitHint: string | null = null;
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

  private _unsubscribe: (() => void) | undefined;
  /** Results of tools that finished before their batch did (see MessageList.liveToolResults). */
  private _liveToolResults = new Map<string, ToolResultMessage>();
  private _autoScroll = true;
  private _lastScrollTop = 0;
  private _resizeObserver: ResizeObserver | undefined;
  private _scrollContainerEl: HTMLElement | undefined;
  private _scrollListener: (() => void) | undefined;
  /** rAF handle for the inner-container auto-scroll loop (thinking / tool blocks). */
  private _innerScrollRAF: number | undefined;
  /** Inner elements we have snapped to bottom at least once. */
  private _innerScrollSeen = new WeakSet<HTMLElement>();
  /** Inner elements the user has manually scrolled away from bottom. */
  private _innerScrollDetached = new WeakSet<HTMLElement>();
  /** Per-element scroll listeners registered lazily during inner auto-scroll. */
  private _innerScrollListeners = new WeakMap<HTMLElement, () => void>();
  /** Elements with a pending programmatic scroll — consumed by the next scroll event. */
  private _innerScrollProgrammaticPending = new WeakSet<HTMLElement>();
  /** Previous value of `_isStreaming` so we can detect edges in `updated()`. */
  private _wasStreaming = false;
  private _utilitiesMenuClickHandler: ((event: MouseEvent) => void) | undefined;
  private _tabContextMenuClickHandler: ((event: MouseEvent) => void) | undefined;
  private readonly _utilitiesMenuId = "pi-utilities-menu";
  private readonly _tabContextMenuId = "pi-tab-context-menu";
  private readonly _contextPillBodyId = "pi-context-pill-body";
  private _onEscapeKey = (event: KeyboardEvent) => {
    if (event.key !== "Escape") {
      return;
    }

    if (this._tabContextMenuRuntimeId) {
      this._closeTabContextMenu({ restoreFocus: true });
      return;
    }

    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu({ restoreFocus: true });
    }
  };
  private _dockObserver: ResizeObserver | undefined;
  private _observedDock: HTMLElement | undefined;
  private _tabRevealKey = "";
  private _tabRevealPending = false;
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
    this._hasMessages = hasRenderableMessages(this.agent.state.messages);
    this._isStreaming = this.agent.state.isStreaming;
    this.requestUpdate();
  }

  /**
   * Show the active tab's activity (compacting, waiting to retry) in the
   * working indicator, also while Pi is responding. Pass `null` to clear.
   */
  setBusyIndicator(label: string | null, hint?: string | null): void {
    this._busyLabel = label;
    this._busyHint = hint ?? null;
    this.requestUpdate();
  }

  /**
   * Show a host wait (e.g. Excel is in cell-edit mode) that overrides the
   * working indicator even while a turn is streaming. Pass `null` to clear.
   */
  setHostWaitIndicator(label: string | null, hint?: string | null): void {
    this._hostWaitLabel = label;
    this._hostWaitHint = hint ?? null;
    this.requestUpdate();
  }

  sendMessage(text: string): void {
    if (this.onSend?.(text, { followUp: false })) {
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
    this._resizeObserver?.disconnect();
    this._resizeObserver = undefined;
    this._dockObserver?.disconnect();
    this._dockObserver = undefined;
    this._observedDock = undefined;
    window.removeEventListener("resize", this._publishDockHeight);

    if (this._scrollContainerEl && this._scrollListener) {
      this._scrollContainerEl.removeEventListener("scroll", this._scrollListener);
    }
    this._scrollContainerEl = undefined;
    this._scrollListener = undefined;
    this._stopInnerAutoScroll();

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

  private readonly _publishDockHeight = (): void => {
    const dock = this._observedDock;
    if (!dock) return;
    const offset = Math.max(0, window.innerHeight - dock.getBoundingClientRect().top);
    document.documentElement.style.setProperty("--pi-dock-height", `${Math.round(offset)}px`);
  };

  /**
   * Publish the dock's height so toasts sit just above it (see toasts.css).
   * Runs after every render: the dock does not exist until an agent is bound,
   * and it is re-created when the sidebar re-renders without one.
   */
  private _observeDock(): void {
    const dock = this.querySelector<HTMLElement>(".pi-dock") ?? undefined;
    if (dock === this._observedDock || typeof ResizeObserver === "undefined") return;
    this._dockObserver?.disconnect();
    this._observedDock = dock;
    if (!dock) return;
    this._dockObserver ??= new ResizeObserver(this._publishDockHeight);
    this._dockObserver.observe(dock);
    window.removeEventListener("resize", this._publishDockHeight);
    window.addEventListener("resize", this._publishDockHeight);
    this._publishDockHeight();
  }

  override updated(changed: PropertyValues<this>) {
    this._setupAutoScroll();
    this._observeDock();
    this._updateSessionTabOverflow();
    // Reveal the active tab when it, its busy marker, or the tab count changes,
    // never on other re-renders: the ‹ › arrows scroll the strip, and pulling
    // it back would make earlier tabs unreachable. Showing an arrow resizes
    // the scroller, so the reveal also re-runs on the render that follows.
    const activeTab = this.sessionTabs.find((tab) => tab.isActive);
    const revealKey = `${activeTab?.runtimeId ?? ""}:${this.sessionTabs.length}:${activeTab?.isBusy ?? false}`;
    const changedKeys: ReadonlyMap<PropertyKey, unknown> = changed;
    const arrowsChanged = changedKeys.has("_tabCanScrollLeft") || changedKeys.has("_tabCanScrollRight");
    if (revealKey !== this._tabRevealKey || (arrowsChanged && this._tabRevealPending)) {
      this._tabRevealKey = revealKey;
      this._tabRevealPending = true;
      requestAnimationFrame(() => { this._tabRevealPending = false; });
      this.querySelector(".pi-session-tab.is-active")?.scrollIntoView({ block: "nearest", inline: "nearest" });
    }

    if (this._isStreaming !== this._wasStreaming) {
      this._wasStreaming = this._isStreaming;
      if (this._isStreaming) {
        this._startInnerAutoScroll();
      } else {
        this._stopInnerAutoScroll();
      }
    }
    if (changed.has("subscriptionWarning") && this.subscriptionWarning) {
      // A late switch to a subscription model must reveal the notice even
      // when the conversation is already scrolled to its latest answer.
      requestAnimationFrame(() => this.querySelector(".pi-subscription-warning")?.scrollIntoView({ block: "nearest" }));
    }
  }

  /**
   * One model of a streaming turn:
   * - the response being streamed renders only in <streaming-message-container>;
   * - at message_end pi-agent-core has pushed the finished response into
   *   state.messages, so the streaming copy is cleared in the same update and
   *   <message-list> alone draws it from then on, its tool cards spinning
   *   until each one's result arrives;
   * - each tool card finishes at its own tool_execution_end.
   */
  private _setupSubscription() {
    this._unsubscribe?.();
    this._liveToolResults = new Map();
    forgetOpenThinking();
    const agent = this.agent;
    if (!agent) return;

    this._hasMessages = hasRenderableMessages(agent.state.messages);
    this._isStreaming = agent.state.isStreaming;

    this._unsubscribe = agent.subscribe((ev: AgentEvent) => {
      switch (ev.type) {
        case "message_start":
        case "message_end":
          if (ev.type === "message_end" && ev.message.role === "assistant") {
            endOpenThinking(ev.message, Date.now());
            this._streamingContainer?.setMessage(null, true);
          }
          this._hasMessages = hasRenderableMessages(agent.state.messages);
          this._isStreaming = agent.state.isStreaming;
          this.requestUpdate();
          break;
        case "agent_start":
          this._liveToolResults = new Map();
          this._thinking = false;
          this._toolLabel = null;
          this._isStreaming = agent.state.isStreaming;
          this.requestUpdate();
          break;
        case "turn_start":
        case "turn_end":
          this._isStreaming = agent.state.isStreaming;
          this.requestUpdate();
          break;
        case "tool_execution_start":
          this._thinking = false;
          this._toolLabel = describePendingToolCall(ev.toolName, ev.args)
            ?? agent.state.tools.find((tool) => tool.name === ev.toolName)?.label ?? ev.toolName;
          this.requestUpdate();
          break;
        case "tool_execution_end": {
          if (agent.state.pendingToolCalls.size === 0) this._toolLabel = null;
          // pi-agent-core types `result` as any. A shape we cannot read just
          // waits for the batch's toolResult message, as before.
          const raw: unknown = ev.result;
          if (!Value.Check(liveToolResultSchema, raw)) break;
          this._liveToolResults.set(ev.toolCallId, {
            role: "toolResult",
            toolCallId: ev.toolCallId,
            toolName: ev.toolName,
            content: raw.content ?? [],
            ...(raw.details === undefined ? {} : { details: raw.details }),
            isError: ev.isError,
            timestamp: Date.now(),
          });
          this.requestUpdate();
          break;
        }
        case "agent_end":
          this._isStreaming = false;
          if (this._streamingContainer) {
            this._streamingContainer.isStreaming = false;
            this._streamingContainer.setMessage(null, true);
          }
          this.requestUpdate();
          break;
        case "message_update": {
          const update = ev.assistantMessageEvent;
          if (ev.message.role === "assistant") {
            this._thinking = update.type === "thinking_start" || update.type === "thinking_delta";
            if (update.type === "thinking_start") {
              recordThinkingStart(thinkingKey(ev.message, update.contentIndex), Date.now());
            } else if (update.type === "thinking_end") {
              recordThinkingEnd(thinkingKey(ev.message, update.contentIndex), Date.now());
            }
          }
          if (this._streamingContainer) {
            const streaming = agent.state.isStreaming;
            this._streamingContainer.isStreaming = streaming;
            this._streamingContainer.setMessage(ev.message, !streaming);
          }
          this.requestUpdate();
          break;
        }
      }
    });
  }

  private _scrollToBottom(container: HTMLElement): void {
    container.scrollTop = container.scrollHeight;
    this._lastScrollTop = container.scrollTop;
  }

  private _scrollToBottomIfNeeded(): void {
    if (!this._autoScroll || !this._scrollContainerEl) {
      return;
    }

    this._scrollToBottom(this._scrollContainerEl);
  }

  /* ── Inner auto-scroll (thinking / tool blocks) ────────── */

  /**
   * Start the rAF loop that keeps expanded thinking-block and tool-card
   * inner scroll areas pinned to the bottom while the outer auto-follow
   * is active.  Only runs while `_isStreaming` is true.
   */
  private _startInnerAutoScroll(): void {
    if (this._innerScrollRAF !== undefined) return;
    const tick = (): void => {
      this._scrollStreamingInnerContainers();
      this._innerScrollRAF = requestAnimationFrame(tick);
    };
    this._innerScrollRAF = requestAnimationFrame(tick);
  }

  private _stopInnerAutoScroll(): void {
    if (this._innerScrollRAF !== undefined) {
      cancelAnimationFrame(this._innerScrollRAF);
      this._innerScrollRAF = undefined;
    }
    // Reset per-stream follow state. Keep listener registry so we don't
    // duplicate listeners if streaming toggles while the same pane stays mounted.
    this._innerScrollSeen = new WeakSet<HTMLElement>();
    this._innerScrollDetached = new WeakSet<HTMLElement>();
  }

  /** Lazily attach a scroll listener that tracks user-initiated detach/re-engage. */
  private _ensureInnerScrollListener(el: HTMLElement): void {
    if (this._innerScrollListeners.has(el)) return;
    const onScroll = () => {
      // Skip the scroll event fired by our programmatic el.scrollTop assignment.
      // The flag is set per-element before the assignment and consumed here.
      if (this._innerScrollProgrammaticPending.has(el)) {
        this._innerScrollProgrammaticPending.delete(el);
        return;
      }

      const dist = el.scrollHeight - el.scrollTop - el.clientHeight;
      if (dist > INNER_SCROLL_DISENGAGE_PX) {
        this._innerScrollDetached.add(el);
      } else if (dist < INNER_SCROLL_REENGAGE_PX) {
        this._innerScrollDetached.delete(el);
      }
    };
    el.addEventListener("scroll", onScroll, { passive: true });
    this._innerScrollListeners.set(el, onScroll);
  }

  /**
   * For each expanded inner-scrollable container inside the *streaming*
   * message, pin it to the bottom — unless the user has manually scrolled
   * away (tracked by a per-element scroll listener, not a distance guess).
   */
  private _scrollStreamingInnerContainers(): void {
    // Intentionally NOT gated on `this._autoScroll`. Inner containers
    // (thinking blocks, tool cards) should auto-follow independently of
    // whether the user has scrolled up in the outer chat. Each inner
    // element has its own detach/re-engage tracking via _innerScrollDetached.
    if (!this._scrollContainerEl) return;

    const streaming = this._scrollContainerEl.querySelector(
      "streaming-message-container",
    );
    if (!streaming) return;

    const containers = streaming.querySelectorAll<HTMLElement>(
      INNER_SCROLLABLE_SELECTOR,
    );

    for (const el of containers) {
      this._ensureInnerScrollListener(el);

      // Not scrollable yet — content still fits inside the viewport.
      if (el.scrollHeight <= el.clientHeight) continue;

      if (!this._innerScrollSeen.has(el)) {
        // First overflow encounter (just expanded / content just exceeded cap).
        this._innerScrollProgrammaticPending.add(el);
        el.scrollTop = el.scrollHeight;
        this._innerScrollSeen.add(el);
        this._innerScrollDetached.delete(el);
        continue;
      }

      // User scrolled up inside this pane — leave it alone.
      if (this._innerScrollDetached.has(el)) continue;

      // Auto-follow — mark as programmatic so the scroll listener skips it.
      this._innerScrollProgrammaticPending.add(el);
      el.scrollTop = el.scrollHeight;
    }
  }

  /* ── Outer auto-scroll ───────────────────────────────────── */

  private _setupAutoScroll() {
    const container = this._scrollContainer;
    if (!container || this._scrollContainerEl === container) return;

    this._resizeObserver?.disconnect();

    if (this._scrollContainerEl && this._scrollListener) {
      this._scrollContainerEl.removeEventListener("scroll", this._scrollListener);
    }

    this._scrollContainerEl = container;
    this._lastScrollTop = container.scrollTop;

    this._resizeObserver = new ResizeObserver(() => {
      this._scrollToBottomIfNeeded();
    });
    this._resizeObserver.observe(container);

    const content = container.querySelector(".pi-messages__inner");
    if (content) {
      this._resizeObserver.observe(content);
    }

    this._scrollToBottomIfNeeded();

    let lastClientHeight = container.clientHeight;
    this._scrollListener = () => {
      const top = container.scrollTop;
      const distFromBottom = container.scrollHeight - top - container.clientHeight;
      // A pane that got shorter (the composer grew, a queue row appeared)
      // moves the scroll position without the user scrolling; keep following.
      const resized = container.clientHeight !== lastClientHeight;
      lastClientHeight = container.clientHeight;
      if (resized) {
        this._scrollToBottomIfNeeded();
        return;
      }
      if (top < this._lastScrollTop && distFromBottom > AUTO_SCROLL_DISENGAGE_PX) this._autoScroll = false;
      else if (distFromBottom < AUTO_SCROLL_REENGAGE_PX) this._autoScroll = true;
      this._lastScrollTop = top;
    };
    container.addEventListener("scroll", this._scrollListener);
  }

  private _onSend = (e: CustomEvent<{ text: string; followUp: boolean }>) => {
    // A new prompt brings the transcript back to its newest message; a steer or
    // follow-up sent while Pi is busy leaves the user's scroll position alone.
    const startsTurn = !this.tabBusy && !this._isStreaming;
    if (this.onSend?.(e.detail.text, { followUp: e.detail.followUp })) {
      if (startsTurn) this._autoScroll = true;
      this._input?.clear();
    }
  };

  private _onAbort = () => { this.onAbort?.(); };

  private _onSessionTabKeyDown = (runtimeId: string, event: KeyboardEvent) => {
    if (event.key === "ContextMenu" || (event.shiftKey && event.key === "F10")) {
      this._openTabContextMenu(runtimeId, event);
      return;
    }

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

  private _onOpenFilesWorkspace = () => {
    this.onOpenFilesWorkspace?.();
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

  private _openTabContextMenu(runtimeId: string, event: MouseEvent | KeyboardEvent): void {
    event.preventDefault();
    event.stopPropagation();

    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu();
    }

    let point = { x: 0, y: 0 };
    if (event instanceof MouseEvent && (event.clientX !== 0 || event.clientY !== 0)) {
      point = { x: event.clientX, y: event.clientY };
    } else if (event.currentTarget instanceof Element) {
      // Keyboard (Shift+F10 / context-menu key): open under the tab.
      const rect = event.currentTarget.getBoundingClientRect();
      point = { x: rect.left, y: rect.bottom };
    }

    this._tabContextMenuRuntimeId = runtimeId;
    const offset = 6;
    const estimatedMenuWidth = 190;
    const estimatedMenuHeight = 220;
    const margin = 8;
    this._tabContextMenuPosition = {
      x: Math.min(point.x + offset, Math.max(margin, window.innerWidth - estimatedMenuWidth - margin)),
      y: Math.min(point.y + offset, Math.max(margin, window.innerHeight - estimatedMenuHeight - margin)),
    };
    this._attachTabContextMenuDocumentListener();
    void this.updateComplete.then(() => {
      const menu = this.querySelector<HTMLElement>(".pi-session-tab-context-menu");
      if (menu) focusInitialMenuItem(menu);
    });
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

  private _closeTabContextMenu(opts: { restoreFocus?: boolean } = {}): void {
    const runtimeId = this._tabContextMenuRuntimeId;
    this._tabContextMenuRuntimeId = null;
    this._tabContextMenuPosition = null;
    this._detachTabContextMenuDocumentListener();
    if (opts.restoreFocus && runtimeId) {
      const index = this.sessionTabs.findIndex((tab) => tab.runtimeId === runtimeId);
      const buttons = this.querySelectorAll<HTMLButtonElement>(".pi-session-tab__main");
      (buttons[index] ?? null)?.focus();
    }
  }

  private _onTabContextMenuKeyDown = (event: KeyboardEvent) => {
    const menu = event.currentTarget;
    if (!(menu instanceof HTMLElement)) return;
    handleMenuKeydown(event, menu, () => this._closeTabContextMenu({ restoreFocus: true }));
  };

  override render() {
    const agent = this.agent;
    if (!agent) return html``;
    const state = agent.state;
    const toolResultsById = new Map(this._liveToolResults);
    for (const message of state.messages) {
      if (message.role === "toolResult") toolResultsById.set(message.toolCallId, message);
    }

    // Derive from agent state directly — _hasMessages may lag behind after
    // batch operations like replaceMessages() that don't fire per-message events.
    const hasMessages = this._hasMessages || hasRenderableMessages(state.messages);

    return html`
      ${this._renderSessionTabs()}
      ${this._renderTabContextMenuOverlay()}
      <div id="pi-notices" class="pi-notices"></div>
      <div class="pi-messages">
        <div class="pi-messages__inner">
          ${this.subscriptionWarning ? html`
            <aside class="pi-subscription-warning" role="status">
              ${t("subscription.warning")}
              <a href="https://claude.ai/settings/usage" target="_blank" rel="noopener noreferrer">https://claude.ai/settings/usage</a>.
              ${t("subscription.warning.setting")}
            </aside>` : nothing}
          ${hasMessages ? html`
            <message-list
              .messages=${state.messages.slice()}
              .tools=${state.tools}
              .pendingToolCalls=${state.pendingToolCalls}
              .liveToolResults=${toolResultsById}
              .isStreaming=${state.isStreaming || this.tabBusy}
              .errorSummary=${this.turnError}
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
      <div class="pi-dock ${this._isStreaming || this._busyLabel !== null || this._hostWaitLabel !== null ? "pi-dock--busy" : ""}">
        <pi-working-indicator
          .active=${this._isStreaming || this._busyLabel !== null || this._hostWaitLabel !== null}
          .primaryText=${this._busyLabel === t("working.stopping") ? this._busyLabel :
            this._hostWaitLabel ?? this._busyLabel ?? this._toolLabel ?? (this._thinking ? undefined : t("working.default"))}
          .hintText=${this._hostWaitLabel !== null ? (this._hostWaitHint ?? undefined) : (this._busyHint ?? undefined)}
        ></pi-working-indicator>
        <div id="pi-widget-slot" class="pi-widget-slot" style="display:none"></div>
        <div class="pi-input-area">
          <pi-input
            .isStreaming=${this._isStreaming}
            .busy=${this._isStreaming || this.tabBusy}
            .stopping=${this._busyLabel === t("working.stopping")}
            @pi-send=${this._onSend}
            @pi-abort=${this._onAbort}
            @pi-files-drop=${this._onFilesDrop}
            @pi-open-files=${this._onOpenFilesWorkspace}
          ></pi-input>
          <div id="pi-status-bar" class="pi-status-bar"></div>
        </div>
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
            aria-label=${t("sidebar.tabs.scroll_left")}
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
                    aria-label=${[
                      t("sidebar.tabs.open", { title: tab.title }),
                      // The button's label replaces its content, so the state
                      // markers below must be named here too.
                      tab.isBusy
                        ? t("sidebar.tabs.busy")
                        : tab.attention === "failed"
                          ? t("sidebar.tabs.failed")
                          : tab.attention === "finished"
                            ? t("sidebar.tabs.unseen")
                            : null,
                      tab.lockState === "waiting_for_lock" ? t("sidebar.tabs.lock") : null,
                    ].filter((part) => part !== null).join(", ")}
                    aria-haspopup="menu"
                    aria-current=${tab.isActive ? "page" : "false"}
                  >
                    <span class="pi-session-tab__title">${tab.title}</span>
                    ${tab.lockState === "waiting_for_lock"
                      ? html`<span class="pi-session-tab__lock" title=${t("sidebar.tabs.lock.tooltip")}>${t("sidebar.tabs.lock")}</span>`
                      : nothing}
                    ${tab.isBusy
                      ? html`<span class="pi-session-tab__busy" role="img" aria-label=${t("sidebar.tabs.busy")} title=${t("sidebar.tabs.busy")}></span>`
                      : tab.attention === "failed"
                        ? html`<span class="pi-session-tab__marker pi-session-tab__marker--failed" role="img" aria-label=${t("sidebar.tabs.failed")} title=${t("sidebar.tabs.failed")}></span>`
                        : tab.attention === "finished"
                          ? html`<span class="pi-session-tab__marker pi-session-tab__marker--finished" role="img" aria-label=${t("sidebar.tabs.unseen")} title=${t("sidebar.tabs.unseen")}></span>`
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
                          ? t("sidebar.tabs.close.wait")
                          : t("sidebar.tabs.close")}
                        aria-label=${t("sidebar.tabs.close")}
                      >
                        ×
                      </button>
                    `
                    : nothing}
                </div>
              `;
            })}
          </div>
          <button
            class="pi-session-tabs__scroll pi-session-tabs__scroll--right"
            type="button"
            ?hidden=${!this._tabCanScrollRight}
            @click=${() => this._scrollTabs(1)}
            aria-label=${t("sidebar.tabs.scroll_right")}
          >
            ›
          </button>
        </div>
        <!-- Outside the scroller so it stays visible however many tabs are open. -->
        <button class="pi-session-tabs__new" @click=${() => this.onCreateTab?.()} aria-label=${t("sidebar.tabs.new")}>+</button>
        <div class="pi-utilities-anchor">
          <button
            class="pi-utilities-btn"
            @click=${() => this._toggleUtilitiesMenu()}
            aria-label=${t("sidebar.utilities.aria")}
            title=${t("sidebar.utilities.aria")}
            aria-haspopup="menu"
            aria-controls=${this._utilitiesMenuId}
            aria-expanded=${this._utilitiesMenuOpen ? "true" : "false"}
          >
            <span class="pi-utilities-btn__icon" aria-hidden="true">${icon(Settings, "sm")}</span>
          </button>
          ${this._utilitiesMenuOpen ? this._renderUtilitiesMenu() : nothing}
        </div>
      </div>
    `;
  }

  private _renderTabMenuItem(label: string, disabled: boolean, action: () => void, danger = false) {
    return html`
      <button
        role="menuitem"
        type="button"
        tabindex="-1"
        class="pi-session-tab-context-menu__item ${danger ? "pi-session-tab-context-menu__item--danger" : ""}"
        ?disabled=${disabled}
        @click=${() => {
          this._closeTabContextMenu({ restoreFocus: true });
          action();
        }}
      >
        ${label}
      </button>
    `;
  }

  private _renderTabContextMenuOverlay() {
    const tab = this.sessionTabs.find((entry) => entry.runtimeId === this._tabContextMenuRuntimeId);
    if (!tab) return nothing;

    const canCloseTabs = this.sessionTabs.length > 1;
    const closeDisabled = !canCloseTabs || tab.lockState === "holding_lock";
    const closeOthersDisabled = this.sessionTabs.length <= 1 || !this.onCloseOtherTabs;
    const tabIndex = this.sessionTabs.findIndex((entry) => entry.runtimeId === tab.runtimeId);
    const moveLeftDisabled = tabIndex <= 0 || !this.onMoveTabLeft;
    const moveRightDisabled = tabIndex >= this.sessionTabs.length - 1 || !this.onMoveTabRight;
    const id = tab.runtimeId;

    return html`
      <div
        class="pi-session-tab-context-menu pi-session-tab-context-menu--floating"
        id=${this._tabContextMenuId}
        role="menu"
        aria-label=${t("sidebar.contextmenu.aria", { title: tab.title })}
        @keydown=${this._onTabContextMenuKeyDown}
        style=${this._tabContextMenuPosition
          ? `left:${this._tabContextMenuPosition.x}px;top:${this._tabContextMenuPosition.y}px;`
          : ""}
      >
        ${this._renderTabMenuItem(t("sidebar.contextmenu.rename"), !this.onRenameTab, () => this.onRenameTab?.(id))}
        ${this._renderTabMenuItem(t("sidebar.contextmenu.duplicate"), !this.onDuplicateTab, () => this.onDuplicateTab?.(id))}
        ${this._renderTabMenuItem(t("sidebar.contextmenu.move_left"), moveLeftDisabled, () => this.onMoveTabLeft?.(id))}
        ${this._renderTabMenuItem(t("sidebar.contextmenu.move_right"), moveRightDisabled, () => this.onMoveTabRight?.(id))}
        ${this._renderTabMenuItem(t("sidebar.contextmenu.close_others"), closeOthersDisabled, () => this.onCloseOtherTabs?.(id))}
        <div class="pi-session-tab-context-menu__divider" role="separator"></div>
        ${this._renderTabMenuItem(t("sidebar.contextmenu.close"), closeDisabled, () => this.onCloseTab?.(id), true)}
      </div>
    `;
  }

  private _toggleUtilitiesMenu() {
    if (this._utilitiesMenuOpen) {
      this._closeUtilitiesMenu({ restoreFocus: true });
      return;
    }

    this._closeTabContextMenu();
    this._utilitiesMenuOpen = true;
    void this.updateComplete.then(() => {
      const menu = this.querySelector<HTMLElement>(".pi-utilities-menu");
      if (menu) focusInitialMenuItem(menu);
    });
    requestAnimationFrame(() => {
      if (!this.isConnected || !this._utilitiesMenuOpen) return;
      this._attachUtilitiesMenuDocumentListener();
    });
  }

  private _attachUtilitiesMenuDocumentListener() {
    if (this._utilitiesMenuClickHandler) return;

    this._utilitiesMenuClickHandler = (event: MouseEvent) => {
      const anchor = this.querySelector(".pi-utilities-anchor");
      const target = event.target;
      if (anchor && target instanceof Node && !anchor.contains(target)) {
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

  private _closeUtilitiesMenu(opts: { restoreFocus?: boolean } = {}) {
    if (!this._utilitiesMenuOpen) return;
    this._utilitiesMenuOpen = false;
    this._detachUtilitiesMenuDocumentListener();
    if (opts.restoreFocus) this.querySelector<HTMLButtonElement>(".pi-utilities-btn")?.focus();
  }

  private _onUtilitiesMenuKeyDown = (event: KeyboardEvent) => {
    const menu = event.currentTarget;
    if (!(menu instanceof HTMLElement)) return;
    handleMenuKeydown(event, menu, () => this._closeUtilitiesMenu({ restoreFocus: true }));
  };

  private _renderUtilitiesItem(iconNode: Parameters<typeof icon>[0], label: string, action: (() => void) | undefined) {
    return html`
      <button role="menuitem" type="button" tabindex="-1" class="pi-utilities-menu__item" @click=${() => {
        this._closeUtilitiesMenu();
        action?.();
      }}>
        <span class="pi-utilities-menu__item-icon" aria-hidden="true">${icon(iconNode, "sm")}</span>
        <span>${label}</span>
      </button>
    `;
  }

  private _renderUtilitiesMenu() {
    return html`
      <div
        class="pi-utilities-menu"
        id=${this._utilitiesMenuId}
        role="menu"
        aria-label=${t("sidebar.utilities.aria")}
        @keydown=${this._onUtilitiesMenuKeyDown}
      >
        ${this._renderUtilitiesItem(Wrench, t("sidebar.menu.setup"), this.onOpenSettings)}
        ${this._renderUtilitiesItem(Puzzle, t("sidebar.menu.extensions"), this.onOpenExtensions)}
        <div class="pi-utilities-menu__divider" role="separator"></div>
        ${this._renderUtilitiesItem(FileText, t("sidebar.menu.files"), this.onOpenFilesWorkspace)}
        ${this._renderUtilitiesItem(Ruler, t("sidebar.menu.rules"), this.onOpenRules)}
        <div class="pi-utilities-menu__divider" role="separator"></div>
        ${this._renderUtilitiesItem(RotateCcw, t("sidebar.menu.resume"), this.onOpenResumePicker)}
        ${this._renderUtilitiesItem(Archive, t("sidebar.menu.backups"), this.onOpenRecovery)}
        <div class="pi-utilities-menu__divider" role="separator"></div>
        ${this._renderUtilitiesItem(Keyboard, t("sidebar.menu.keyboard_shortcuts"), this.onOpenShortcuts)}
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
        t("sidebar.context_pill.no_snapshots"),
        "",
        t("sidebar.context_pill.hint"),
      ].join("\n");

      return html`
        <div class="pi-message-gutter">
          <div class="pi-context-pill">
            <button
              type="button"
              class="pi-context-pill__header"
              aria-controls=${this._contextPillBodyId}
              aria-expanded=${expanded ? "true" : "false"}
              @click=${this._toggleContextPill}
            >
              <span>${t("sidebar.context_pill.no_calls")}</span>
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

    const prefixChangeSummary = latestSnapshot.prefixChanged
      ? t("sidebar.context_pill.yes", { reasons: formatPrefixChange(latestSnapshot.prefixChangeReasons) })
      : t("sidebar.context_pill.no");

    const summaryRows = [
      `| | value |`,
      `|---|---|`,
      `| Call | #${call} (${latestSnapshot.isToolContinuation ? t("sidebar.context_pill.continuation") : t("sidebar.context_pill.first")}) |`,
      `| Prefix changed | ${prefixChangeSummary} |`,
      `| Prefix churn totals | ${this._payloadStats.prefixChanges} calls (model ${this._payloadStats.prefixModelChanges}, system ${this._payloadStats.prefixSystemPromptChanges}, tools ${this._payloadStats.prefixToolChanges}) |`,
      `| System prompt | ${systemChars.toLocaleString()} chars |`,
      `| Tool schemas (${toolCount}) | ${toolSchemaChars.toLocaleString()} chars |`,
      `| Tool bundle | \`${latestSnapshot.toolBundle}\` |`,
      `| Messages (${messageCount}) | ${messageChars.toLocaleString()} chars |`,
      `| **Total** | **${total.toLocaleString()} chars** |`,
      `| Provider/model | \`${latestSnapshot.provider}/${latestSnapshot.modelId}\` |`,
    ];

    const summaryMd = summaryRows.join("\n");

    const recentMd = [
      `| call | phase | bundle | tools | prefix | total chars | payload shape |`,
      `|---|---|---|---|---|---|---|`,
      ...sessionSnapshots.slice(-8).reverse().map((snapshot) => {
        const phase = snapshot.isToolContinuation ? t("sidebar.context_pill.continuation") : t("sidebar.context_pill.first");
        const tools = snapshot.toolsIncluded ? String(snapshot.toolCount) : t("sidebar.context_pill.stripped");
        const prefix = snapshot.prefixChanged
          ? formatPrefixChange(snapshot.prefixChangeReasons)
          : t("sidebar.context_pill.stable");
        const payloadShape = formatPayloadShape(snapshot.payloadShape);
        return `| #${snapshot.call} | ${phase} | \`${snapshot.toolBundle}\` | ${tools} | ${prefix} | ${snapshot.totalChars.toLocaleString()} | ${payloadShape} |`;
      }),
    ].join("\n");

    // Tools table
    const toolsTableMd = ctx?.tools
      ? [
          `| tool | description | schema |`,
          `|---|---|---|`,
          ...ctx.tools.map((t) => {
            const schemaSize = JSON.stringify(t.parameters).length;
            const desc = (t.description.split("\n")[0] ?? "").slice(0, 80);
            return `| \`${t.name}\` | ${desc} | ${formatK(schemaSize)} |`;
          }),
        ].join("\n")
      : "*(tools were stripped on this call or context snapshot is unavailable)*";

    // System prompt rendered as markdown (not in a code fence)
    const systemMd = ctx?.systemPrompt ?? "*(none captured for this call)*";

    const phaseLabel = latestSnapshot.isToolContinuation ? ` · ${t("sidebar.context_pill.continuation")}` : ` · ${t("sidebar.context_pill.first")}`;

    return html`
      <div class="pi-message-gutter">
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
                <span class="pi-context-pill__section-label">${t("sidebar.context_pill.recent_calls")}</span>
                <markdown-block .content=${recentMd}></markdown-block>
              </div>
              <div class="pi-context-pill__section">
                <div class="pi-context-pill__section-header">
                  <span class="pi-context-pill__section-label">${t("sidebar.context_pill.tools")}</span>
                  ${ctx?.tools ? html`<button class="pi-context-pill__copy" @click=${this._copyToolsJson}>${t("sidebar.context_pill.copy_json")}</button>` : nothing}
                </div>
                <markdown-block .content=${toolsTableMd}></markdown-block>
              </div>
              <div class="pi-context-pill__section">
                <span class="pi-context-pill__section-label">${t("sidebar.context_pill.system_prompt")}</span>
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

  /**
   * A starter hint is added to the end of the draft, never replacing it. It
   * is inserted as an edit so Ctrl/Cmd+Z takes it back out.
   */
  private _applyHintPrompt(prompt: string): void {
    const input = this._input;
    if (!input) {
      this.sendMessage(prompt);
      return;
    }

    const textarea = input.getTextarea();
    textarea.focus();
    const end = textarea.value.length;
    textarea.setSelectionRange(end, end);
    const insert = textarea.value.trim().length > 0 ? `\n\n${prompt}` : prompt;
    // execCommand("insertText") is the only edit that keeps the textarea's
    // undo stack; WKWebView and WebView2 support it.
    const inserted = document.execCommand("insertText", false, insert);
    if (!inserted) {
      input.value = `${textarea.value}${insert}`;
    }
  }

  private _renderEmptyState() {
    return html`
      <div class="pi-empty">
        <div class="pi-empty__content">
          <div class="pi-empty__logo">π</div>
          <p class="pi-empty__tagline">
            ${t("sidebar.empty.tagline")}
          </p>
          <div class="pi-empty__hints">
            ${this.emptyHints.map((hint) => html`
              <button
                class="pi-empty__hint"
                aria-describedby="pi-empty-note"
                @click=${() => this._applyHintPrompt(hint.prompt)}
              >
                <span class="pi-empty__hint-label">${hint.label}</span>
                <span class="pi-empty__hint-preview">${this._summarizeHintPrompt(hint.prompt)}</span>
              </button>
            `)}
          </div>
          <p class="pi-empty__note" id="pi-empty-note">${t("sidebar.empty.hint.title")}</p>
        </div>
      </div>
    `;
  }
}
