/**
 * <thinking-block> — collapsible reasoning section.
 *
 * First-party replacement for pi-web-ui's ThinkingBlock (docs/ui-ownership.md).
 * Keeps the same tag name and DOM contract relied on by theme CSS:
 * - `.thinking-block` card wrapper
 * - `.thinking-header` with chevron span first, label span second
 *   (CSS reorders them visually via flex `order`)
 * - expanded body rendered as a nested <markdown-block .isThinking>
 *
 * Label lifecycle:
 * - this block is the part of the response still streaming → "Thinking…"
 * - finished, with its start and end observed by the sidebar in this
 *   taskpane → "Thought for Xs" / "Xm Xs"
 * - otherwise (restored history, reload mid-response) → "Thought"
 *
 * The duration and expanded state live in transcript-state.ts, keyed by
 * message and content position, so they survive the element being replaced
 * when the response moves from the streaming container into the list.
 */

import { html, LitElement } from "lit";
import { customElement, property } from "lit/decorators.js";
import { ChevronRight } from "lucide";

import { icon } from "../icons.js";
import { t } from "../../language/index.js";
import { isDisclosureExpanded, setDisclosureExpanded, subscribeDisclosure } from "./transcript-state.js";
import "./markdown-block.js";

function formatThinkingDuration(elapsedMs: number): string {
  const totalSeconds = Math.max(1, Math.round(elapsedMs / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;

  if (minutes <= 0) {
    return `${totalSeconds}s`;
  }

  return `${minutes}m ${seconds}s`;
}

@customElement("thinking-block")
export class ThinkingBlock extends LitElement {
  @property() content = "";
  @property({ type: Boolean }) isStreaming = false;
  @property() disclosureKey = "";
  @property({ type: Number }) durationMs: number | undefined;

  private _unsubscribe: (() => void) | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
    this._unsubscribe = subscribeDisclosure(() => this.requestUpdate());
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
  }

  override render() {
    const expanded = isDisclosureExpanded(this.disclosureKey);
    const labelClass = this.isStreaming
      ? "pi-thinking-label pi-thinking-label--streaming"
      : "pi-thinking-label";
    const chevronClass = expanded
      ? "pi-thinking-chevron pi-thinking-chevron--open"
      : "pi-thinking-chevron";

    const label = this.isStreaming
      ? t("messages.thinking")
      : this.durationMs !== undefined
        ? t("messages.thoughtFor", { duration: formatThinkingDuration(this.durationMs) })
        : t("compat.thought");

    return html`
      <div class="thinking-block">
        <button
          type="button"
          class="thinking-header"
          aria-expanded=${expanded ? "true" : "false"}
          @click=${() => setDisclosureExpanded(this.disclosureKey, !expanded)}
        >
          <span class=${chevronClass}>${icon(ChevronRight, "sm")}</span>
          <span class=${labelClass}>${label}</span>
        </button>
        ${expanded
          ? html`<markdown-block .content=${this.content} .isThinking=${true}></markdown-block>`
          : ""}
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "thinking-block": ThinkingBlock;
  }
}
