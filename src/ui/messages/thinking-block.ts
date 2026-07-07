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
 * Owns the label lifecycle previously provided by the
 * src/compat/thinking-duration.ts MutationObserver patch:
 * - streaming → shimmering "Thinking…"
 * - finished after streaming here → "Thought for Xs" / "Xm Xs"
 * - restored history (never streamed in this element) → "Thought"
 */

import { html, LitElement, type PropertyValues } from "lit";
import { customElement, property, state } from "lit/decorators.js";
import { ChevronRight } from "lucide";

import { icon } from "../icons.js";
import { t } from "../../language/index.js";
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

  @state() private _expanded = false;
  @state() private _completedLabel: string | null = null;

  private _startedAtMs: number | null = null;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  protected override willUpdate(changed: PropertyValues<this>): void {
    super.willUpdate(changed);

    if (this.isStreaming) {
      if (this._startedAtMs === null) {
        this._startedAtMs = Date.now();
      }
      this._completedLabel = null;
      return;
    }

    if (this._completedLabel === null) {
      this._completedLabel =
        this._startedAtMs !== null
          ? t("messages.thoughtFor", { duration: formatThinkingDuration(Date.now() - this._startedAtMs) })
          : t("compat.thought");
    }
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  private _toggle(): void {
    this._expanded = !this._expanded;
  }

  override render() {
    const labelClass = this.isStreaming
      ? "pi-thinking-label pi-thinking-label--streaming"
      : "pi-thinking-label";
    const chevronClass = this._expanded
      ? "pi-thinking-chevron pi-thinking-chevron--open"
      : "pi-thinking-chevron";

    const label = this.isStreaming
      ? t("messages.thinking")
      : this._completedLabel ?? t("compat.thought");

    return html`
      <div class="thinking-block">
        <div class="thinking-header" @click=${this._toggle}>
          <span class=${chevronClass}>${icon(ChevronRight, "sm")}</span>
          <span class=${labelClass}>${label}</span>
        </div>
        ${this._expanded
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
