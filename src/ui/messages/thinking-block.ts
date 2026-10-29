import { html, LitElement } from "lit";
import { customElement, property } from "lit/decorators.js";
import { ChevronRight } from "lucide";

import { icon } from "../icons.js";
import { t } from "../../language/index.js";
import { isDisclosureExpanded, setDisclosureExpanded, subscribeDisclosure } from "./transcript-state.js";
import "./markdown-block.js";

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

    let label = t("compat.thought");
    if (this.isStreaming) {
      label = t("messages.thinking");
    } else if (this.durationMs !== undefined) {
      const seconds = Math.max(1, Math.round(this.durationMs / 1000));
      const duration = seconds < 60 ? `${seconds}s` : `${Math.floor(seconds / 60)}m ${seconds % 60}s`;
      label = t("messages.thoughtFor", { duration });
    }

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
