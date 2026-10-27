/** Real activity labels; playful phrases only while the model is thinking. */
import { html, LitElement } from "lit";
import { customElement, property, state } from "lit/decorators.js";
import { pickWhimsicalMessage } from "./whimsical-messages.js";
import { t } from "../language/index.js";

@customElement("pi-working-indicator")
export class WorkingIndicator extends LitElement {
  @property({ type: Boolean }) active = false;
  /** A real activity label; absent only while thinking. */
  @property({ type: String }) primaryText?: string;
  @property({ type: String }) hintText?: string;
  @state() private _whimsical = "";
  private _timer: ReturnType<typeof setInterval> | undefined;

  protected override createRenderRoot() { return this; }

  override updated(): void {
    if (this.active && !this.primaryText) {
      if (this._timer) return;
      this._whimsical = pickWhimsicalMessage("");
      this._timer = setInterval(() => {
        this._whimsical = pickWhimsicalMessage(this._whimsical);
      }, 6000);
    } else {
      this._stopRotation();
    }
  }

  private _stopRotation(): void {
    if (this._timer) clearInterval(this._timer);
    this._timer = undefined;
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._stopRotation();
  }

  override render() {
    if (!this.active) return html``;
    return html`
      <div class="pi-working pi-working--fixed">
        <span class="pi-working__pill">
          <span class="pi-working__spinner" aria-hidden="true"></span>
          <span class="pi-working__label" title=${this.primaryText || this._whimsical || t("working.default")}>
            ${this.primaryText || this._whimsical || t("working.default")}
          </span>
        </span>
        <span class="pi-working__hint">${this.hintText || t("working.hint.escape")}</span>
      </div>
    `;
  }
}
