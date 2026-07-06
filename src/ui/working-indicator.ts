/**
 * Pi for Excel — Working indicator with rotating hints and whimsical status.
 *
 * Shows while the agent is streaming. Two independently rotating texts:
 * - Left: whimsical "working" phrases (rotate every ~6s)
 * - Right: feature discovery hints (rotate every ~4.5s, starts with "escape to interrupt")
 *
 * Staggered timers so both don't change simultaneously.
 */

import { html, LitElement } from "lit";
import { customElement, property, state } from "lit/decorators.js";
import { pickWhimsicalMessage } from "./whimsical-messages.js";
import { t } from "../language/index.js";

// Locale keys — resolved lazily via t() so the active language (set at boot,
// after module import) is respected. Do NOT call t() at module scope.
const HINT_KEYS: string[] = [
  "working.hint.escape",
  "working.hint.reasoning",
  "working.hint.commands",
  "working.hint.collapse",
  "working.hint.redirect",
];

@customElement("pi-working-indicator")
export class WorkingIndicator extends LitElement {
  @property({ type: Boolean }) active = false;

  /** Optional fixed label (disables whimsical rotation). */
  @property({ type: String }) primaryText?: string;

  /** Optional fixed hint (disables hint rotation). */
  @property({ type: String }) hintText?: string;

  @state() private _whimsical = t("working.default");
  @state() private _hintIndex = 0;
  @state() private _fadingWhimsical = false;
  @state() private _fadingHint = false;

  private _whimsicalTimer?: ReturnType<typeof setInterval>;
  private _hintTimer?: ReturnType<typeof setInterval>;
  private _staggerTimeout?: ReturnType<typeof setTimeout>;

  protected override createRenderRoot() { return this; }

  override connectedCallback() {
    super.connectedCallback();
    this.style.display = "block";
    // If already active when connected, start immediately
    if (this.active) this._startRotation();
  }

  override updated(changed: Map<string, DynamicValue>) {
    if (changed.has("active") || changed.has("primaryText") || changed.has("hintText")) {
      if (this.active) this._startRotation();
      else this._stopRotation();
    }
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    this._stopRotation();
  }

  private _startRotation() {
    // If we're in fixed-text mode, don't start timers.
    if (this.primaryText || this.hintText) {
      this._stopRotation();
      this._fadingWhimsical = false;
      this._fadingHint = false;
      this._whimsical = this.primaryText || t("working.default");
      this._hintIndex = 0;
      return;
    }

    // Idempotent — safe to call from both connectedCallback and updated
    if (this._hintTimer) return;
    this._stopRotation();
    // Reset to initial state — random hint from the start
    this._whimsical = t("working.default");
    this._hintIndex = Math.floor(Math.random() * HINT_KEYS.length);
    this._fadingWhimsical = false;
    this._fadingHint = false;

    // First whimsical swap after 1s, then every ~6s
    this._staggerTimeout = setTimeout(() => {
      if (!this.active) return;
      this._rotateWhimsical();
      this._whimsicalTimer = setInterval(() => this._rotateWhimsical(), 6000);
    }, 1000);
    // Hints rotate every ~4.5s
    this._hintTimer = setInterval(() => this._rotateHint(), 4500);
  }

  private _stopRotation() {
    if (this._staggerTimeout) { clearTimeout(this._staggerTimeout); this._staggerTimeout = undefined; }
    if (this._whimsicalTimer) { clearInterval(this._whimsicalTimer); this._whimsicalTimer = undefined; }
    if (this._hintTimer) { clearInterval(this._hintTimer); this._hintTimer = undefined; }
  }

  private _rotateHint() {
    this._fadingHint = true;
    setTimeout(() => {
      // Random pick, avoiding current
      let next: number;
      do { next = Math.floor(Math.random() * HINT_KEYS.length); } while (next === this._hintIndex && HINT_KEYS.length > 1);
      this._hintIndex = next;
      this._fadingHint = false;
    }, 250); // half of the CSS transition duration
  }

  private _rotateWhimsical() {
    this._fadingWhimsical = true;
    setTimeout(() => {
      this._whimsical = pickWhimsicalMessage(this._whimsical);
      this._fadingWhimsical = false;
    }, 250);
  }

  override render() {
    if (!this.active) return html``;

    const left = this.primaryText || this._whimsical;
    const right = this.hintText || t(HINT_KEYS[this._hintIndex]);

    const fixed = Boolean(this.primaryText || this.hintText);

    return html`
      <div class="pi-working ${fixed ? "pi-working--fixed" : ""}">
        ${fixed
          ? html`
            <span class="pi-working__pill">
              <span class="pi-working__spinner" aria-hidden="true"></span>
              ${left}
            </span>
          `
          : html`
            <span class="pi-working__text ${this._fadingWhimsical ? "pi-working--fading" : ""}">
              ${left}
            </span>
          `
        }

        <span class="pi-working__hint ${fixed ? "" : (this._fadingHint ? "pi-working--fading" : "")}">
          ${right}
        </span>
      </div>
    `;
  }
}
