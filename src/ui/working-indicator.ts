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

const HINTS: string[] = [
  "press Esc to stop",
  "Shift+Tab to change thinking depth",
  "type / to see commands",
  "Ctrl+O to hide details",
  "press Enter to guide the response",
];

@customElement("pi-working-indicator")
export class WorkingIndicator extends LitElement {
  @property({ type: Boolean }) active = false;

  @state() private _whimsical = "Working…";
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

  override updated(changed: Map<string, unknown>) {
    if (changed.has("active")) {
      if (this.active) this._startRotation();
      else this._stopRotation();
    }
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    this._stopRotation();
  }

  private _startRotation() {
    // Idempotent — safe to call from both connectedCallback and updated
    if (this._hintTimer) return;
    this._stopRotation();
    // Reset to initial state — random hint from the start
    this._whimsical = "Working…";
    this._hintIndex = Math.floor(Math.random() * HINTS.length);
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
      do { next = Math.floor(Math.random() * HINTS.length); } while (next === this._hintIndex && HINTS.length > 1);
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
    return html`
      <div class="pi-working">
        <span class="pi-working__text ${this._fadingWhimsical ? "pi-working--fading" : ""}">
          ${this._whimsical}
        </span>
        <span class="pi-working__hint ${this._fadingHint ? "pi-working--fading" : ""}">
          ${HINTS[this._hintIndex]}
        </span>
      </div>
    `;
  }
}
