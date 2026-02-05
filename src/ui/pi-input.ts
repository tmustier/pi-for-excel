/**
 * Pi for Excel — Chat input component.
 *
 * A frosted glass card with auto-growing textarea and embedded send/abort button.
 * Purpose-built for a narrow sidebar. Replaces pi-web-ui's MessageEditor.
 *
 * Events:
 *   'pi-send'  → detail: { text: string }
 *   'pi-abort' → (no detail)
 *
 * Usage:
 *   <pi-input .isStreaming=${false} placeholder="Ask about your spreadsheet…"></pi-input>
 */

import { html, LitElement, css } from "lit";
import { customElement, property, query, state } from "lit/decorators.js";

@customElement("pi-input")
export class PiInput extends LitElement {
  /* ── Public properties ─────────────────────────────── */
  @property({ type: Boolean }) isStreaming = false;
  @property() placeholder = "Type a message…";

  /* ── Internal state ────────────────────────────────── */
  @state() private _value = "";

  @query("textarea") private _textarea!: HTMLTextAreaElement;

  /* ── Public API ────────────────────────────────────── */

  get value(): string { return this._value; }
  set value(v: string) {
    this._value = v;
    // Sync the textarea if it exists
    if (this._textarea) {
      this._textarea.value = v;
      this._autoGrow();
    }
  }

  getTextarea(): HTMLTextAreaElement { return this._textarea; }

  clear(): void {
    this._value = "";
    if (this._textarea) {
      this._textarea.value = "";
      this._autoGrow();
    }
  }

  focus(): void {
    this._textarea?.focus();
  }

  /* ── Light DOM (shares page styles) ────────────────── */
  protected override createRenderRoot() { return this; }

  /* ── Handlers ──────────────────────────────────────── */

  private _onInput = (e: Event) => {
    this._value = (e.target as HTMLTextAreaElement).value;
    this._autoGrow();
    // Bubble a raw input event so command-menu can listen
    this.dispatchEvent(new Event("input", { bubbles: true }));
  };

  private _onKeydown = (e: KeyboardEvent) => {
    // Enter (no modifier) → send; Shift+Enter → newline
    if (e.key === "Enter" && !e.shiftKey && !e.altKey) {
      // Don't handle if streaming (taskpane.ts handles steer/follow-up)
      if (this.isStreaming) return;
      if (!this._value.trim()) return;
      // Don't handle if it starts with / (slash command — taskpane.ts handles)
      if (this._value.startsWith("/")) return;

      e.preventDefault();
      this._send();
      return;
    }

    if (e.key === "Escape" && this.isStreaming) {
      e.preventDefault();
      this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }));
    }
  };

  private _send() {
    const text = this._value.trim();
    if (!text) return;
    this.dispatchEvent(new CustomEvent("pi-send", { bubbles: true, detail: { text } }));
    // Don't clear here — let the parent clear after confirming send
  }

  private _onAbort = () => {
    this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }));
  };

  private _autoGrow() {
    const ta = this._textarea;
    if (!ta) return;
    ta.style.height = "auto";
    ta.style.height = Math.min(ta.scrollHeight, window.innerHeight * 0.4) + "px";
  }

  override firstUpdated() {
    this._textarea?.focus();
  }

  /* ── Render ────────────────────────────────────────── */

  override render() {
    const hasContent = this._value.trim().length > 0;

    return html`
      <div class="pi-input-card">
        <div class="pi-input-card__sheen"></div>
        <textarea
          class="pi-input-textarea"
          .value=${this._value}
          placeholder=${this.isStreaming ? "Steer (Enter) · Follow-up (⌥Enter)…" : this.placeholder}
          rows="1"
          @input=${this._onInput}
          @keydown=${this._onKeydown}
        ></textarea>
        ${this.isStreaming
          ? html`
            <button class="pi-input-btn pi-input-btn--abort" @click=${this._onAbort} aria-label="Stop">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><rect x="4" y="4" width="16" height="16" rx="2"/></svg>
            </button>`
          : html`
            <button
              class="pi-input-btn pi-input-btn--send ${hasContent ? "" : "is-disabled"}"
              @click=${() => this._send()}
              aria-label="Send"
              ?disabled=${!hasContent}
            >
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
            </button>`
        }
      </div>
    `;
  }
}
