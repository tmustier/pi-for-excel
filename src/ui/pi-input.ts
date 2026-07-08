/**
 * Pi for Excel — Chat input component.
 *
 * A clean card with auto-growing textarea and embedded send/abort button.
 * Purpose-built for a narrow sidebar. Replaces pi-web-ui's MessageEditor.
 *
 * Events:
 *   'pi-send'        → detail: { text: string }
 *   'pi-abort'       → (no detail)
 *   'pi-files-drop'  → detail: { files: File[] }
 *   'pi-open-files'  → (no detail)
 */

import { html, LitElement } from "lit";
import { icon } from "./icons.js";
import { customElement, property, state, query } from "lit/decorators.js";
import { FileText } from "lucide";

import { doesUiClaimStreamingEscape } from "../utils/escape-guard.js";
import { t } from "../language/index.js";
import { getSendText, resolveInputAutoGrowHeight, shouldSendOnEnter } from "./pi-input-behavior.js";

const PLACEHOLDER_HINT_KEYS = [
  "input.placeholder.ask",
  "input.placeholder.commands",
  "input.placeholder.edit",
  "input.placeholder.summarize",
];

function getPlaceholderHintKey(index: number): string {
  return PLACEHOLDER_HINT_KEYS[index] ?? PLACEHOLDER_HINT_KEYS[0] ?? "input.placeholder.ask";
}

@customElement("pi-input")
export class PiInput extends LitElement {
  @property({ type: Boolean }) isStreaming = false;

  @state() private _value = "";
  @state() private _placeholderIndex = 0;
  @state() private _isDragOver = false;
  @query("textarea") private _textarea!: HTMLTextAreaElement;

  private _placeholderTimer: ReturnType<typeof setInterval> | undefined;

  get value(): string { return this._value; }
  set value(v: string) {
    this._value = v;
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

  override focus(): void { this._textarea?.focus(); }

  protected override createRenderRoot() { return this; }

  private _readTextareaValue(): string {
    return this._textarea?.value ?? this._value;
  }

  private _syncValueFromTextarea(): string {
    const nextValue = this._readTextareaValue();
    if (nextValue !== this._value) {
      this._value = nextValue;
    }
    return nextValue;
  }

  private _onInput = (e: Event) => {
    if (e.target instanceof HTMLTextAreaElement) {
      this._value = e.target.value;
    } else {
      this._syncValueFromTextarea();
    }
    this._autoGrow();
    this.dispatchEvent(new Event("input", { bubbles: true }));
  };

  private _onKeydown = (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      const value = this._syncValueFromTextarea();
      if (!shouldSendOnEnter({ key: e.key, shiftKey: e.shiftKey, isStreaming: this.isStreaming, value })) return;
      e.preventDefault();
      this._send();
      return;
    }

    if (e.key === "Escape" && this.isStreaming) {
      if (doesUiClaimStreamingEscape(e.target)) return;
      e.preventDefault();
      this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }));
    }
  };

  private _onDragEnter = (event: DragEvent) => {
    if (!event.dataTransfer || event.dataTransfer.files.length === 0) return;
    event.preventDefault();
    this._isDragOver = true;
  };

  private _onDragOver = (event: DragEvent) => {
    if (!event.dataTransfer) return;
    if (event.dataTransfer.files.length === 0) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = "copy";
    this._isDragOver = true;
  };

  private _onDragLeave = (event: DragEvent) => {
    const related = event.relatedTarget;
    if (related instanceof Node && this.contains(related)) return;
    this._isDragOver = false;
  };

  private _dispatchFiles(files: File[]): void {
    if (files.length === 0) return;

    this.dispatchEvent(new CustomEvent<{ files: File[] }>("pi-files-drop", {
      bubbles: true,
      detail: { files },
    }));
  }

  private _onDrop = (event: DragEvent) => {
    event.preventDefault();
    this._isDragOver = false;

    const transfer = event.dataTransfer;
    if (!transfer || transfer.files.length === 0) return;

    const files = Array.from(transfer.files);
    this._dispatchFiles(files);
  };

  private _openFilesWorkspace = (event?: Event) => {
    event?.preventDefault();
    event?.stopPropagation();
    this.dispatchEvent(new CustomEvent("pi-open-files", { bubbles: true }));
  };

  private _onActionMouseDown = (event: MouseEvent) => {
    event.preventDefault();
  };

  private _onSendClick = (event: MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    this._send();
  };

  private _onAbortClick = (event: MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }));
  };

  private _send() {
    const text = getSendText(this._syncValueFromTextarea());
    if (!text) return;
    this.dispatchEvent(new CustomEvent("pi-send", { bubbles: true, detail: { text } }));
  }

  private _autoGrow() {
    const ta = this._textarea;
    if (!ta) return;
    ta.style.height = "auto";
    const height = resolveInputAutoGrowHeight({
      scrollHeight: ta.scrollHeight,
      viewportHeight: window.innerHeight,
      cssMaxHeight: this._getTextareaCssMaxHeight(ta),
    });
    ta.style.height = `${height}px`;
  }

  private _getTextareaCssMaxHeight(ta: HTMLTextAreaElement): number {
    if (typeof window.getComputedStyle !== "function") return Number.NaN;
    const computed = window.getComputedStyle(ta).maxHeight;
    if (!computed || computed === "none") return Number.NaN;
    const parsed = Number.parseFloat(computed);
    return Number.isFinite(parsed) ? parsed : Number.NaN;
  }

  override connectedCallback() {
    super.connectedCallback();
    // Rotate placeholder hints every 8s (mostly default, occasionally slash hint)
    this._placeholderTimer = setInterval(() => {
      if (this.isStreaming || this._value) return; // don't rotate while typing or streaming
      this._placeholderIndex = (this._placeholderIndex + 1) % PLACEHOLDER_HINT_KEYS.length;
    }, 8000);
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    if (this._placeholderTimer) { clearInterval(this._placeholderTimer); this._placeholderTimer = undefined; }
  }

  override firstUpdated() { this._textarea?.focus(); }

  override render() {
    const hasContent = this._value.trim().length > 0;

    return html`
      <div
        class="pi-input-card ${this._isDragOver ? "is-drag-over" : ""}"
        @dragenter=${this._onDragEnter}
        @dragover=${this._onDragOver}
        @dragleave=${this._onDragLeave}
        @drop=${this._onDrop}
      >
        <button
          class="pi-input-btn pi-input-btn--attach"
          type="button"
          @mousedown=${this._onActionMouseDown}
          @click=${this._openFilesWorkspace}
          aria-label=${t("input.attach.aria")}
          title=${t("input.attach.aria")}
        >
          ${icon(FileText, "sm")}
        </button>
        <textarea
          class="pi-input-textarea"
          .value=${this._value}
          placeholder=${this.isStreaming ? t("input.streaming.placeholder") : t(getPlaceholderHintKey(this._placeholderIndex))}
          rows="1"
          aria-label=${t("input.chat.aria")}
          autocomplete="off"
          @input=${this._onInput}
          @change=${this._onInput}
          @keyup=${this._onInput}
          @keydown=${this._onKeydown}
        ></textarea>
        ${this._isDragOver
          ? html`<div class="pi-input-drop-hint">${t("input.drop.hint")}</div>`
          : null}
        ${this.isStreaming
          ? html`
            <button
              class="pi-input-btn pi-input-btn--abort"
              type="button"
              @mousedown=${this._onActionMouseDown}
              @click=${this._onAbortClick}
              aria-label=${t("input.stop.aria")}
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><rect x="4" y="4" width="16" height="16" rx="2"/></svg>
            </button>`
          : html`
            <button
              class="pi-input-btn pi-input-btn--send ${hasContent ? "" : "is-disabled"}"
              type="button"
              @mousedown=${this._onActionMouseDown}
              @click=${this._onSendClick}
              aria-label=${t("input.send.aria")}
              aria-disabled=${hasContent ? "false" : "true"}
            >
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
            </button>`
        }
      </div>
    `;
  }
}
