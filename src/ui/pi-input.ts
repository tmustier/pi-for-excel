import { html, LitElement } from "lit";
import { icon } from "./icons.js";
import { customElement, property, state, query } from "lit/decorators.js";
import { keyed } from "lit/directives/keyed.js";
import { FileText } from "lucide";

import { t } from "../language/index.js";
import { isMacPlatform } from "../utils/platform.js";
import { getSendText, isSubmitKey, resolveInputAutoGrowHeight } from "./pi-input-behavior.js";

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
  @property({ attribute: false }) attachments: readonly string[] = [];
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: Boolean }) busy = false;
  @property({ type: Boolean }) stopping = false;

  @state() private _value = "";
  @state() private _placeholderIndex = 0;
  @state() private _isDragOver = false;
  @query("textarea") private _textarea!: HTMLTextAreaElement;

  private _placeholderTimer: ReturnType<typeof setInterval> | undefined;
  private _draftOwner = 0;

  get value(): string { return this._value; }
  set value(v: string) {
    const changed = v !== this._value;
    this._value = v;
    if (this._textarea) {
      this._textarea.value = v;
      this._autoGrow();
    }
    if (changed) this.dispatchEvent(new Event("input", { bubbles: true }));
  }

  getTextarea(): HTMLTextAreaElement { return this._textarea; }

  loadDraft(text: string): void {
    const hadFocus = this._textarea !== null && document.activeElement === this._textarea;
    this._draftOwner += 1;
    this._value = text;
    this.requestUpdate();
    void this.updateComplete.then(() => {
      const ta = this._textarea;
      if (!ta) return;
      ta.value = this._value;
      this._autoGrow();
      if (hadFocus) ta.focus();
    });
  }

  clear(): void {
    this.value = "";
  }

  submit(opts?: { followUp?: boolean }): void {
    this._send(opts?.followUp === true);
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
      this.dispatchEvent(new Event("input", { bubbles: true }));
    }
    return nextValue;
  }

  /**
   * `change` and `keyup` also land here because legacy WPS WebViews can miss
   * `input`. Only a real value change needs the resize, which forces a layout
   * of the whole pane, so repeats for the same edit (and keys such as arrows)
   * return early.
   */
  private _onInput = (e: Event) => {
    const nextValue = e.target instanceof HTMLTextAreaElement ? e.target.value : this._readTextareaValue();
    if (nextValue === this._value) return;

    this._value = nextValue;
    this._autoGrow();
    this.dispatchEvent(new Event("input", { bubbles: true }));
  };

  private _onKeydown = (e: KeyboardEvent) => {
    if (!isSubmitKey(e)) return;
    e.preventDefault();
    this._send(e.altKey);
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
    this._send(false);
  };

  private _onAbortClick = (event: MouseEvent) => {
    event.preventDefault();
    event.stopPropagation();
    this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }));
  };

  private _send(followUp: boolean) {
    const text = getSendText(this._syncValueFromTextarea());
    if (!text && this.attachments.length === 0) return;
    this.dispatchEvent(new CustomEvent("pi-send", { bubbles: true, detail: { text: text ?? "", followUp } }));
  }

  private _autoGrow() {
    const ta = this._textarea;
    if (!ta) return;
    const card = ta.closest<HTMLElement>(".pi-input-card");
    if (card) card.style.minHeight = `${card.offsetHeight}px`;
    ta.style.height = "auto";
    const height = resolveInputAutoGrowHeight({
      scrollHeight: ta.scrollHeight,
      viewportHeight: window.innerHeight,
      cssMaxHeight: this._getTextareaCssMaxHeight(ta),
    });
    ta.style.height = `${height}px`;
    if (card) card.style.minHeight = "";
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
    this._placeholderTimer = setInterval(() => {
      if (this.isStreaming || this._value) return;
      this._placeholderIndex = (this._placeholderIndex + 1) % PLACEHOLDER_HINT_KEYS.length;
    }, 8000);
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    if (this._placeholderTimer) { clearInterval(this._placeholderTimer); this._placeholderTimer = undefined; }
  }

  override firstUpdated() { this._textarea?.focus(); }

  override render() {
    const hasContent = this._value.trim().length > 0 || this.attachments.length > 0;

    return html`
      <div
        class="pi-input-card ${this.busy ? "is-busy" : ""} ${this._isDragOver ? "is-drag-over" : ""}"
        @dragenter=${this._onDragEnter}
        @dragover=${this._onDragOver}
        @dragleave=${this._onDragLeave}
        @drop=${this._onDrop}
      >
        ${this.attachments.length > 0 ? html`
          <div class="pi-input-attachments">
            ${this.attachments.map((path) => html`
              <button type="button" class="pi-input-attachment"
                aria-label=${t("input.attachment.remove", { path })}
                @click=${() => this.dispatchEvent(new CustomEvent("pi-attachment-remove", { bubbles: true, detail: { path } }))}
              >${path} ×</button>`)}
          </div>` : null}
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
        ${keyed(this._draftOwner, html`<textarea
          class="pi-input-textarea"
          .value=${this._value}
          placeholder=${this.isStreaming ? t("input.streaming.placeholder", { modifier: isMacPlatform() ? "Option" : "Alt" }) : t(getPlaceholderHintKey(this._placeholderIndex))}
          rows="1"
          aria-label=${t("input.chat.aria")}
          autocomplete="off"
          @input=${this._onInput}
          @change=${this._onInput}
          @keyup=${this._onInput}
          @keydown=${this._onKeydown}
        ></textarea>`)}
        ${this._isDragOver
          ? html`<div class="pi-input-drop-hint">${t("input.drop.hint")}</div>`
          : null}
        ${this.busy
          ? html`
            <button
              class="pi-input-btn pi-input-btn--abort"
              type="button"
              @mousedown=${this._onActionMouseDown}
              @click=${this._onAbortClick}
              aria-label=${t("input.stop.aria")}
              ?disabled=${this.stopping}
            >
              <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><rect x="4" y="4" width="16" height="16" rx="2"/></svg>
            </button>`
          : null}
        ${!this.busy || hasContent ? html`
            <button
              class="pi-input-btn pi-input-btn--send ${hasContent ? "" : "is-disabled"}"
              type="button"
              @mousedown=${this._onActionMouseDown}
              @click=${this._onSendClick}
              aria-label=${t("input.send.aria")}
              aria-disabled=${hasContent ? "false" : "true"}
            >
              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
            </button>` : null}
      </div>
    `;
  }
}
