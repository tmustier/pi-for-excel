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
 *   'pi-input-action' → detail: { action: PiInputAction }
 */

import { html, LitElement } from "lit";
import { customElement, property, state, query } from "lit/decorators.js";

import { doesOverlayClaimEscape } from "../utils/escape-guard.js";

const PLACEHOLDER_HINTS = [
  "Ask Pi anything about your workbook…",
  "Type / for commands…",
  "Ask Pi anything about your workbook…",
  "Tell Pi what to change…",
];

export type PiInputAction = "open-files" | "open-rules" | "open-resume" | "open-backups";

@customElement("pi-input")
export class PiInput extends LitElement {
  @property({ type: Boolean }) isStreaming = false;
  @property({ type: Boolean }) hasRecoveryCheckpoints = false;

  @state() private _value = "";
  @state() private _placeholderIndex = 0;
  @state() private _isDragOver = false;
  @state() private _actionsMenuOpen = false;
  @query("textarea") private _textarea!: HTMLTextAreaElement;
  @query(".pi-input-file") private _fileInput?: HTMLInputElement;

  private _placeholderTimer?: ReturnType<typeof setInterval>;
  private _actionsMenuDocumentClickHandler?: (event: MouseEvent) => void;
  private _actionsMenuDocumentKeydownHandler?: (event: KeyboardEvent) => void;

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

  focus(): void { this._textarea?.focus(); }

  protected override createRenderRoot() { return this; }

  private _onInput = (e: Event) => {
    this._value = (e.target as HTMLTextAreaElement).value;
    this._autoGrow();
    this.dispatchEvent(new Event("input", { bubbles: true }));
  };

  private _onKeydown = (e: KeyboardEvent) => {
    if (e.key === "Enter" && !e.shiftKey) {
      if (this.isStreaming) return;
      if (!this._value.trim()) return;
      if (this._value.startsWith("/")) return;
      e.preventDefault();
      this._send();
      return;
    }

    if (e.key === "Escape" && this._actionsMenuOpen) {
      e.preventDefault();
      this._closeActionsMenu();
      return;
    }

    if (e.key === "Escape" && this.isStreaming) {
      if (doesOverlayClaimEscape(e.target)) return;
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

  private _openFilePicker = () => {
    this._fileInput?.click();
  };

  private _onFileInputChange = (event: Event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement) || !target.files) return;

    const files = Array.from(target.files);
    this._dispatchFiles(files);
    target.value = "";
  };

  private _dispatchInputAction(action: PiInputAction): void {
    this.dispatchEvent(new CustomEvent<{ action: PiInputAction }>("pi-input-action", {
      bubbles: true,
      detail: { action },
    }));
  }

  private _attachActionsMenuDocumentListener(): void {
    if (!this._actionsMenuDocumentClickHandler) {
      this._actionsMenuDocumentClickHandler = (event: MouseEvent) => {
        const anchor = this.querySelector(".pi-input-actions-anchor");
        const target = event.target;
        if (anchor && target instanceof Node && anchor.contains(target)) {
          return;
        }

        this._closeActionsMenu();
      };

      document.addEventListener("click", this._actionsMenuDocumentClickHandler, true);
    }

    if (!this._actionsMenuDocumentKeydownHandler) {
      this._actionsMenuDocumentKeydownHandler = (event: KeyboardEvent) => {
        if (event.key !== "Escape") {
          return;
        }

        event.preventDefault();
        event.stopPropagation();
        this._closeActionsMenu();

        const actionsButton = this.querySelector<HTMLButtonElement>(".pi-input-btn--actions");
        actionsButton?.focus();
      };

      document.addEventListener("keydown", this._actionsMenuDocumentKeydownHandler, true);
    }
  }

  private _detachActionsMenuDocumentListener(): void {
    if (this._actionsMenuDocumentClickHandler) {
      document.removeEventListener("click", this._actionsMenuDocumentClickHandler, true);
      this._actionsMenuDocumentClickHandler = undefined;
    }

    if (this._actionsMenuDocumentKeydownHandler) {
      document.removeEventListener("keydown", this._actionsMenuDocumentKeydownHandler, true);
      this._actionsMenuDocumentKeydownHandler = undefined;
    }
  }

  private _openActionsMenu(): void {
    if (this._actionsMenuOpen) return;

    this._actionsMenuOpen = true;
    this._attachActionsMenuDocumentListener();
  }

  private _closeActionsMenu(): void {
    if (!this._actionsMenuOpen) return;

    this._actionsMenuOpen = false;
    this._detachActionsMenuDocumentListener();
  }

  private _toggleActionsMenu(): void {
    if (this._actionsMenuOpen) {
      this._closeActionsMenu();
      return;
    }

    this._openActionsMenu();
  }

  private _onActionImportFiles = () => {
    this._closeActionsMenu();
    this._openFilePicker();
  };

  private _onActionClick(action: PiInputAction) {
    this._closeActionsMenu();
    this._dispatchInputAction(action);
  }

  private _send() {
    const text = this._value.trim();
    if (!text) return;
    this.dispatchEvent(new CustomEvent("pi-send", { bubbles: true, detail: { text } }));
  }

  private _autoGrow() {
    const ta = this._textarea;
    if (!ta) return;
    ta.style.height = "auto";
    ta.style.height = Math.min(ta.scrollHeight, window.innerHeight * 0.4) + "px";
  }

  override connectedCallback() {
    super.connectedCallback();
    // Rotate placeholder hints every 8s (mostly default, occasionally slash hint)
    this._placeholderTimer = setInterval(() => {
      if (this.isStreaming || this._value) return; // don't rotate while typing or streaming
      this._placeholderIndex = (this._placeholderIndex + 1) % PLACEHOLDER_HINTS.length;
    }, 8000);
  }

  override disconnectedCallback() {
    super.disconnectedCallback();
    if (this._placeholderTimer) { clearInterval(this._placeholderTimer); this._placeholderTimer = undefined; }
    this._closeActionsMenu();
  }

  override firstUpdated() { this._textarea?.focus(); }

  override render() {
    const hasContent = this._value.trim().length > 0;
    const backupsTitle = this.hasRecoveryCheckpoints
      ? "Browse and restore backups"
      : "Browse backups and create a manual full-workbook backup";

    return html`
      <div
        class="pi-input-card ${this._isDragOver ? "is-drag-over" : ""}"
        @dragenter=${this._onDragEnter}
        @dragover=${this._onDragOver}
        @dragleave=${this._onDragLeave}
        @drop=${this._onDrop}
      >
        <input
          class="pi-input-file"
          type="file"
          multiple
          @change=${this._onFileInputChange}
        />
        <textarea
          class="pi-input-textarea"
          .value=${this._value}
          placeholder=${this.isStreaming ? "Guide response (↵) · New question (⌥↵)" : PLACEHOLDER_HINTS[this._placeholderIndex]}
          rows="1"
          @input=${this._onInput}
          @keydown=${this._onKeydown}
        ></textarea>

        <div class="pi-input-actions-anchor">
          <button
            class="pi-input-btn pi-input-btn--actions"
            type="button"
            @click=${this._toggleActionsMenu}
            aria-label="Input actions"
            title="Input actions"
            aria-haspopup="menu"
            aria-expanded=${this._actionsMenuOpen ? "true" : "false"}
          >
            +
          </button>

          ${this._actionsMenuOpen
            ? html`
              <div class="pi-input-actions-menu" role="menu" aria-label="Input actions">
                <button
                  type="button"
                  role="menuitem"
                  class="pi-input-actions-menu__item"
                  @click=${this._onActionImportFiles}
                >
                  Import files…
                </button>
                <div class="pi-input-actions-menu__divider" role="separator"></div>
                <button
                  type="button"
                  role="menuitem"
                  class="pi-input-actions-menu__item"
                  @click=${() => this._onActionClick("open-files")}
                >
                  Files…
                </button>
                <button
                  type="button"
                  role="menuitem"
                  class="pi-input-actions-menu__item"
                  @click=${() => this._onActionClick("open-rules")}
                >
                  Rules…
                </button>
                <button
                  type="button"
                  role="menuitem"
                  class="pi-input-actions-menu__item"
                  @click=${() => this._onActionClick("open-resume")}
                >
                  Resume session…
                </button>
                <button
                  type="button"
                  role="menuitem"
                  class="pi-input-actions-menu__item"
                  @click=${() => this._onActionClick("open-backups")}
                  title=${backupsTitle}
                >
                  Backups…
                </button>
              </div>
            `
            : null}
        </div>

        <button
          class="pi-input-btn pi-input-btn--attach"
          type="button"
          @click=${this._openFilePicker}
          aria-label="Import files into Files"
          title="Import files into Files"
        >
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m21.44 11.05-8.49 8.49a5.5 5.5 0 0 1-7.78-7.78l9.2-9.19a3.5 3.5 0 0 1 4.95 4.95l-9.19 9.2a1.5 1.5 0 0 1-2.12-2.13l8.49-8.48"/></svg>
        </button>
        ${this._isDragOver
          ? html`<div class="pi-input-drop-hint">Drop files to import into Files</div>`
          : null}
        ${this.isStreaming
          ? html`
            <button class="pi-input-btn pi-input-btn--abort" @click=${() => this.dispatchEvent(new CustomEvent("pi-abort", { bubbles: true }))} aria-label="Stop">
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
