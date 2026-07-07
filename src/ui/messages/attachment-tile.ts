/**
 * <attachment-tile> — compact tile for a message attachment.
 *
 * First-party replacement for pi-web-ui's AttachmentTile (docs/ui-ownership.md).
 * Pi for Excel cannot create attachments (document parsing is not bundled),
 * but restored/shared sessions may contain user-with-attachments messages;
 * this keeps them rendering. Preview click-through is intentionally not
 * supported (the upstream overlay was always stubbed out in this app).
 */

import { html, LitElement } from "lit";
import { customElement, property } from "lit/decorators.js";
import { FileSpreadsheet, FileText } from "lucide";

import { icon } from "../icons.js";
import type { Attachment } from "../../messages/attachments.js";

@customElement("attachment-tile")
export class AttachmentTile extends LitElement {
  @property({ type: Object }) attachment?: Attachment;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override render() {
    const attachment = this.attachment;
    if (!attachment) return html``;

    if (attachment.preview) {
      const mimeType = attachment.type === "image" ? attachment.mimeType : "image/png";
      return html`
        <img
          class="pi-attachment pi-attachment--img"
          src="data:${mimeType};base64,${attachment.preview}"
          alt=${attachment.fileName}
          title=${attachment.fileName}
        />
      `;
    }

    const isExcel =
      attachment.mimeType.includes("spreadsheetml") ||
      attachment.fileName.toLowerCase().endsWith(".xlsx") ||
      attachment.fileName.toLowerCase().endsWith(".xls");

    return html`
      <div class="pi-attachment pi-attachment--doc" title=${attachment.fileName}>
        ${icon(isExcel ? FileSpreadsheet : FileText, "md")}
        <span class="pi-attachment__name">${attachment.fileName}</span>
      </div>
    `;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "attachment-tile": AttachmentTile;
  }
}
