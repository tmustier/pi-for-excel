/**
 * Stub module for `@earendil-works/pi-web-ui/dist/dialogs/AttachmentOverlay.js`.
 *
 * The real overlay renders PDF/DOCX/XLSX previews and bundles large deps.
 */

export class AttachmentOverlay {
  static open(_attachment: unknown, _onClose?: () => void): void {
    // Keep this non-throwing so a stray click on an attachment tile doesn't crash the UI.
    console.warn("[pi] Attachment preview is not available in this build");
  }
}
