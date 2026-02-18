/**
 * MIME helpers for Files dialog open/download actions.
 */

const ACTIVE_CONTENT_MIME_TYPES = new Set<string>([
  "text/html",
  "application/xhtml+xml",
  "image/svg+xml",
  "text/javascript",
  "application/javascript",
]);

/**
 * Returns a MIME type safe to open via blob URL under the add-in origin.
 */
export function resolveSafeFilesDialogBlobMimeType(mimeType: string): string {
  const trimmed = mimeType.trim();
  if (trimmed.length === 0) {
    return "application/octet-stream";
  }

  const baseType = trimmed.split(";", 1)[0]?.trim().toLowerCase() ?? "";
  if (ACTIVE_CONTENT_MIME_TYPES.has(baseType)) {
    return "application/octet-stream";
  }

  return trimmed;
}
