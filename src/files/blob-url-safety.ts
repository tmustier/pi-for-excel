const ACTIVE_CONTENT_MIME_TYPES = new Set<string>([
  "text/html",
  "application/xhtml+xml",
  "image/svg+xml",
  "text/javascript",
  "application/javascript",
]);

export function resolveSafeBlobUrlMimeType(mimeType: string): string {
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
