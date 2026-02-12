/**
 * MIME/type helpers for workspace files.
 */

import type { WorkspaceFileKind } from "./types.js";

const TEXT_MIME_PREFIXES = ["text/"];
const TEXT_MIME_EXACT = new Set([
  "application/json",
  "application/xml",
  "application/javascript",
  "application/x-javascript",
  "application/x-yaml",
  "application/yaml",
  "application/x-sh",
  "application/sql",
  "application/toml",
  "application/x-httpd-php",
]);

const EXTENSION_TO_MIME: Record<string, string> = {
  txt: "text/plain",
  md: "text/markdown",
  markdown: "text/markdown",
  csv: "text/csv",
  tsv: "text/tab-separated-values",
  json: "application/json",
  yaml: "application/yaml",
  yml: "application/yaml",
  xml: "application/xml",
  html: "text/html",
  htm: "text/html",
  css: "text/css",
  js: "application/javascript",
  mjs: "application/javascript",
  ts: "text/plain",
  py: "text/plain",
  sh: "application/x-sh",
  sql: "application/sql",
  toml: "application/toml",
  ini: "text/plain",
  log: "text/plain",
  pdf: "application/pdf",
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
  gif: "image/gif",
  svg: "image/svg+xml",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
};

function getExtension(fileName: string): string {
  const lastDot = fileName.lastIndexOf(".");
  if (lastDot <= 0 || lastDot === fileName.length - 1) return "";
  return fileName.slice(lastDot + 1).toLowerCase();
}

export function inferMimeType(fileName: string, hint?: string): string {
  const normalizedHint = hint?.trim();
  if (normalizedHint) return normalizedHint;

  const ext = getExtension(fileName);
  if (!ext) return "application/octet-stream";

  return EXTENSION_TO_MIME[ext] ?? "application/octet-stream";
}

export function isTextMimeType(mimeType: string): boolean {
  const normalized = mimeType.trim().toLowerCase();
  if (!normalized) return false;

  if (TEXT_MIME_EXACT.has(normalized)) return true;
  return TEXT_MIME_PREFIXES.some((prefix) => normalized.startsWith(prefix));
}

export function inferFileKind(fileName: string, mimeTypeHint?: string): WorkspaceFileKind {
  const mimeType = inferMimeType(fileName, mimeTypeHint);
  return isTextMimeType(mimeType) ? "text" : "binary";
}

export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return "0 B";
  if (bytes < 1024) return `${bytes} B`;

  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(kb >= 100 ? 0 : kb >= 10 ? 1 : 2)} KB`;

  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(mb >= 100 ? 0 : mb >= 10 ? 1 : 2)} MB`;

  const gb = mb / 1024;
  return `${gb.toFixed(gb >= 100 ? 0 : gb >= 10 ? 1 : 2)} GB`;
}
