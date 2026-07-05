/**
 * Host-neutral workbook identity primitives.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never a raw
 * host document URL/path. Host adapters pass any raw identity source through
 * this module so the privacy-preserving semantics stay centralized.
 */

import { t } from "../language/index.js";

export interface WorkbookContext {
  /**
   * A stable, local-only identifier when available.
   *
   * Current Office strategy:
   * - if the host document URL exists, normalize it (drop query/hash) and return a SHA-256 hash
   * - otherwise return null (ephemeral/unknown)
   */
  workbookId: string | null;

  /** Best-effort workbook file name derived from the host document URL/path. */
  workbookName: string | null;

  /** Where the identity came from (useful for debugging / future migration). */
  source: "document.url" | "unknown";
}

export function createUnknownWorkbookContext(): WorkbookContext {
  return {
    workbookId: null,
    workbookName: null,
    source: "unknown",
  };
}

function bufferToHex(buf: ArrayBuffer): string {
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function stripQueryAndHash(raw: string): string {
  return raw.split(/[?#]/, 1)[0]?.trim() ?? "";
}

function getWorkbookNameFromUrl(url: string): string | null {
  const readName = (raw: string): string | null => {
    const cleaned = stripQueryAndHash(raw);
    if (!cleaned) return null;

    try {
      const decoded = decodeURIComponent(cleaned).trim();
      return decoded.length > 0 ? decoded : null;
    } catch {
      return cleaned.length > 0 ? cleaned : null;
    }
  };

  try {
    const parsed = new URL(url);
    const normalizedPathname = parsed.pathname.replace(/\\/g, "/");
    const fromPathname = normalizedPathname.split("/").at(-1);
    if (fromPathname) {
      const name = readName(fromPathname);
      if (name) return name;
    }
  } catch {
    // Fall back to plain path parsing below.
  }

  const normalized = url.replace(/\\/g, "/");
  const fromPath = normalized.split("/").at(-1);
  return fromPath ? readName(fromPath) : null;
}

function normalizeWorkbookIdentitySource(url: string): string {
  const trimmed = url.trim();
  if (trimmed.length === 0) return trimmed;

  try {
    const parsed = new URL(trimmed);
    parsed.search = "";
    parsed.hash = "";
    parsed.pathname = parsed.pathname.replace(/\\/g, "/");
    return parsed.toString();
  } catch {
    return stripQueryAndHash(trimmed).replace(/\\/g, "/");
  }
}

export function formatWorkbookLabel(context: WorkbookContext): string {
  if (context.workbookName) return context.workbookName;

  if (context.workbookId) {
    const shortId = context.workbookId.slice(0, 18);
    return `Workbook (${shortId}…)`;
  }

  return t("workbook.currentWorkbook");
}

function fnv1a32Hex(bytes: Uint8Array): string {
  let hash = 0x811c9dc5;
  for (const b of bytes) {
    hash ^= b;
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

async function sha256Hex(input: string): Promise<string> {
  const bytes = new TextEncoder().encode(input);

  // WebCrypto SHA-256 (preferred)
  const subtle = globalThis.crypto?.subtle;
  if (subtle?.digest) {
    try {
      const buf = await subtle.digest("SHA-256", bytes);
      return bufferToHex(buf);
    } catch {
      // fall through to FNV-1a
    }
  }

  // Fallback: non-cryptographic but stable hash. Only used when WebCrypto is unavailable.
  return fnv1a32Hex(bytes);
}

/**
 * Best-effort workbook context from a host-provided document URL/path.
 *
 * IMPORTANT: callers should persist only `workbookId` (the hash), never the raw URL/path.
 */
export async function getWorkbookContextFromDocumentUrl(url: string | null): Promise<WorkbookContext> {
  if (!url) {
    return createUnknownWorkbookContext();
  }

  const identitySource = normalizeWorkbookIdentitySource(url);
  const hash = await sha256Hex(identitySource);
  return {
    workbookId: `url_sha256:${hash}`,
    workbookName: getWorkbookNameFromUrl(url),
    source: "document.url",
  };
}
