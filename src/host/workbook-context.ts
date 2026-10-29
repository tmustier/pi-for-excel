import { t } from "../language/index.js";

export interface WorkbookContext {
  workbookId: string | null;

  workbookName: string | null;

  source: "document.url" | "wps.activeWorkbook.fullName" | "unknown";
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

  const subtle = globalThis.crypto?.subtle;
  if (subtle?.digest) {
    try {
      const buf = await subtle.digest("SHA-256", bytes);
      return bufferToHex(buf);
    } catch {
    }
  }

  return fnv1a32Hex(bytes);
}

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

export async function getWorkbookContextFromWpsFullName(
  fullName: string | null,
  workbookName: string | null,
): Promise<WorkbookContext> {
  const trimmedFullName = fullName?.trim() ?? "";
  if (!trimmedFullName) {
    return {
      workbookId: null,
      workbookName,
      source: "unknown",
    };
  }

  const identitySource = normalizeWorkbookIdentitySource(trimmedFullName);
  const hash = await sha256Hex(identitySource);
  return {
    workbookId: `wps_path_sha256:${hash}`,
    workbookName,
    source: "wps.activeWorkbook.fullName",
  };
}
