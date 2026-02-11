/**
 * Security policy for extension module sources.
 *
 * Keep this intentionally small:
 * - local module specifiers are allowed by default
 * - blob: module URLs are allowed (used for paste-code extensions)
 * - remote http(s) URLs are blocked by default
 * - an explicit localStorage opt-in can temporarily re-enable remote URLs
 */

const LOCAL_SPECIFIER_PREFIXES = ["./", "../", "/"];
const REMOTE_PROTOCOLS = new Set(["http:", "https:"]);
const BLOB_PROTOCOL = "blob:";

export const ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY = "pi.allowRemoteExtensionUrls";

export type ExtensionSourceKind = "local-module" | "blob-url" | "remote-url" | "unsupported";

/**
 * Classify a string extension source into local/remote/unsupported.
 */
export function classifyExtensionSource(source: string): ExtensionSourceKind {
  const specifier = source.trim();
  if (specifier.length === 0) return "unsupported";

  // `//host/path` is protocol-relative and resolves to a remote module URL.
  if (specifier.startsWith("//")) return "remote-url";

  for (const prefix of LOCAL_SPECIFIER_PREFIXES) {
    if (specifier.startsWith(prefix)) return "local-module";
  }

  let parsed: URL;
  try {
    parsed = new URL(specifier);
  } catch {
    return "unsupported";
  }

  if (parsed.protocol === BLOB_PROTOCOL) {
    return "blob-url";
  }

  return REMOTE_PROTOCOLS.has(parsed.protocol) ? "remote-url" : "unsupported";
}

/**
 * Parse an explicit unsafe opt-in flag for remote extension URLs.
 */
export function isRemoteExtensionOptIn(raw: string | null | undefined): boolean {
  return raw === "1" || raw === "true";
}
