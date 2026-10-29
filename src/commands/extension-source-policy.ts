const LOCAL_SPECIFIER_PREFIXES = ["./", "../", "/"];
const REMOTE_PROTOCOLS = new Set(["http:", "https:"]);
const BLOB_PROTOCOL = "blob:";

export const ALLOW_REMOTE_EXTENSION_URLS_STORAGE_KEY = "pi.allowRemoteExtensionUrls";

export type ExtensionSourceKind = "local-module" | "blob-url" | "remote-url" | "unsupported";

export function classifyExtensionSource(source: string): ExtensionSourceKind {
  const specifier = source.trim();
  if (specifier.length === 0) return "unsupported";

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

export function isRemoteExtensionOptIn(raw: string | null | undefined): boolean {
  return raw === "1" || raw === "true";
}
