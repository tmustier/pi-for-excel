/**
 * Pure policy helpers used by markdown safety patching.
 *
 * These helpers are kept framework-free so they can be regression-tested under
 * Node without browser rendering dependencies.
 */

const ALLOWED_LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);
const DISABLED_MARKDOWN_EXTENSION_NAMES = new Set(["inlineMathDollar", "blockMathDollar"]);

export type MarkdownImageRenderPlan =
  | { kind: "link"; href: string; label: string }
  | { kind: "text"; label: string };

function getBaseUrlForResolution(): string {
  if (typeof window !== "undefined" && typeof window.location?.href === "string") {
    return window.location.href;
  }
  return "https://localhost/";
}

/** Allowlist-based URL protocol check used for markdown links/images. */
export function isAllowedMarkdownUrl(raw: string): boolean {
  const trimmed = raw.trim();
  if (!trimmed) return false;

  try {
    const parsed = new URL(trimmed, getBaseUrlForResolution());
    return ALLOWED_LINK_PROTOCOLS.has(parsed.protocol);
  } catch {
    return false;
  }
}

/**
 * mini-lit enables `$...$` and `$$...$$` KaTeX extensions by default.
 * In spreadsheet/finance prose this collides with currency notation and
 * causes large chunks of text to render as math. We disable only the
 * dollar-delimited math extensions and keep `\(...\)` / `\[...\]` support.
 */
export function isMarkdownExtensionDisabledByPolicy(name: string): boolean {
  return DISABLED_MARKDOWN_EXTENSION_NAMES.has(name);
}

export function getMarkdownImageLabel(alt: string): string {
  const trimmedAlt = alt.trim();
  if (trimmedAlt.length > 0) {
    return `image: ${trimmedAlt}`;
  }
  return "image";
}

/**
 * Decide image fallback rendering. We never render markdown images as <img>.
 */
export function createMarkdownImageRenderPlan(href: string, alt: string): MarkdownImageRenderPlan {
  const label = getMarkdownImageLabel(alt);

  if (isAllowedMarkdownUrl(href)) {
    return {
      kind: "link",
      href,
      label,
    };
  }

  return {
    kind: "text",
    label,
  };
}
