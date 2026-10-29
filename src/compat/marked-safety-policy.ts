const ALLOWED_LINK_PROTOCOLS = new Set(["http:", "https:", "mailto:", "tel:"]);
const DISABLED_MARKDOWN_EXTENSION_NAMES = new Set([
  "inlineMathDollar",
  "blockMathDollar",
  "inlineMathLatex",
  "blockMathLatex",
]);

export type MarkdownImageRenderPlan =
  | { kind: "link"; href: string; label: string }
  | { kind: "text"; label: string };

function getBaseUrlForResolution(): string {
  if (typeof window !== "undefined" && typeof window.location?.href === "string") {
    return window.location.href;
  }
  return "https://localhost/";
}

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
