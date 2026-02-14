/**
 * Markdown preprocessing helpers for UI rendering.
 */

const COMMON_FRONTMATTER_KEYS = new Set([
  "title",
  "date",
  "description",
  "name",
  "author",
  "authors",
  "slug",
  "tags",
  "category",
  "categories",
  "layout",
  "draft",
  "published",
  "updated",
  "summary",
  "excerpt",
]);

function isLikelyYamlFrontmatterBlock(block: string): boolean {
  let sawMapping = false;

  for (const rawLine of block.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith("#")) continue;

    // YAML mapping entry (supports dashed keys and optional value).
    const mappingMatch = line.match(/^([A-Za-z_][A-Za-z0-9_-]*)\s*:\s*(?:.*)?$/);
    if (mappingMatch) {
      const key = mappingMatch[1];
      const normalizedKey = key.toLowerCase();
      // Keep prose guard: allow arbitrary lowercase keys, but only allow
      // title-case/uppercase keys for canonical frontmatter fields.
      if (key === normalizedKey || COMMON_FRONTMATTER_KEYS.has(normalizedKey)) {
        sawMapping = true;
        continue;
      }
      return false;
    }

    // YAML list item (e.g., under a mapping key).
    if (/^-\s+\S/.test(line)) {
      continue;
    }

    // Anything else is likely markdown/text, not frontmatter metadata.
    return false;
  }

  return sawMapping;
}

/**
 * Remove YAML frontmatter only when the opening block is likely metadata.
 *
 * This avoids dropping ordinary markdown that happens to begin with
 * thematic breaks (`---`).
 */
export function stripYamlFrontmatter(text: string): string {
  const match = text.match(/^---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*(?:\r?\n|$)/);
  if (!match) return text;

  const frontmatterBody = match[1];
  if (!isLikelyYamlFrontmatterBlock(frontmatterBody)) return text;

  return text.slice(match[0].length);
}
