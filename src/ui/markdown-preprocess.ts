/**
 * Markdown preprocessing helpers for UI rendering.
 */

function isLikelyYamlFrontmatterBlock(block: string): boolean {
  let sawMapping = false;

  for (const rawLine of block.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith("#")) continue;

    // YAML mapping entry (supports dashed keys and optional value).
    if (/^[a-z_][a-z0-9_-]*\s*:\s*(?:.*)?$/.test(line)) {
      sawMapping = true;
      continue;
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
