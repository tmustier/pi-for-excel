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

const YAML_MAPPING_RE = /^\s*([A-Za-z_][A-Za-z0-9_.-]*|"(?:[^"\\]|\\.)+"|'(?:[^'\\]|\\.)+')\s*:\s*(.*)$/u;

function countLeadingSpaces(text: string): number {
  return text.length - text.trimStart().length;
}

function normalizeFrontmatterKey(rawKey: string): string {
  const key = rawKey.trim();
  if ((key.startsWith("\"") && key.endsWith("\"")) || (key.startsWith("'") && key.endsWith("'"))) {
    return key.slice(1, -1);
  }

  return key;
}

function isLikelyYamlFrontmatterBlock(block: string): boolean {
  let sawMapping = false;
  let blockScalarIndent: number | null = null;

  for (const rawLine of block.split(/\r?\n/)) {
    const line = rawLine.trim();
    const lineIndent = countLeadingSpaces(rawLine);

    if (blockScalarIndent !== null) {
      if (!line) {
        continue;
      }

      if (lineIndent > blockScalarIndent) {
        continue;
      }

      blockScalarIndent = null;
    }

    if (!line) continue;
    if (line.startsWith("#")) continue;

    if (/^-\s+\S/u.test(line)) {
      continue;
    }

    const mappingMatch = YAML_MAPPING_RE.exec(rawLine);
    if (mappingMatch) {
      const rawKey = mappingMatch[1];
      if (rawKey === undefined) {
        return false;
      }

      const key = normalizeFrontmatterKey(rawKey);
      const normalizedKey = key.toLowerCase();

      if (key !== normalizedKey && !COMMON_FRONTMATTER_KEYS.has(normalizedKey)) {
        return false;
      }

      sawMapping = true;

      const value = (mappingMatch[2] ?? "").trim();
      if (/^[>|][+-]?\d*$/u.test(value)) {
        blockScalarIndent = lineIndent;
      }

      continue;
    }

    if (lineIndent > 0 && sawMapping) {
      continue;
    }

    return false;
  }

  return sawMapping;
}

export function stripYamlFrontmatter(text: string): string {
  const match = text.match(/^\uFEFF?---[ \t]*\r?\n([\s\S]*?)\r?\n---[ \t]*(?:\r?\n|$)/u);
  if (!match) return text;

  const frontmatterBody = match[1];
  if (frontmatterBody === undefined || !isLikelyYamlFrontmatterBlock(frontmatterBody)) return text;

  return text.slice(match[0].length);
}
