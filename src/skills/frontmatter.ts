/**
 * Agent Skill markdown/frontmatter parsing.
 */

export interface ParsedSkillFrontmatter {
  name: string;
  description: string;
  compatibility?: string;
}

export interface ParsedSkillDocument {
  frontmatter: ParsedSkillFrontmatter;
  body: string;
}

const FRONTMATTER_RE = /^---\s*\n([\s\S]*?)\n---\s*\n?/u;

function unquoteYamlScalar(value: string): string {
  const trimmed = value.trim();
  if (trimmed.length < 2) return trimmed;

  const startsWithQuote = trimmed.startsWith('"') || trimmed.startsWith("'");
  const endsWithQuote = trimmed.endsWith('"') || trimmed.endsWith("'");

  if (!startsWithQuote || !endsWithQuote) return trimmed;
  return trimmed.slice(1, -1).trim();
}

function parseTopLevelFrontmatter(frontmatterBlock: string): Record<string, string> {
  const values: Record<string, string> = {};

  for (const rawLine of frontmatterBlock.split(/\r?\n/u)) {
    const line = rawLine.trimEnd();
    if (line.length === 0) continue;
    if (line.startsWith("#")) continue;

    // Ignore nested YAML blocks (e.g. metadata:) and list items.
    if (rawLine.startsWith(" ") || rawLine.startsWith("\t") || line.startsWith("- ")) {
      continue;
    }

    const separatorIndex = line.indexOf(":");
    if (separatorIndex <= 0) continue;

    const key = line.slice(0, separatorIndex).trim().toLowerCase();
    if (key.length === 0) continue;

    const rawValue = line.slice(separatorIndex + 1).trim();
    values[key] = unquoteYamlScalar(rawValue);
  }

  return values;
}

export function parseSkillDocument(markdown: string): ParsedSkillDocument | null {
  const match = FRONTMATTER_RE.exec(markdown);
  if (!match) return null;

  const fullMatch = match[0];
  const frontmatterBlock = match[1];
  if (frontmatterBlock === undefined) return null;

  const frontmatterValues = parseTopLevelFrontmatter(frontmatterBlock);
  const name = frontmatterValues.name?.trim() ?? "";
  const description = frontmatterValues.description?.trim() ?? "";

  if (name.length === 0 || description.length === 0) {
    return null;
  }

  const compatibility = frontmatterValues.compatibility?.trim();
  const body = markdown.slice(fullMatch.length).trimStart();

  return {
    frontmatter: {
      name,
      description,
      ...(compatibility && compatibility.length > 0 ? { compatibility } : {}),
    },
    body,
  };
}
