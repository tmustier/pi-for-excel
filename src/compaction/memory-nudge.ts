import type { AgentMessage } from "@mariozechner/pi-agent-core";

import { extractTextBlocks } from "../utils/content.js";

const AUTO_CONTEXT_PREFIX = "[Auto-context]";
const DEFAULT_MAX_SNIPPETS = 3;
const MAX_SNIPPET_CHARS = 180;

const MEMORY_CUE_PATTERNS: readonly RegExp[] = [
  /\bremember(?:\s+this|\s+that)?\b/i,
  /\bdon['’]t\s+forget\b/i,
  /\bkeep\s+(?:this|that)\s+in\s+mind\b/i,
  /\bfor\s+future\s+reference\b/i,
  /\bmake\s+a\s+note\b/i,
  /\bplease\s+save\b/i,
];

export interface CompactionMemoryCueSummary {
  cueCount: number;
  snippets: string[];
}

function normalizeWhitespace(text: string): string {
  return text.replace(/\s+/gu, " ").trim();
}

function isMemoryCueText(text: string): boolean {
  return MEMORY_CUE_PATTERNS.some((pattern) => pattern.test(text));
}

function toSnippet(text: string): string {
  if (text.length <= MAX_SNIPPET_CHARS) return text;
  return `${text.slice(0, MAX_SNIPPET_CHARS - 1).trimEnd()}…`;
}

function messageText(message: AgentMessage): string {
  if (message.role !== "user" && message.role !== "user-with-attachments") {
    return "";
  }

  if (typeof message.content === "string") {
    return normalizeWhitespace(message.content);
  }

  return normalizeWhitespace(extractTextBlocks(message.content));
}

export function collectCompactionMemoryCues(
  messages: readonly AgentMessage[],
  maxSnippets = DEFAULT_MAX_SNIPPETS,
): CompactionMemoryCueSummary {
  const safeMaxSnippets = Math.max(0, maxSnippets);
  const snippets: string[] = [];
  const seenSnippetKeys = new Set<string>();
  let cueCount = 0;

  for (const message of messages) {
    const text = messageText(message);
    if (!text) continue;
    if (text.startsWith(AUTO_CONTEXT_PREFIX)) continue;
    if (!isMemoryCueText(text)) continue;

    cueCount += 1;

    if (snippets.length >= safeMaxSnippets) continue;

    const snippet = toSnippet(text);
    const snippetKey = snippet.toLowerCase();
    if (seenSnippetKeys.has(snippetKey)) continue;

    seenSnippetKeys.add(snippetKey);
    snippets.push(snippet);
  }

  return {
    cueCount,
    snippets,
  };
}

export function buildCompactionMemoryFocusInstruction(
  summary: CompactionMemoryCueSummary,
): string | null {
  if (summary.cueCount <= 0) return null;

  const lines: string[] = [
    "Before finalizing the summary, identify durable memory that should be written to workspace files.",
    "- Behavioral preferences/rules should be called out for the instructions tool.",
    "- Factual workbook/domain memory should be called out for notes/ or workbooks/<name>/notes.md.",
    "- If durable memory is present, include a short bullet list in Critical Context under 'Memory to persist'.",
  ];

  if (summary.snippets.length > 0) {
    lines.push("Potential user cues:");
    for (const snippet of summary.snippets) {
      lines.push(`- \"${snippet}\"`);
    }
  }

  return lines.join("\n");
}

export function mergeCompactionAdditionalFocus(
  ...focusParts: Array<string | null | undefined>
): string | undefined {
  const normalized = focusParts
    .map((part) => (typeof part === "string" ? part.trim() : ""))
    .filter((part) => part.length > 0);

  if (normalized.length === 0) {
    return undefined;
  }

  return normalized.join("\n\n");
}
