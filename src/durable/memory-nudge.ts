import type { Message } from "@earendil-works/pi-ai";
import type { CompactionHooks } from "@earendil-works/pi-durable";

const MEMORY_CUES = [
  /\bremember(?:\s+this|\s+that)?\b/iu,
  /\bdon['’]t\s+forget\b/iu,
  /\bkeep\s+(?:this|that)\s+in\s+mind\b/iu,
  /\bfor\s+future\s+reference\b/iu,
  /\bmake\s+a\s+note\b/iu,
  /\bplease\s+save\b/iu,
];

function userText(message: Message): string {
  if (message.role !== "user") return "";
  const text = typeof message.content === "string" ? message.content : message.content
    .filter((block) => block.type === "text").map((block) => block.text).join("\n");
  return text.replace(/\s+/gu, " ").trim();
}

/** Pure request-only hook: never edits history or replaces native summary placement. */
export const compactionMemoryNudge: CompactionHooks["beforeSummary"] = (request) => {
  const messages = request.messages.filter((message) => !userText(message).startsWith("[Auto-context]"));
  const snippets: string[] = [];
  const seen = new Set<string>();
  let hasCue = false;
  for (const message of messages) {
    const text = userText(message);
    if (!MEMORY_CUES.some((pattern) => pattern.test(text))) continue;
    hasCue = true;
    const snippet = text.length <= 180 ? text : `${text.slice(0, 179).trimEnd()}…`;
    const key = snippet.toLowerCase();
    if (snippets.length >= 3 || seen.has(key)) continue;
    seen.add(key);
    snippets.push(snippet);
  }
  const focus = hasCue ? [
    "Before finalizing the summary, identify durable memory that should be written to workspace files.",
    "- Behavioral preferences/rules should be called out for the instructions tool.",
    "- Factual workbook/domain memory should be called out for notes/ or workbooks/<name>/notes.md.",
    "- If durable memory is present, include a short bullet list in Critical Context under 'Memory to persist'.",
    "Potential user cues:",
    ...snippets.map((snippet) => `- "${snippet}"`),
  ].join("\n") : undefined;
  const instructions = [request.instructions?.trim(), focus].filter(Boolean).join("\n\n");
  return { messages, ...(instructions ? { instructions } : {}) };
};
