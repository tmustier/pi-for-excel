import type { AgentToolResult } from "@earendil-works/pi-agent-core";

export function appendMutationResultNote<TDetails>(result: AgentToolResult<TDetails>, note: string): void {
  const trimmedNote = note.trim();
  if (trimmedNote.length === 0) return;

  const first = result.content[0];
  if (!first || first.type !== "text") return;

  first.text = `${first.text}\n\n${trimmedNote}`;
}

export function appendMutationResultNotes<TDetails>(
  result: AgentToolResult<TDetails>,
  notes: readonly string[],
): void {
  for (const note of notes) {
    appendMutationResultNote(result, note);
  }
}
