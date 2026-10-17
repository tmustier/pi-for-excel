import type { JsonValue } from "@earendil-works/chord";
import type { AgentToolResult } from "@earendil-works/pi-agent-core";

export interface WorkbookMutationCall {
  /** Stable Durable logical invocation ID, scoped by conversation/task; never the provider call ID alone. */
  mutationId: string;
  /** Provider-owned correlation ID; may repeat across tasks and sessions. */
  toolCallId: string;
  toolName: string;
  workbookId: string | null;
  sessionId: string;
  params: JsonValue;
}

/** Entered under the workbook coordinator's write lock, before calling any host code. */
export interface WorkbookMutationGate {
  execute<T>(call: WorkbookMutationCall, run: () => Promise<AgentToolResult<T>>): Promise<AgentToolResult<T>>;
}
