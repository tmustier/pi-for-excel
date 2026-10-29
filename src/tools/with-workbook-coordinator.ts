import type { AgentTool, AgentToolResult, AgentToolUpdateCallback } from "@earendil-works/pi-agent-core";
import type { Static, TSchema } from "typebox";
import { Value } from "typebox/value";

import type { ExecutionMode } from "../execution/mode.js";
import type { WorkbookCoordinator, WorkbookOperationContext } from "../workbook/coordinator.js";
import { cancelCellEditWaits } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import {
  getToolApprovalRequirement,
  getToolContextImpact,
  getToolExecutionMode,
  type ToolContextImpact,
} from "./execution-policy.js";
import { isUnsupportedHostTool } from "./unsupported-host-tool.js";
import { mutationTargetSchema, requireExplicitMutationTarget } from "./mutation/target.js";

export interface WorkbookCoordinatorContextProvider {
  getWorkbookId: () => Promise<string | null>;
  getSessionId: () => string;
}

export interface WorkbookMutationEvent {
  workbookId: string | null;
  sessionId: string;
  toolName: string;
  impact: ToolContextImpact;
  revision: number;
}

export interface WorkbookWriteWindow {
  toolCallId: string;
  toolName: string;
  startedAt: number;
  settledAt: number;
}

export interface WorkbookMutationObserver {
  onWriteCommitted?: (event: WorkbookMutationEvent) => void;
  onWriteStarted?: (window: WorkbookWriteWindow) => void;
}

export interface ToolApprovalRequest {
  toolCallId: string;
  executionMode: ExecutionMode;
  reason: "confirm_mode" | "always" | "irreversible";
  toolName: string;
  params: unknown;
  signal?: AbortSignal;
  overwrites?: OverwrittenCells;
  unbackableCellCount?: number;
}

export interface ApprovalPrecheck<TParameters extends TSchema, TDetails> {
  alwaysCheckBeforeApproval?: boolean;
  checkBeforeApproval?: (
    toolCallId: string,
    params: Static<TParameters>,
  ) => Promise<ApprovalPrecheckResult<TDetails>>;
}

export interface OverwrittenCells {
  count: number;
  sample: Array<{ address: string; content: string }>;
}

export type ApprovalPrecheckResult<TDetails> =
  | { refused: AgentToolResult<TDetails> }
  | { overwrites: OverwrittenCells }
  | { unbackableCellCount: number }
  | null;

export interface WorkbookExecutionPolicy {
  getExecutionMode?: () => Promise<ExecutionMode>;
  requestToolApproval?: (request: ToolApprovalRequest) => Promise<boolean>;
}

function makeContext(args: {
  workbookId: string;
  sessionId: string;
  toolName: string;
}): WorkbookOperationContext {
  return {
    workbookId: args.workbookId,
    sessionId: args.sessionId,
    opId: crypto.randomUUID(),
    toolName: args.toolName,
  };
}

function throwIfAborted(signal: AbortSignal | undefined, mutation = false): void {
  if (!signal?.aborted) return;

  throw new Error(mutation ? "Stopped before running; nothing changed in the workbook." : "Aborted");
}

export interface StopReporting {
  reportsStop?: true;
}

export const STOPPED_AFTER_CHANGE_NOTE =
  "⏹ Stop was pressed while this ran; the change above was already applied and stands.";

function appendStopNote<TDetails>(result: AgentToolResult<TDetails>): AgentToolResult<TDetails> {
  const [first, ...rest] = result.content;
  if (!first || first.type !== "text") return result;
  return { ...result, content: [{ ...first, text: `${first.text}\n\n${STOPPED_AFTER_CHANGE_NOTE}` }, ...rest] };
}

export const STOP_GRACE_MS: Readonly<Record<"read" | "write", number>> = { read: 500, write: 10_000 };

function settleOnStop<T>(
  work: Promise<T>,
  signal: AbortSignal | undefined,
  mode: "read" | "write",
): Promise<T> {
  if (!signal) return work;
  return new Promise<T>((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout> | undefined;
    const onAbort = () => {
      cancelCellEditWaits();
      timer = setTimeout(() => {
        work.catch(() => {});
        reject(mode === "write"
          ? Object.assign(new Error("Stopped by the user while Excel was still working on this change. The outcome is unknown: the change may still be applied. Read the affected range back before relying on it."), { outcomeUnknown: true })
          : new Error("Stopped by the user before Excel answered."));
      }, STOP_GRACE_MS[mode]);
    };
    if (signal.aborted) onAbort();
    else signal.addEventListener("abort", onAbort, { once: true });
    work.then(resolve, reject).finally(() => {
      if (timer !== undefined) clearTimeout(timer);
      signal.removeEventListener("abort", onAbort);
    });
  });
}

type PendingMutations = Set<Promise<void>>;

function wrapTool<TParameters extends TSchema, TDetails>(
  tool: AgentTool<TParameters, TDetails> & ApprovalPrecheck<TParameters, TDetails> & StopReporting,
  coordinator: WorkbookCoordinator,
  contextProvider: WorkbookCoordinatorContextProvider,
  mutationObserver: WorkbookMutationObserver | undefined,
  executionPolicy: WorkbookExecutionPolicy,
  pendingMutations: PendingMutations,
): AgentTool<TParameters, TDetails> {
  return {
    ...tool,
    execute: (toolCallId, params, signal, onUpdate) => {
      const mode = getToolExecutionMode(tool.name, params);
      if (mode === "read") {
        const earlierMutations = [...pendingMutations];
        if (earlierMutations.length === 0) return run(mode, toolCallId, params, signal, onUpdate);
        return Promise.all(earlierMutations).then(() => run(mode, toolCallId, params, signal, onUpdate));
      }

      const running = run(mode, toolCallId, params, signal, onUpdate);
      const settled = running.then(() => undefined, () => undefined);
      pendingMutations.add(settled);
      void settled.then(() => pendingMutations.delete(settled));
      return running;
    },
  };

  async function run(
    mode: ReturnType<typeof getToolExecutionMode>,
    toolCallId: string,
    params: Static<TParameters>,
    signal: AbortSignal | undefined,
    onUpdate: AgentToolUpdateCallback<TDetails> | undefined,
  ): Promise<AgentToolResult<TDetails>> {
    if (mode === "mutate" && Value.Check(mutationTargetSchema, params)) {
      requireExplicitMutationTarget(tool.name, params);
    }
    const contextWorkbookId = await contextProvider.getWorkbookId();
    const coordinatorWorkbookId = contextWorkbookId ?? "workbook:unknown";
    const sessionId = contextProvider.getSessionId();
    const context = makeContext({
      workbookId: coordinatorWorkbookId,
      sessionId,
      toolName: tool.name,
    });

    throwIfAborted(signal, mode !== "read");

    let requirement = getToolApprovalRequirement(tool.name, params);
    const executionMode = requirement === "none"
      ? "yolo"
      : await executionPolicy.getExecutionMode?.() ?? "yolo";
    if (tool.alwaysCheckBeforeApproval || requirement === "always" || requirement === "irreversible" || (requirement === "confirm_mode" && executionMode === "safe")) {
      const { checkBeforeApproval } = tool;
      const checked = checkBeforeApproval
        ? await coordinator.runRead(
          makeContext({ workbookId: coordinatorWorkbookId, sessionId, toolName: tool.name }),
          () => checkBeforeApproval(toolCallId, params),
        )
        : null;
      if (checked && "refused" in checked) return checked.refused;
      if (checked && "unbackableCellCount" in checked) requirement = "irreversible";
      if (requirement === "always" || requirement === "irreversible" || (requirement === "confirm_mode" && executionMode === "safe")) {
        if (!executionPolicy.requestToolApproval) {
          throw new Error("This action requires explicit user approval, but confirmation UI is unavailable.");
        }
        const approved = await executionPolicy.requestToolApproval({
          toolCallId,
          executionMode,
          reason: requirement,
          toolName: tool.name,
          params,
          ...(signal ? { signal } : {}),
          ...(checked && "overwrites" in checked ? { overwrites: checked.overwrites } : {}),
          ...(checked && "unbackableCellCount" in checked ? { unbackableCellCount: checked.unbackableCellCount } : {}),
        });
        throwIfAborted(signal);
        if (!approved) throw new Error(requirement === "confirm_mode"
          ? "The user declined this workbook change when asked (Confirm mode)."
          : "The user declined this action when asked.");
      }
    }

    if (mode === "read") {
      return coordinator.runRead(context, () => {
        throwIfAborted(signal);
        return settleOnStop(tool.execute(toolCallId, params, signal, onUpdate), signal, "read");
      });
    }

    throwIfAborted(signal, true);

    const out = await coordinator.runWrite(
      context,
      async () => {
        throwIfAborted(signal, true);
        const window: WorkbookWriteWindow = {
          toolCallId,
          toolName: tool.name,
          startedAt: Date.now(),
          settledAt: Number.POSITIVE_INFINITY,
        };
        try {
          mutationObserver?.onWriteStarted?.(window);
        } catch (error) {
          console.warn("[pi] Workbook write observer failed:", getErrorMessage(error));
        }
        try {
          const work = tool.execute(toolCallId, params, signal, onUpdate);
          return await (tool.reportsStop ? work : settleOnStop(work, signal, "write"));
        } finally {
          window.settledAt = Date.now();
        }
      },
    );

    if (mutationObserver?.onWriteCommitted) {
      const impact = getToolContextImpact(tool.name, params);
      try {
        mutationObserver.onWriteCommitted({
          workbookId: contextWorkbookId,
          sessionId,
          toolName: tool.name,
          impact,
          revision: out.revision,
        });
      } catch (error) {
        console.warn("[pi] Workbook mutation observer failed:", getErrorMessage(error));
      }
    }

    if (signal?.aborted && !tool.reportsStop && out.result.isError !== true) {
      return appendStopNote(out.result);
    }
    return out.result;
  }
}

export function withWorkbookCoordinator(
  tools: AgentTool[],
  coordinator: WorkbookCoordinator,
  contextProvider: WorkbookCoordinatorContextProvider,
  mutationObserver?: WorkbookMutationObserver,
  executionPolicy: WorkbookExecutionPolicy = {},
): AgentTool[] {
  const pendingMutations: PendingMutations = new Set();
  return tools.map((tool) => isUnsupportedHostTool(tool)
    ? tool
    : wrapTool(tool, coordinator, contextProvider, mutationObserver, executionPolicy, pendingMutations));
}
