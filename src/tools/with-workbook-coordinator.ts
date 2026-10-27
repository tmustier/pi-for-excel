/**
 * Tool wrapper that routes mutating tool calls through the workbook coordinator.
 */

import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import type { Static, TSchema } from "typebox";

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

/** When one of Pi's mutating tool calls held the workbook lock. */
export interface WorkbookWriteWindow {
  toolCallId: string;
  toolName: string;
  startedAt: number;
  /** `Infinity` while the call runs; set when it settles. */
  settledAt: number;
}

export interface WorkbookMutationObserver {
  onWriteCommitted?: (event: WorkbookMutationEvent) => void;
  /**
   * Every mutating call as it starts, committed or not (for telling Pi's edits
   * from the user's). The wrapper sets `settledAt` on the same object when the
   * call settles.
   */
  onWriteStarted?: (window: WorkbookWriteWindow) => void;
}

export interface ToolApprovalRequest {
  executionMode: ExecutionMode;
  /** Why this call asks: Confirm mode, or an action that always asks (see execution-policy). */
  reason: "confirm_mode" | "always";
  toolName: string;
  params: unknown;
  /** Aborts when the turn is stopped; the request should then be withdrawn. */
  signal?: AbortSignal;
  /** Existing content the change replaces, when the tool could tell. */
  overwrites?: OverwrittenCells;
}

/**
 * A tool may check, before approval is asked, whether the call would be
 * refused anyway (write_cells: invalid input or overwrite protection). A
 * result ends the call: the user is not asked about a change that will not
 * happen. It runs as a workbook read; the tool checks again when it runs.
 */
export interface ApprovalPrecheck<TParameters extends TSchema, TDetails> {
  checkBeforeApproval?: (
    toolCallId: string,
    params: Static<TParameters>,
  ) => Promise<ApprovalPrecheckResult<TDetails>>;
}

/** Non-empty cells a change would replace, so the approval can name them. */
export interface OverwrittenCells {
  count: number;
  /** A few of them, in reading order: address and current value or formula. */
  sample: Array<{ address: string; content: string }>;
}

/**
 * `refused`: the call ends with this result and nothing is asked.
 * `overwrites`: the call goes to approval, which names these cells.
 */
export type ApprovalPrecheckResult<TDetails> =
  | { refused: AgentToolResult<TDetails> }
  | { overwrites: OverwrittenCells }
  | null;

export interface WorkbookExecutionPolicy {
  /** Default: yolo */
  getExecutionMode?: () => Promise<ExecutionMode>;
  /**
   * Asked before the call waits for the workbook lock, for calls that
   * `getToolApprovalRequirement` says need approval in the current mode.
   * Return `true` to proceed, `false` to decline.
   */
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

  // A mutation stopped here never reached Excel; say so, so the agent knows.
  throw new Error(mutation ? "Stopped before running; nothing changed in the workbook." : "Aborted");
}

/**
 * A mutation tool that checks Stop itself and says in its result what was
 * written before it stopped (write_cells, fill_formula). Other mutation tools
 * run as one Excel batch; when Stop comes while one runs, the wrapper adds
 * that its change stands.
 */
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

/**
 * After Stop, how long a tool call may take to finish by itself before the
 * wrapper gives up on it. A write that finishes reports precisely what it
 * wrote (and its backup), so it gets long enough for a slow Excel on the web
 * batch; a hung Excel still ends the turn after this.
 */
export const STOP_GRACE_MS: Readonly<Record<"read" | "write", number>> = { read: 500, write: 10_000 };

/**
 * Settle a tool call promptly after Stop, even when Excel never answers, so
 * the workbook lock is released and the turn can end. The abandoned call may
 * still complete in Excel; the error says so.
 */
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

function defaultGetExecutionMode(): Promise<ExecutionMode> {
  return Promise.resolve("yolo");
}

function defaultRequestToolApproval(_request: ToolApprovalRequest): Promise<boolean> {
  return Promise.reject(new Error(
    "This action requires explicit user approval, but confirmation UI is unavailable.",
  ));
}

/** The approval this call needs in the current mode, or null. */
async function getNeededApproval(args: {
  policy: WorkbookExecutionPolicy;
  toolName: string;
  params: unknown;
}): Promise<{ requirement: "confirm_mode" | "always"; executionMode: ExecutionMode } | null> {
  const { policy, toolName, params } = args;
  const requirement = getToolApprovalRequirement(toolName, params);
  if (requirement === "none") return null;

  const getExecutionMode = policy.getExecutionMode ?? defaultGetExecutionMode;
  const executionMode = await getExecutionMode();
  if (requirement === "confirm_mode" && executionMode !== "safe") return null;
  return { requirement, executionMode };
}

async function requireApproval(args: {
  policy: WorkbookExecutionPolicy;
  requirement: "confirm_mode" | "always";
  executionMode: ExecutionMode;
  toolName: string;
  params: unknown;
  signal: AbortSignal | undefined;
  overwrites: OverwrittenCells | undefined;
}): Promise<void> {
  const { requirement, executionMode } = args;
  const requestToolApproval = args.policy.requestToolApproval ?? defaultRequestToolApproval;
  const approved = await requestToolApproval({
    executionMode,
    reason: requirement,
    toolName: args.toolName,
    params: args.params,
    ...(args.signal ? { signal: args.signal } : {}),
    ...(args.overwrites ? { overwrites: args.overwrites } : {}),
  });

  throwIfAborted(args.signal);
  if (!approved) {
    throw new Error(requirement === "always"
      ? "The user declined this action in the approval dialog."
      : "The user declined this workbook change in the approval dialog (Confirm mode).");
  }
}

/**
 * Mutations started earlier in this tool set that have not finished. A model can
 * issue a write and a verifying read in one batch; the batch's calls start in
 * order but run concurrently, so without this a read can see the value from
 * before the write. Reads wait for mutations already started; mutations do not
 * wait for reads. Registration happens before the first `await` in `execute`,
 * so the order of the batch decides who waits.
 */
type PendingMutations = Set<Promise<void>>;

function wrapTool<TParameters extends TSchema, TDetails>(
  tool: AgentTool<TParameters, TDetails> & ApprovalPrecheck<TParameters, TDetails>,
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
      const contextWorkbookId = await contextProvider.getWorkbookId();
      const coordinatorWorkbookId = contextWorkbookId ?? "workbook:unknown";
      const sessionId = contextProvider.getSessionId();
      const context = makeContext({
        workbookId: coordinatorWorkbookId,
        sessionId,
        toolName: tool.name,
      });

      throwIfAborted(signal);

      // Ask before waiting for the lock, so a pending approval never holds
      // up other tabs' writes.
      const needed = await getNeededApproval({ policy: executionPolicy, toolName: tool.name, params });
      if (needed) {
        const { checkBeforeApproval } = tool;
        const checked = checkBeforeApproval
          ? await coordinator.runRead(
            makeContext({ workbookId: coordinatorWorkbookId, sessionId, toolName: tool.name }),
            () => checkBeforeApproval(toolCallId, params),
          )
          : null;
        if (checked && "refused" in checked) return checked.refused;

        await requireApproval({
          policy: executionPolicy,
          ...needed,
          toolName: tool.name,
          params,
          signal,
          overwrites: checked?.overwrites,
        });
      }

      if (mode === "read") {
        const earlierMutations = [...pendingMutations];
        if (earlierMutations.length === 0) return run(mode, toolCallId, params, signal, onUpdate);
        return Promise.all(earlierMutations).then(() => run(mode, toolCallId, params, signal, onUpdate));
      }

      const out = await coordinator.runWrite(
        context,
        () => {
          throwIfAborted(signal);
          return tool.execute(toolCallId, params, signal, onUpdate);
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

      return out.result;
    },
  };

  async function run(
    mode: ReturnType<typeof getToolExecutionMode>,
    toolCallId: string,
    params: Static<TParameters>,
    signal: AbortSignal | undefined,
    onUpdate: AgentToolUpdateCallback<TDetails> | undefined,
  ): Promise<AgentToolResult<TDetails>> {
    const contextWorkbookId = await contextProvider.getWorkbookId();
    const coordinatorWorkbookId = contextWorkbookId ?? "workbook:unknown";
    const sessionId = contextProvider.getSessionId();
    const context = makeContext({
      workbookId: coordinatorWorkbookId,
      sessionId,
      toolName: tool.name,
    });

    throwIfAborted(signal, mode !== "read");

    if (mode === "read") {
      return coordinator.runRead(context, () => {
        throwIfAborted(signal);
        return settleOnStop(tool.execute(toolCallId, params, signal, onUpdate), signal, "read");
      });
    }

    await requireMutationApprovalIfNeeded({
      policy: executionPolicy,
      toolName: tool.name,
      params,
    });

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
          // A tool that reports its own Stop result is not cut short (#792 N-4):
          // its per-batch excelRun timeout bounds how long it takes.
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
