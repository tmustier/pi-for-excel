/**
 * Tool wrapper that routes mutating tool calls through the workbook coordinator.
 */

import type { AgentTool, AgentToolResult, AgentToolUpdateCallback } from "@earendil-works/pi-agent-core";
import type { Static, TSchema } from "typebox";

import {
  formatExecutionModeLabel,
  type ExecutionMode,
} from "../execution/mode.js";
import type { WorkbookCoordinator, WorkbookOperationContext } from "../workbook/coordinator.js";
import { cancelCellEditWaits } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import { getToolContextImpact, getToolExecutionMode, type ToolContextImpact } from "./execution-policy.js";
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

export interface MutationApprovalRequest {
  executionMode: ExecutionMode;
  toolName: string;
  params: unknown;
}

export interface WorkbookExecutionPolicy {
  /** Default: yolo */
  getExecutionMode?: () => Promise<ExecutionMode>;
  /**
   * Called for mutating tools when execution mode is "safe".
   * Return `true` to proceed, `false` to block.
   */
  requestMutationApproval?: (request: MutationApprovalRequest) => Promise<boolean>;
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
 * After Stop, how long a tool call may take to finish by itself (and report
 * precisely what it did) before the wrapper gives up on it.
 */
export const STOP_GRACE_MS = 500;

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
      }, STOP_GRACE_MS);
    };
    if (signal.aborted) onAbort();
    else signal.addEventListener("abort", onAbort, { once: true });
    work.then(resolve, reject).finally(() => {
      if (timer !== undefined) clearTimeout(timer);
      signal.removeEventListener("abort", onAbort);
    });
  });
}

function isWorkbookCoordinatorParamsShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function readApprovalParams(params: unknown): { action: string | null; range: string | null } {
  if (!isWorkbookCoordinatorParamsShape(params)) return { action: null, range: null };

  const action = typeof params.action === "string" ? params.action.trim() : "";
  const range = typeof params.range === "string" ? params.range.trim() : "";
  return { action: action.length > 0 ? action : null, range: range.length > 0 ? range : null };
}

export function buildMutationApprovalMessage(request: MutationApprovalRequest): string {
  const modeLabel = formatExecutionModeLabel(request.executionMode);
  const { action, range } = readApprovalParams(request.params);

  const toolLabel = action ? `${request.toolName}:${action}` : request.toolName;
  const lines = [
    `Allow workbook mutation in ${modeLabel} mode?`,
    "",
    `Tool: ${toolLabel}`,
  ];

  if (range) {
    lines.push(`Range: ${range}`);
  }

  lines.push("", "Tip: run /yolo on to switch to Auto mode (no confirmations).");
  return lines.join("\n");
}

function defaultGetExecutionMode(): Promise<ExecutionMode> {
  return Promise.resolve("yolo");
}

function defaultRequestMutationApproval(_request: MutationApprovalRequest): Promise<boolean> {
  return Promise.reject(new Error(
    "Confirm mode requires explicit user approval, but confirmation UI is unavailable.",
  ));
}

async function requireMutationApprovalIfNeeded(args: {
  policy: WorkbookExecutionPolicy;
  toolName: string;
  params: unknown;
}): Promise<void> {
  const getExecutionMode = args.policy.getExecutionMode ?? defaultGetExecutionMode;
  const executionMode = await getExecutionMode();

  if (executionMode !== "safe") {
    return;
  }

  const requestMutationApproval = args.policy.requestMutationApproval ?? defaultRequestMutationApproval;
  const approved = await requestMutationApproval({
    executionMode,
    toolName: args.toolName,
    params: args.params,
  });

  if (!approved) {
    throw new Error("Mutation cancelled by user (Confirm mode).");
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
  tool: AgentTool<TParameters, TDetails> & StopReporting,
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

      throwIfAborted(signal);

      const out = await coordinator.runWrite(
        context,
        async () => {
          throwIfAborted(signal);
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
            return await settleOnStop(tool.execute(toolCallId, params, signal, onUpdate), signal, "write");
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
        return tool.execute(toolCallId, params, signal, onUpdate);
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
      () => {
        throwIfAborted(signal, true);
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
