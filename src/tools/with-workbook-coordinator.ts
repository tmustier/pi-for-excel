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

export interface WorkbookMutationObserver {
  onWriteCommitted?: (event: WorkbookMutationEvent) => void;
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

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) return;

  // Keep message text compatible with existing abort handling paths.
  throw new Error("Aborted");
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
  tool: AgentTool<TParameters, TDetails>,
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
    const contextWorkbookId = await contextProvider.getWorkbookId();
    const coordinatorWorkbookId = contextWorkbookId ?? "workbook:unknown";
    const sessionId = contextProvider.getSessionId();
    const context = makeContext({
      workbookId: coordinatorWorkbookId,
      sessionId,
      toolName: tool.name,
    });

    throwIfAborted(signal);

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

    throwIfAborted(signal);

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
