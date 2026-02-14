/**
 * Tool wrapper that routes mutating tool calls through the workbook coordinator.
 */

import type { AgentTool } from "@mariozechner/pi-agent-core";
import type { TSchema } from "@sinclair/typebox";

import {
  formatExecutionModeLabel,
  type ExecutionMode,
} from "../execution/mode.js";
import type { WorkbookCoordinator, WorkbookOperationContext } from "../workbook/coordinator.js";
import { getErrorMessage } from "../utils/errors.js";
import { getToolContextImpact, getToolExecutionMode, type ToolContextImpact } from "./execution-policy.js";

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

function isRecordObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getActionParam(params: unknown): string | null {
  if (!isRecordObject(params)) return null;

  const action = params.action;
  return typeof action === "string" && action.trim().length > 0 ? action.trim() : null;
}

function getRangeParam(params: unknown): string | null {
  if (!isRecordObject(params)) return null;

  const range = params.range;
  return typeof range === "string" && range.trim().length > 0 ? range.trim() : null;
}

function buildMutationApprovalMessage(request: MutationApprovalRequest): string {
  const modeLabel = formatExecutionModeLabel(request.executionMode);
  const action = getActionParam(request.params);
  const range = getRangeParam(request.params);

  const toolLabel = action ? `${request.toolName}:${action}` : request.toolName;
  const lines = [
    `Allow workbook mutation in ${modeLabel} mode?`,
    "",
    `Tool: ${toolLabel}`,
  ];

  if (range) {
    lines.push(`Range: ${range}`);
  }

  lines.push("", "Tip: run /yolo on to disable pre-execution mutation confirmations.");
  return lines.join("\n");
}

function defaultGetExecutionMode(): Promise<ExecutionMode> {
  return Promise.resolve("yolo");
}

function defaultRequestMutationApproval(request: MutationApprovalRequest): Promise<boolean> {
  if (typeof window === "undefined" || typeof window.confirm !== "function") {
    return Promise.reject(new Error(
      "Safe mode requires explicit user approval, but confirmation UI is unavailable.",
    ));
  }

  return Promise.resolve(window.confirm(buildMutationApprovalMessage(request)));
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
    throw new Error("Mutation cancelled by user (Safe mode).");
  }
}

function wrapTool<TParameters extends TSchema, TDetails>(
  tool: AgentTool<TParameters, TDetails>,
  coordinator: WorkbookCoordinator,
  contextProvider: WorkbookCoordinatorContextProvider,
  mutationObserver: WorkbookMutationObserver | undefined,
  executionPolicy: WorkbookExecutionPolicy,
): AgentTool<TParameters, TDetails> {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
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
        } catch (error: unknown) {
          console.warn("[pi] Workbook mutation observer failed:", getErrorMessage(error));
        }
      }

      return out.result;
    },
  };
}

export function withWorkbookCoordinator(
  tools: AgentTool[],
  coordinator: WorkbookCoordinator,
  contextProvider: WorkbookCoordinatorContextProvider,
  mutationObserver?: WorkbookMutationObserver,
  executionPolicy: WorkbookExecutionPolicy = {},
): AgentTool[] {
  return tools.map((tool) => wrapTool(tool, coordinator, contextProvider, mutationObserver, executionPolicy));
}
