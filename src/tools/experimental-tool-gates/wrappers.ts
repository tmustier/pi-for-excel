import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  buildPythonBridgeGateErrorMessage,
  buildTmuxBridgeGateErrorMessage,
  defaultGetApprovedPythonBridgeUrl,
  defaultSetApprovedPythonBridgeUrl,
  evaluatePythonBridgeGate,
  evaluateTmuxBridgeGate,
} from "./evaluation.js";
import {
  EXECUTE_OFFICE_JS_TOOL_NAME,
  PYTHON_BRIDGE_ONLY_TOOL_NAMES,
  PYTHON_FALLBACK_TOOL_NAMES,
  TMUX_TOOL_NAME,
  type ExperimentalToolGateDependencies,
  type OfficeJsExecuteApprovalRequest,
  type PythonBridgeApprovalRequest,
  type PythonBridgeGateReason,
  type TmuxBridgeGateReason,
} from "./types.js";
import type {
  LibreOfficeBridgeDetails,
  PythonBridgeDetails,
  PythonTransformRangeDetails,
  TmuxBridgeDetails,
} from "../tool-details.js";

function isRecordObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getRecordValue(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === "string" && value.trim().length > 0 ? value.trim() : undefined;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (!signal?.aborted) {
    return;
  }

  throw new Error("Aborted");
}

function getTmuxActionFromParams(params: unknown): string {
  if (!isRecordObject(params)) {
    return "list_sessions";
  }

  const action = params.action;
  return typeof action === "string" && action.trim().length > 0
    ? action.trim()
    : "list_sessions";
}

function getPythonActionForTool(toolName: string): string {
  return toolName === "python_transform_range"
    ? "transform_range"
    : "run_python";
}

function buildGateErrorText(message: string, skillName: string): string {
  return `${message}\nSkill: ${skillName}`;
}

function buildTmuxGateErrorResult(args: {
  reason: TmuxBridgeGateReason;
  bridgeUrl?: string;
  params: unknown;
}): AgentToolResult<TmuxBridgeDetails> {
  const message = buildTmuxBridgeGateErrorMessage(args.reason);

  return {
    content: [{
      type: "text",
      text: buildGateErrorText(message, "tmux-bridge"),
    }],
    details: {
      kind: "tmux_bridge",
      ok: false,
      action: getTmuxActionFromParams(args.params),
      bridgeUrl: args.bridgeUrl,
      error: message,
      gateReason: args.reason,
      skillHint: "tmux-bridge",
    },
  };
}

function buildPythonGateErrorResult(args: {
  reason: PythonBridgeGateReason;
  bridgeUrl?: string;
  toolName: string;
}): AgentToolResult<PythonBridgeDetails> {
  const message = buildPythonBridgeGateErrorMessage(args.reason);

  return {
    content: [{
      type: "text",
      text: buildGateErrorText(message, "python-bridge"),
    }],
    details: {
      kind: "python_bridge",
      ok: false,
      action: getPythonActionForTool(args.toolName),
      bridgeUrl: args.bridgeUrl,
      error: message,
      gateReason: args.reason,
      skillHint: "python-bridge",
    },
  };
}

function buildPythonTransformRangeGateErrorResult(args: {
  reason: PythonBridgeGateReason;
  bridgeUrl?: string;
}): AgentToolResult<PythonTransformRangeDetails> {
  const message = buildPythonBridgeGateErrorMessage(args.reason);

  return {
    content: [{
      type: "text",
      text: buildGateErrorText(message, "python-bridge"),
    }],
    details: {
      kind: "python_transform_range",
      blocked: false,
      bridgeUrl: args.bridgeUrl,
      error: message,
      gateReason: args.reason,
      skillHint: "python-bridge",
    },
  };
}

function buildLibreOfficeGateErrorResult(args: {
  reason: PythonBridgeGateReason;
  bridgeUrl?: string;
}): AgentToolResult<LibreOfficeBridgeDetails> {
  const message = buildPythonBridgeGateErrorMessage(args.reason);

  return {
    content: [{
      type: "text",
      text: buildGateErrorText(message, "python-bridge"),
    }],
    details: {
      kind: "libreoffice_bridge",
      ok: false,
      action: "convert",
      bridgeUrl: args.bridgeUrl,
      error: message,
      gateReason: args.reason,
      skillHint: "python-bridge",
    },
  };
}

export function buildPythonBridgeApprovalMessage(
  toolName: string,
  bridgeUrl: string,
  params: unknown,
): string {
  const title = "Allow local Python / LibreOffice execution?";

  if (isRecordObject(params)) {
    if (toolName === "python_run") {
      const code = getRecordValue(params, "code") ?? "(no code)";
      const previewLine = code.split("\n")[0] ?? code;
      return `${title}\n\nTool: python_run\nBridge: ${bridgeUrl}\nCode preview: ${previewLine}`;
    }

    if (toolName === "libreoffice_convert") {
      const inputPath = getRecordValue(params, "input_path") ?? "(unknown input)";
      const targetFormat = getRecordValue(params, "target_format") ?? "(unknown format)";
      return `${title}\n\nTool: libreoffice_convert\nBridge: ${bridgeUrl}\nInput: ${inputPath}\nTarget: ${targetFormat.toUpperCase()}`;
    }

    if (toolName === "python_transform_range") {
      const range = getRecordValue(params, "range") ?? "(unknown range)";
      const output = getRecordValue(params, "output_start_cell") ?? "(source top-left)";
      return `${title}\n\nTool: python_transform_range\nBridge: ${bridgeUrl}\nRange: ${range}\nOutput start: ${output}`;
    }
  }

  return `${title}\n\nTool: ${toolName}\nBridge: ${bridgeUrl}`;
}

function defaultRequestPythonBridgeApproval(
  _request: PythonBridgeApprovalRequest,
): Promise<boolean> {
  // Fails open when no UI approval handler is injected.
  return Promise.resolve(true);
}

export function buildOfficeJsExecuteApprovalMessage(
  request: OfficeJsExecuteApprovalRequest,
): string {
  const explanation = request.explanation.trim().length > 0
    ? request.explanation.trim()
    : "(no explanation provided)";

  const firstLine = request.code
    .split("\n")
    .map((line) => line.trim())
    .find((line) => line.length > 0)
    ?? "(no code preview)";

  return [
    "Allow direct Office.js execution?",
    "",
    `Action: ${explanation}`,
    `Code preview: ${firstLine}`,
  ].join("\n");
}

function defaultRequestOfficeJsExecuteApproval(
  _request: OfficeJsExecuteApprovalRequest,
): Promise<boolean> {
  return Promise.reject(new Error(
    "Office.js execution requires explicit user approval, but confirmation UI is unavailable.",
  ));
}

function getOfficeJsExecuteApprovalRequest(params: unknown): OfficeJsExecuteApprovalRequest {
  if (!isRecordObject(params)) {
    return {
      explanation: "",
      code: "",
    };
  }

  const explanation = typeof params.explanation === "string"
    ? params.explanation
    : "";

  const code = typeof params.code === "string"
    ? params.code
    : "";

  return {
    explanation,
    code,
  };
}

function wrapTmuxToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      const gate = await evaluateTmuxBridgeGate(dependencies);
      if (!gate.allowed) {
        const reason = gate.reason ?? "bridge_unreachable";
        return buildTmuxGateErrorResult({
          reason,
          bridgeUrl: gate.bridgeUrl,
          params,
        });
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function wrapExecuteOfficeJsToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  const requestApproval =
    dependencies.requestOfficeJsExecuteApproval
    ?? defaultRequestOfficeJsExecuteApproval;

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      // Auto mode trusts Office.js — skip approval prompt.
      const mode = await (dependencies.getExecutionMode?.() ?? Promise.resolve("safe" as const));
      if (mode !== "yolo") {
        const approved = await requestApproval(getOfficeJsExecuteApprovalRequest(params));
        if (!approved) {
          throw new Error("Office.js execution cancelled by user.");
        }
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function createPythonBridgeApprover(
  dependencies: ExperimentalToolGateDependencies,
): (toolName: string, bridgeUrl: string, params: unknown) => Promise<void> {
  const requestApproval = dependencies.requestPythonBridgeApproval ?? defaultRequestPythonBridgeApproval;
  const getApprovedBridgeUrl =
    dependencies.getApprovedPythonBridgeUrl
    ?? defaultGetApprovedPythonBridgeUrl;
  const setApprovedBridgeUrl =
    dependencies.setApprovedPythonBridgeUrl
    ?? defaultSetApprovedPythonBridgeUrl;

  return async (toolName: string, bridgeUrl: string, params: unknown): Promise<void> => {
    const cachedApprovalUrl = await getApprovedBridgeUrl();
    if (cachedApprovalUrl === bridgeUrl) {
      return;
    }

    const approved = await requestApproval({
      toolName,
      bridgeUrl,
      params,
    });
    if (!approved) {
      throw new Error("Python/LibreOffice execution cancelled by user.");
    }

    await setApprovedBridgeUrl(bridgeUrl);
  };
}

/**
 * Python tools with a built-in fallback (`python_run`, `python_transform_range`).
 *
 * Behavior:
 * - If bridge is configured + reachable, request approval once per bridge URL.
 * - If bridge URL is missing/invalid, allow tool execution so it can fall back
 *   to in-browser Pyodide.
 * - If bridge is configured but unreachable, keep blocking with an explicit
 *   reachability error.
 */
function wrapPythonToolWithOptionalBridgeApproval(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  const approveBridgeUsage = createPythonBridgeApprover(dependencies);

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      const gate = await evaluatePythonBridgeGate(dependencies);
      if (gate.allowed) {
        const bridgeUrl = gate.bridgeUrl;
        if (!bridgeUrl) {
          throw new Error("Python bridge gate did not return a bridge URL.");
        }

        throwIfAborted(signal);
        await approveBridgeUsage(tool.name, bridgeUrl, params);
      } else {
        const reason = gate.reason ?? "bridge_unreachable";

        // Missing/invalid URL means the tool can still run via Pyodide fallback.
        if (reason === "bridge_unreachable") {
          if (tool.name === "python_transform_range") {
            return buildPythonTransformRangeGateErrorResult({
              reason,
              bridgeUrl: gate.bridgeUrl,
            });
          }

          return buildPythonGateErrorResult({
            reason,
            bridgeUrl: gate.bridgeUrl,
            toolName: tool.name,
          });
        }
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

/**
 * Python tools that strictly require the native bridge (`libreoffice_convert`).
 */
function wrapPythonBridgeOnlyToolWithApprovalGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
): AgentTool {
  const approveBridgeUsage = createPythonBridgeApprover(dependencies);

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      const gate = await evaluatePythonBridgeGate(dependencies);
      if (!gate.allowed) {
        const reason = gate.reason ?? "bridge_unreachable";
        return buildLibreOfficeGateErrorResult({
          reason,
          bridgeUrl: gate.bridgeUrl,
        });
      }

      const bridgeUrl = gate.bridgeUrl;
      if (!bridgeUrl) {
        throw new Error("Python bridge gate did not return a bridge URL.");
      }

      throwIfAborted(signal);
      await approveBridgeUsage(tool.name, bridgeUrl, params);

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

/**
 * Apply execution gates to tool calls.
 *
 * Current rules:
 * - `tmux` requires a healthy bridge URL (custom override or default localhost URL).
 * - `python_run` and `python_transform_range` can run with Pyodide fallback
 *   when bridge URL is missing/invalid; bridge approval is required when a
 *   reachable bridge is present.
 * - `libreoffice_convert` strictly requires a configured + reachable bridge and
 *   approval (no Pyodide fallback).
 * - `execute_office_js` requires explicit user confirmation on every call.
 * - `files` has no gate — read, write, and delete are always available.
 */
export function applyExperimentalToolGates(
  tools: AgentTool[],
  dependencies: ExperimentalToolGateDependencies = {},
): Promise<AgentTool[]> {
  const gatedTools: AgentTool[] = [];

  for (const tool of tools) {
    if (tool.name === TMUX_TOOL_NAME) {
      gatedTools.push(wrapTmuxToolWithHardGate(tool, dependencies));
      continue;
    }

    if (tool.name === EXECUTE_OFFICE_JS_TOOL_NAME) {
      gatedTools.push(wrapExecuteOfficeJsToolWithHardGate(tool, dependencies));
      continue;
    }

    if (PYTHON_FALLBACK_TOOL_NAMES.has(tool.name)) {
      gatedTools.push(wrapPythonToolWithOptionalBridgeApproval(tool, dependencies));
      continue;
    }

    if (PYTHON_BRIDGE_ONLY_TOOL_NAMES.has(tool.name)) {
      gatedTools.push(wrapPythonBridgeOnlyToolWithApprovalGate(tool, dependencies));
      continue;
    }

    gatedTools.push(tool);
  }

  return Promise.resolve(gatedTools);
}
