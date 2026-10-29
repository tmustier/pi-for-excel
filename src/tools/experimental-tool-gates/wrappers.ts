import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";

import { assessOfficeJsCodeRisk } from "./office-js-risk.js";
import {
  wrapExtensionsManagerToolWithChangeGate,
  wrapFilesToolWithChangeGate,
} from "./local-change-gates.js";
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
  EXECUTE_WPS_JS_TOOL_NAME,
  EXTENSIONS_MANAGER_TOOL_NAME,
  FILES_TOOL_NAME,
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
import { isUnsupportedHostTool } from "../unsupported-host-tool.js";

function isExperimentalGateParamsShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function getGateParamString(record: Record<string, unknown>, key: string): string | undefined {
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
  if (!isExperimentalGateParamsShape(params)) {
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
      ...(args.bridgeUrl !== undefined ? { bridgeUrl: args.bridgeUrl } : {}),
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
      ...(args.bridgeUrl !== undefined ? { bridgeUrl: args.bridgeUrl } : {}),
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
      ...(args.bridgeUrl !== undefined ? { bridgeUrl: args.bridgeUrl } : {}),
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
      ...(args.bridgeUrl !== undefined ? { bridgeUrl: args.bridgeUrl } : {}),
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

  if (isExperimentalGateParamsShape(params)) {
    if (toolName === "python_run") {
      const code = getGateParamString(params, "code") ?? "(no code)";
      const previewLine = code.split("\n")[0] ?? code;
      return `${title}\n\nTool: python_run\nBridge: ${bridgeUrl}\nCode preview: ${previewLine}`;
    }

    if (toolName === "libreoffice_convert") {
      const inputPath = getGateParamString(params, "input_path") ?? "(unknown input)";
      const targetFormat = getGateParamString(params, "target_format") ?? "(unknown format)";
      return `${title}\n\nTool: libreoffice_convert\nBridge: ${bridgeUrl}\nInput: ${inputPath}\nTarget: ${targetFormat.toUpperCase()}`;
    }

    if (toolName === "python_transform_range") {
      const range = getGateParamString(params, "range") ?? "(unknown range)";
      const output = getGateParamString(params, "output_start_cell") ?? "(source top-left)";
      return `${title}\n\nTool: python_transform_range\nBridge: ${bridgeUrl}\nRange: ${range}\nOutput start: ${output}`;
    }
  }

  return `${title}\n\nTool: ${toolName}\nBridge: ${bridgeUrl}`;
}

function defaultRequestPythonBridgeApproval(
  _request: PythonBridgeApprovalRequest,
): Promise<boolean> {
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

  const apiName = request.apiName ?? "Office.js";
  const hostApiLabel = apiName === "WPS JSAPI" ? "the WPS JSAPI" : "the Excel API";
  const lines = [
    `Allow direct ${apiName} execution?`,
    "",
    `Action: ${explanation}`,
    `Code preview: ${firstLine}`,
  ];

  const riskIdentifiers = request.riskIdentifiers ?? [];
  if (riskIdentifiers.length > 0) {
    lines.push(
      "",
      `⚠️ Code references browser capabilities beyond ${hostApiLabel}: ${riskIdentifiers.join(", ")}.`,
      "These can reach network, storage, or credentials — review the code before allowing.",
    );
  }

  return lines.join("\n");
}

function defaultRequestOfficeJsExecuteApproval(
  _request: OfficeJsExecuteApprovalRequest,
): Promise<boolean> {
  return Promise.reject(new Error(
    "Direct workbook JS execution requires explicit user approval, but confirmation UI is unavailable.",
  ));
}

function getDirectJsExecuteApprovalRequest(
  toolCallId: string,
  params: unknown,
  apiName: "Office.js" | "WPS JSAPI",
): OfficeJsExecuteApprovalRequest {
  if (!isExperimentalGateParamsShape(params)) {
    return {
      toolCallId,
      explanation: "",
      code: "",
      apiName,
    };
  }

  const explanation = typeof params.explanation === "string"
    ? params.explanation
    : "";

  const code = typeof params.code === "string"
    ? params.code
    : "";

  return {
    toolCallId,
    explanation,
    code,
    apiName,
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
          ...(gate.bridgeUrl !== undefined ? { bridgeUrl: gate.bridgeUrl } : {}),
          params,
        });
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function wrapDirectWorkbookJsToolWithHardGate(
  tool: AgentTool,
  dependencies: ExperimentalToolGateDependencies,
  apiName: "Office.js" | "WPS JSAPI",
): AgentTool {
  const requestApproval =
    dependencies.requestOfficeJsExecuteApproval
    ?? defaultRequestOfficeJsExecuteApproval;

  return {
    ...tool,
    execute: async (toolCallId, params, signal, onUpdate) => {
      throwIfAborted(signal);

      const request = getDirectJsExecuteApprovalRequest(toolCallId, params, apiName);
      const risk = assessOfficeJsCodeRisk(request.code);

      const mode = await (dependencies.getExecutionMode?.() ?? Promise.resolve("safe" as const));
      if (mode !== "yolo" || risk.flagged || risk.destructive) {
        const approved = await requestApproval({
          ...request,
          ...(risk.flagged ? { riskIdentifiers: risk.identifiers } : {}),
          ...(risk.destructive ? { destructive: true } : {}),
          ...(signal ? { signal } : {}),
        });
        throwIfAborted(signal);
        if (!approved) {
          throw new Error(`${apiName} execution cancelled by user.`);
        }
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

function createPythonBridgeApprover(
  dependencies: ExperimentalToolGateDependencies,
): (toolName: string, bridgeUrl: string, params: unknown, signal: AbortSignal | undefined) => Promise<void> {
  const requestApproval = dependencies.requestPythonBridgeApproval ?? defaultRequestPythonBridgeApproval;
  const getApprovedBridgeUrl =
    dependencies.getApprovedPythonBridgeUrl
    ?? defaultGetApprovedPythonBridgeUrl;
  const setApprovedBridgeUrl =
    dependencies.setApprovedPythonBridgeUrl
    ?? defaultSetApprovedPythonBridgeUrl;

  return async (toolName: string, bridgeUrl: string, params: unknown, signal: AbortSignal | undefined): Promise<void> => {
    const cachedApprovalUrl = await getApprovedBridgeUrl();
    if (cachedApprovalUrl === bridgeUrl) {
      return;
    }

    const approved = await requestApproval({
      toolName,
      bridgeUrl,
      params,
      ...(signal ? { signal } : {}),
    });
    throwIfAborted(signal);
    if (!approved) {
      throw new Error("Python/LibreOffice execution cancelled by user.");
    }

    await setApprovedBridgeUrl(bridgeUrl);
  };
}

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
        await approveBridgeUsage(tool.name, bridgeUrl, params, signal);
      } else {
        const reason = gate.reason ?? "bridge_unreachable";

        if (reason === "bridge_unreachable") {
          if (tool.name === "python_transform_range") {
            return buildPythonTransformRangeGateErrorResult({
              reason,
              ...(gate.bridgeUrl !== undefined ? { bridgeUrl: gate.bridgeUrl } : {}),
            });
          }

          return buildPythonGateErrorResult({
            reason,
            ...(gate.bridgeUrl !== undefined ? { bridgeUrl: gate.bridgeUrl } : {}),
            toolName: tool.name,
          });
        }
      }

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

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
          ...(gate.bridgeUrl !== undefined ? { bridgeUrl: gate.bridgeUrl } : {}),
        });
      }

      const bridgeUrl = gate.bridgeUrl;
      if (!bridgeUrl) {
        throw new Error("Python bridge gate did not return a bridge URL.");
      }

      throwIfAborted(signal);
      await approveBridgeUsage(tool.name, bridgeUrl, params, signal);

      throwIfAborted(signal);
      return tool.execute(toolCallId, params, signal, onUpdate);
    },
  };
}

export function applyExperimentalToolGates(
  tools: AgentTool[],
  dependencies: ExperimentalToolGateDependencies = {},
): Promise<AgentTool[]> {
  const gatedTools: AgentTool[] = [];

  for (const tool of tools) {
    if (isUnsupportedHostTool(tool)) {
      gatedTools.push(tool);
      continue;
    }

    if (tool.name === TMUX_TOOL_NAME) {
      gatedTools.push(wrapTmuxToolWithHardGate(tool, dependencies));
      continue;
    }

    if (tool.name === EXECUTE_OFFICE_JS_TOOL_NAME) {
      gatedTools.push(wrapDirectWorkbookJsToolWithHardGate(tool, dependencies, "Office.js"));
      continue;
    }

    if (tool.name === EXECUTE_WPS_JS_TOOL_NAME) {
      gatedTools.push(wrapDirectWorkbookJsToolWithHardGate(tool, dependencies, "WPS JSAPI"));
      continue;
    }

    if (tool.name === FILES_TOOL_NAME) {
      gatedTools.push(wrapFilesToolWithChangeGate(tool, dependencies));
      continue;
    }

    if (tool.name === EXTENSIONS_MANAGER_TOOL_NAME) {
      gatedTools.push(wrapExtensionsManagerToolWithChangeGate(tool, dependencies));
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
