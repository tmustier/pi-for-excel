function isToolsExecuteOfficeJsPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * execute_office_js — run direct Office.js code with explicit user intent.
 *
 * Security posture:
 * - always available (not behind /experimental)
 * - each execution requires a brief explanation (shown in tool cards and approval prompts)
 * - each execution requires explicit user approval
 */

import { runMutation } from "./mutation/run-mutation.js";
import { Type, type Static } from "typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";


import { getErrorMessage } from "../utils/errors.js";
import { describeExcelFailure, parseExcelError } from "./mutation/excel-failure.js";

const MAX_CODE_CHARS = 20_000;
const MAX_EXPLANATION_CHARS = 50;
const MAX_RESULT_CHARS = 8_000;

const schema = Type.Object({
  code: Type.String({
    minLength: 1,
    maxLength: MAX_CODE_CHARS,
    description:
      "Async function body that receives context: Excel.RequestContext. "
      + "Use load() + await context.sync() before reading properties. "
      + "Return JSON-serializable data.",
  }),
  explanation: Type.String({
    minLength: 1,
    maxLength: MAX_EXPLANATION_CHARS,
    description:
      "Very brief description of what this Office.js action does (max 50 chars).",
  }),
});

type Params = Static<typeof schema>;

type ExecuteOfficeJsRunner = (context: Excel.RequestContext) => Promise<unknown>;

interface ExecuteOfficeJsToolDependencies {
  runCode: (code: string) => Promise<unknown>;
}

function normalizeExplanation(explanation: string): string {
  const trimmed = explanation.trim();
  if (trimmed.length === 0) {
    throw new Error("explanation must not be empty");
  }

  if (trimmed.length > MAX_EXPLANATION_CHARS) {
    throw new Error(`explanation must be at most ${MAX_EXPLANATION_CHARS} characters`);
  }

  return trimmed;
}

function normalizeCode(code: string): string {
  const trimmed = code.trim();
  if (trimmed.length === 0) {
    throw new Error("code must not be empty");
  }

  if (trimmed.length > MAX_CODE_CHARS) {
    throw new Error(`code exceeds ${MAX_CODE_CHARS.toLocaleString()} character limit`);
  }

  if (/\bExcel\.run\s*\(/u.test(trimmed)) {
    throw new Error("Do not call Excel.run() in code; context is provided automatically.");
  }

  return trimmed;
}

type OfficeJsRunnerCandidate = (context: Excel.RequestContext) => unknown;

function isOfficeJsRunnerCandidate(value: unknown): value is OfficeJsRunnerCandidate {
  return typeof value === "function";
}

async function loadOfficeJsRunner(code: string): Promise<ExecuteOfficeJsRunner> {
  const moduleSource = [
    "export default async function execute(context) {",
    code,
    "}",
  ].join("\n");

  const blob = new Blob([moduleSource], { type: "text/javascript" });
  const blobUrl = URL.createObjectURL(blob);

  try {
    const moduleNamespace: unknown = await import(/* @vite-ignore */ blobUrl);
    if (!isToolsExecuteOfficeJsPayloadShape(moduleNamespace)) {
      throw new Error("Compiled Office.js module did not export a valid function.");
    }

    const maybeRunner = moduleNamespace.default;
    if (!isOfficeJsRunnerCandidate(maybeRunner)) {
      throw new Error("Compiled Office.js module must export a default async function.");
    }

    return (context: Excel.RequestContext): Promise<unknown> => {
      const rawResult = maybeRunner(context);
      return Promise.resolve(rawResult);
    };
  } catch (error) {
    throw new Error(`Invalid Office.js code: ${getErrorMessage(error)}`);
  } finally {
    URL.revokeObjectURL(blobUrl);
  }
}

async function defaultRunCode(code: string): Promise<unknown> {
  const runner = await loadOfficeJsRunner(code);

  return runMutation(async (context) => {
    return runner(context);
  });
}

function jsonSafeReplacer(_key: string, value: unknown): unknown {
  if (typeof value === "bigint") {
    return value.toString();
  }

  return value;
}

function serializeResult(result: unknown): { text: string; truncated: boolean } {
  let serialized: string;

  try {
    const maybeSerialized = JSON.stringify(result, jsonSafeReplacer);
    serialized = maybeSerialized ?? "null";
  } catch (error) {
    throw new Error(`Result is not JSON-serializable: ${getErrorMessage(error)}`);
  }

  if (serialized.length <= MAX_RESULT_CHARS) {
    return { text: serialized, truncated: false };
  }

  return {
    text: `${serialized.slice(0, MAX_RESULT_CHARS)}\n…`,
    truncated: true,
  };
}

const defaultDependencies: ExecuteOfficeJsToolDependencies = {
  runCode: defaultRunCode,
};

export function createExecuteOfficeJsTool(
  dependencies: Partial<ExecuteOfficeJsToolDependencies> = {},
): AgentTool<typeof schema, undefined> {
  const resolvedDependencies: ExecuteOfficeJsToolDependencies = {
    runCode: dependencies.runCode ?? defaultDependencies.runCode,
  };

  return {
    name: "execute_office_js",
    label: "Execute Office.js",
    description:
      "Run direct Office.js (Excel JavaScript API) code with the provided Excel.RequestContext. "
      + "Use only when structured tools cannot express the operation.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const explanation = normalizeExplanation(params.explanation);
        const code = normalizeCode(params.code);
        const result = await resolvedDependencies.runCode(code);
        const serialized = serializeResult(result);

        const truncatedNote = serialized.truncated
          ? `\n\nℹ️ Result truncated to ${MAX_RESULT_CHARS.toLocaleString()} characters.`
          : "";

        const fencedResult = [
          `Executed Office.js: ${explanation}`,
          "",
          "Result:",
          "```json",
          serialized.text,
          "```",
        ].join("\n");

        return {
          content: [{
            type: "text",
            text: `${fencedResult}${truncatedNote}`,
          }],
          details: undefined,
        };
      } catch (error) {
        const parsed = parseExcelError(error);
        // Office.js codes (AccessDenied, cell-edit mode, payload limits) get the same reasons as the
        // structured tools; errors thrown by the code itself keep their message.
        const message = parsed.code ? (await describeExcelFailure(parsed)).reason : getErrorMessage(error);
        const ranEarlier = parsed.outcomeUnknown
          // Excel.run cannot be cancelled; runMutation rejects the code's later syncs, nothing more.
          ? "\n\nPi cannot stop Office.js code once it runs. It blocked this code's later Excel calls, so nothing after the " +
            "unanswered step reached Excel, but that step may still be applied, and steps before it already ran. " +
            "Work the code does outside Excel (timers, network) may still run. Read the affected ranges back before retrying."
          : parsed.code
            ? "\n\nSteps in the code before the failing call may already have changed the workbook; read back before retrying."
            : "";

        return {
          content: [{
            type: "text",
            text: `Error executing Office.js: ${message}${ranEarlier}`,
          }],
          details: undefined,
          isError: true,
        };
      }
    },
  };
}
