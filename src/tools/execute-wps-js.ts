function isToolsExecuteWpsJsPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * execute_wps_js — run direct WPS ET JSAPI code with explicit user intent.
 *
 * WPS exposes a synchronous VBA-like object model. The user-provided function
 * body receives `Application` (the active WPS ET Application) and `wps` in scope.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  getWpsEtApplication,
  getWpsGlobalForTaskPane,
  type WpsEtApplication,
  type WpsGlobal,
} from "../host/wps/jsapi.js";
import { getErrorMessage } from "../utils/errors.js";

const MAX_CODE_CHARS = 20_000;
const MAX_EXPLANATION_CHARS = 50;
const MAX_RESULT_CHARS = 8_000;

const schema = Type.Object({
  code: Type.String({
    minLength: 1,
    maxLength: MAX_CODE_CHARS,
    description:
      "Synchronous function body with Application (WPS ET Application) and wps in scope. "
      + "Use the VBA-like WPS JSAPI directly and return JSON-serializable data.",
  }),
  explanation: Type.String({
    minLength: 1,
    maxLength: MAX_EXPLANATION_CHARS,
    description:
      "Very brief description of what this WPS JSAPI action does (max 50 chars).",
  }),
});

type Params = Static<typeof schema>;

type ExecuteWpsJsRunner = (Application: WpsEtApplication, wps: WpsGlobal | null) => DynamicValue;

interface ExecuteWpsJsToolDependencies {
  runCode: (code: string) => Promise<DynamicValue>;
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
    throw new Error("Do not call Excel.run() in WPS JSAPI code; use the provided Application object.");
  }

  return trimmed;
}

type WpsJsRunnerCandidate = (Application: WpsEtApplication, wps: WpsGlobal | null) => DynamicValue;

function isWpsJsRunnerCandidate(value: DynamicValue): value is WpsJsRunnerCandidate {
  return typeof value === "function";
}

async function loadWpsJsRunner(code: string): Promise<ExecuteWpsJsRunner> {
  const moduleSource = [
    "export default function execute(Application, wps) {",
    code,
    "}",
  ].join("\n");

  const blob = new Blob([moduleSource], { type: "text/javascript" });
  const blobUrl = URL.createObjectURL(blob);

  try {
    const moduleNamespace: DynamicValue = await import(/* @vite-ignore */ blobUrl);
    if (!isToolsExecuteWpsJsPayloadShape(moduleNamespace)) {
      throw new Error("Compiled WPS JSAPI module did not export a valid function.");
    }

    const maybeRunner = moduleNamespace.default;
    if (!isWpsJsRunnerCandidate(maybeRunner)) {
      throw new Error("Compiled WPS JSAPI module must export a default function.");
    }

    return (Application: WpsEtApplication, wps: WpsGlobal | null): DynamicValue => maybeRunner(Application, wps);
  } catch (error) {
    throw new Error(`Invalid WPS JSAPI code: ${getErrorMessage(error)}`);
  } finally {
    URL.revokeObjectURL(blobUrl);
  }
}

async function defaultRunCode(code: string): Promise<DynamicValue> {
  const app = getWpsEtApplication();
  if (!app) {
    throw new Error("WPS ET Application is unavailable.");
  }

  const runner = await loadWpsJsRunner(code);
  return runner(app, getWpsGlobalForTaskPane());
}

function jsonSafeReplacer(_key: string, value: DynamicValue): DynamicValue {
  if (typeof value === "bigint") {
    return value.toString();
  }

  return value;
}

function serializeResult(result: DynamicValue): { text: string; truncated: boolean } {
  let serialized: string;

  try {
    const maybeSerialized = JSON.stringify(result, jsonSafeReplacer, 2);
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

const defaultDependencies: ExecuteWpsJsToolDependencies = {
  runCode: defaultRunCode,
};

export function createExecuteWpsJsTool(
  dependencies: Partial<ExecuteWpsJsToolDependencies> = {},
): AgentTool<typeof schema, undefined> {
  const resolvedDependencies: ExecuteWpsJsToolDependencies = {
    runCode: dependencies.runCode ?? defaultDependencies.runCode,
  };

  return {
    name: "execute_wps_js",
    label: "Execute WPS JSAPI",
    description:
      "Run direct WPS Spreadsheets JSAPI code with Application in scope. "
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
          `Executed WPS JSAPI: ${explanation}`,
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
        const message = getErrorMessage(error);

        return {
          content: [{
            type: "text",
            text: `Error executing WPS JSAPI: ${message}`,
          }],
          details: undefined,
        };
      }
    },
  };
}
