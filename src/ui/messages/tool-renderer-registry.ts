/**
 * Registry of tool renderers + default JSON fallback renderer.
 *
 * First-party replacement for pi-web-ui's tools/renderer-registry and
 * DefaultRenderer (docs/ui-ownership.md). <tool-message> calls renderTool()
 * to render a tool call; registered renderers (src/ui/tool-renderers.ts)
 * take precedence, otherwise the JSON fallback below is used.
 */

import type { ToolResultMessage } from "@earendil-works/pi-ai";
import { html, type TemplateResult } from "lit";
import { Code, Loader } from "lucide";

import { icon } from "../icons.js";
import { t } from "../../language/index.js";

export interface ToolRenderResult {
  content: TemplateResult;
  isCustom: boolean;
}

export interface ToolRenderer<TParams = DynamicValue, TDetails = DynamicValue> {
  render(
    params: TParams | undefined,
    result: ToolResultMessage<TDetails> | undefined,
    isStreaming?: boolean,
  ): ToolRenderResult;
}

const toolRenderers = new Map<string, ToolRenderer>();

export function registerToolRenderer(toolName: string, renderer: ToolRenderer): void {
  toolRenderers.set(toolName, renderer);
}

export function getToolRenderer(toolName: string): ToolRenderer | undefined {
  return toolRenderers.get(toolName);
}

/* ── Default JSON fallback renderer ───────────────────────── */

type FallbackState = "inprogress" | "complete" | "error";

function formatJson(value: DynamicValue): string {
  if (typeof value === "string") {
    try {
      return JSON.stringify(JSON.parse(value), null, 2);
    } catch {
      return value;
    }
  }

  try {
    return JSON.stringify(value, null, 2) ?? String(value);
  } catch {
    return String(value);
  }
}

function renderFallbackHeader(state: FallbackState, label: string): TemplateResult {
  return html`
    <div class="pi-tool-fallback__header" data-state=${state}>
      ${icon(Code, "sm")}
      <span>${label}</span>
      ${state === "inprogress"
        ? html`<span class="pi-tool-card__spinner" aria-hidden="true">${icon(Loader, "sm")}</span>`
        : ""}
    </div>
  `;
}

function renderFallbackSection(label: string, code: string, language: string): TemplateResult {
  return html`
    <div>
      <div class="pi-tool-fallback__label">${label}</div>
      <code-block .code=${code} language=${language}></code-block>
    </div>
  `;
}

function renderDefaultTool(
  params: DynamicValue,
  result: ToolResultMessage<DynamicValue> | undefined,
  isStreaming: boolean,
): ToolRenderResult {
  const state: FallbackState = result
    ? (result.isError ? "error" : "complete")
    : isStreaming
      ? "inprogress"
      : "complete";

  const paramsJson = params === undefined || params === null ? "" : formatJson(params);

  if (result) {
    let output =
      result.content
        ?.filter((c) => c.type === "text")
        .map((c) => c.text)
        .join("\n") || t("messages.noOutput");
    let outputLanguage = "text";
    try {
      output = JSON.stringify(JSON.parse(output), null, 2);
      outputLanguage = "json";
    } catch {
      // Not JSON — keep as plain text.
    }

    return {
      content: html`
        <div class="pi-tool-fallback">
          ${renderFallbackHeader(state, t("messages.toolCall"))}
          ${paramsJson ? renderFallbackSection(t("messages.toolInput"), paramsJson, "json") : ""}
          ${renderFallbackSection(t("messages.toolOutput"), output, outputLanguage)}
        </div>
      `,
      isCustom: false,
    };
  }

  if (isStreaming && (!paramsJson || paramsJson === "{}" || paramsJson === "null")) {
    return {
      content: html`
        <div class="pi-tool-fallback">
          ${renderFallbackHeader(state, t("messages.preparingParams"))}
        </div>
      `,
      isCustom: false,
    };
  }

  return {
    content: html`
      <div class="pi-tool-fallback">
        ${renderFallbackHeader(state, t("messages.toolCall"))}
        ${paramsJson ? renderFallbackSection(t("messages.toolInput"), paramsJson, "json") : ""}
      </div>
    `,
    isCustom: false,
  };
}

/**
 * Render a tool call: registered renderer if available, JSON fallback otherwise.
 */
export function renderTool(
  toolName: string,
  params: DynamicValue,
  result: ToolResultMessage<DynamicValue> | undefined,
  isStreaming: boolean,
): ToolRenderResult {
  const renderer = getToolRenderer(toolName);
  if (renderer) {
    return renderer.render(params, result, isStreaming);
  }
  return renderDefaultTool(params, result, isStreaming);
}
