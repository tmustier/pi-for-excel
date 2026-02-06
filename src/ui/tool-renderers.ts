/**
 * Tool renderers for Pi-for-Excel.
 *
 * pi-web-ui's default tool renderer displays tool output in a <code-block>.
 * Our Excel tools intentionally return markdown (tables, headings, lists), so
 * render tool output using <markdown-block> for readability.
 */

import type { ImageContent, TextContent, ToolResultMessage } from "@mariozechner/pi-ai";
import {
  registerToolRenderer,
  renderHeader,
  type ToolRenderer,
  type ToolRenderResult,
} from "@mariozechner/pi-web-ui";
import { html, type TemplateResult } from "lit";
import { Code } from "lucide";

const EXCEL_TOOL_NAMES = [
  "get_workbook_overview",
  "read_range",
  "read_selection",
  "get_range_as_csv",
  "get_all_objects",
  "write_cells",
  "fill_formula",
  "search_workbook",
  "modify_structure",
  "format_cells",
  "conditional_format",
  "trace_dependencies",
  "get_recent_changes",
] as const;

function formatParamsJson(params: unknown): string {
  if (params === undefined) return "";

  // pi-ai's ToolCall.arguments are usually objects, but some providers may
  // stream/emit a JSON string. Handle both.
  try {
    if (typeof params === "string") {
      try {
        return JSON.stringify(JSON.parse(params), null, 2);
      } catch {
        return params;
      }
    }
    return JSON.stringify(params, null, 2);
  } catch {
    return String(params);
  }
}

function splitToolResultContent(result: ToolResultMessage<unknown>): {
  text: string;
  images: ImageContent[];
} {
  const text = (result.content ?? [])
    .filter((c): c is TextContent => c.type === "text")
    .map((c) => c.text)
    .join("\n");

  const images = (result.content ?? [])
    .filter((c): c is ImageContent => c.type === "image");

  return { text, images };
}

function tryFormatJsonOutput(text: string): { isJson: boolean; formatted: string } {
  const trimmed = text.trim();
  if (!trimmed) return { isJson: false, formatted: text };

  try {
    const parsed = JSON.parse(trimmed);
    return { isJson: true, formatted: JSON.stringify(parsed, null, 2) };
  } catch {
    return { isJson: false, formatted: text };
  }
}

function detectStandaloneImagePath(text: string): string | null {
  const t = text.trim();
  if (!t) return null;
  if (t.includes("\n")) return null;

  const isImage = /\.(png|jpe?g|gif|webp|svg)$/i.test(t);
  if (!isImage) return null;

  const isUnixAbs = t.startsWith("/");
  const isWinAbs = /^[A-Za-z]:\\/.test(t);
  const isFileUrl = t.startsWith("file://");

  return (isUnixAbs || isWinAbs || isFileUrl) ? t : null;
}

function pathBasename(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean);
  return parts[parts.length - 1] ?? path;
}

function toFileUrl(path: string): string {
  if (path.startsWith("file://")) return path;

  // Windows: C:\\Users\\me\\file.png → file:///C:/Users/me/file.png
  const win = /^([A-Za-z]):\\(.*)$/.exec(path);
  if (win) {
    const drive = win[1].toUpperCase();
    const rest = win[2]
      .split("\\")
      .map((seg) => encodeURIComponent(seg))
      .join("/");
    return `file:///${drive}:/${rest}`;
  }

  // Unix: /var/folders/... → file:///var/folders/...
  const encoded = path
    .split("/")
    .map((seg) => encodeURIComponent(seg))
    .join("/");
  return `file://${encoded}`;
}

function renderImages(images: ImageContent[]): TemplateResult {
  if (!images.length) return html``;

  return html`
    <div class="mt-2 grid grid-cols-1 gap-2">
      ${images.map((img) => {
        const src = `data:${img.mimeType};base64,${img.data}`;
        return html`
          <div class="border border-border rounded-lg overflow-hidden bg-background">
            <img src=${src} class="block w-full h-auto" />
          </div>
        `;
      })}
    </div>
  `;
}

function createExcelMarkdownRenderer(toolName: string): ToolRenderer<unknown, unknown> {
  return {
    render(params: unknown, result: ToolResultMessage<unknown> | undefined, isStreaming?: boolean): ToolRenderResult {
      const state = result
        ? (result.isError ? "error" : "complete")
        : isStreaming
          ? "inprogress"
          : "complete";

      const paramsJson = formatParamsJson(params);

      // With result: show input + rendered output
      if (result) {
        const { text, images } = splitToolResultContent(result);
        const standaloneImagePath = detectStandaloneImagePath(text);
        const json = tryFormatJsonOutput(text);

        const headerText = html`<span class="font-mono">${toolName}</span>`;

        return {
          content: html`
            <div class="space-y-3">
              ${renderHeader(state, Code, headerText)}

              ${paramsJson
                ? html`
                  <div>
                    <div class="text-xs font-medium mb-1 text-muted-foreground">Input</div>
                    <code-block .code=${paramsJson} language="json"></code-block>
                  </div>
                `
                : ""}

              <div>
                <div class="text-xs font-medium mb-1 text-muted-foreground">Output</div>

                ${standaloneImagePath
                  ? html`
                    <div class="text-sm">
                      <div>
                        Image file: <a
                          href=${toFileUrl(standaloneImagePath)}
                          target="_blank"
                          rel="noopener noreferrer"
                          class="underline"
                        >${pathBasename(standaloneImagePath)}</a>
                      </div>
                      <div class="mt-1 text-xs font-mono text-muted-foreground break-all">${standaloneImagePath}</div>
                    </div>
                  `
                  : json.isJson
                    ? html`<code-block .code=${json.formatted} language="json"></code-block>`
                    : html`<markdown-block .content=${text || "(no output)"}></markdown-block>`}

                ${renderImages(images)}
              </div>
            </div>
          `,
          isCustom: false,
        };
      }

      // Streaming/pending: show header + input
      if (paramsJson) {
        const headerText = html`<span class="font-mono">${toolName}</span>`;

        return {
          content: html`
            <div class="space-y-3">
              ${renderHeader(state, Code, headerText)}
              <div>
                <div class="text-xs font-medium mb-1 text-muted-foreground">Input</div>
                <code-block .code=${paramsJson} language="json"></code-block>
              </div>
            </div>
          `,
          isCustom: false,
        };
      }

      // No params or result yet
      const headerText = html`<span class="font-mono">${toolName}</span>`;
      return {
        content: html`
          <div>
            ${renderHeader(state, Code, headerText)}
          </div>
        `,
        isCustom: false,
      };
    },
  };
}

for (const name of EXCEL_TOOL_NAMES) {
  registerToolRenderer(name, createExcelMarkdownRenderer(name));
}
