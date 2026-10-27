/**
 * <user-message>, <assistant-message>, <tool-message> — chat message elements.
 *
 * First-party replacements for pi-web-ui's Messages.js components
 * (docs/ui-ownership.md). Tag names, light DOM, and the CSS-facing structure
 * (`.user-message-container`, `.pi-assistant-body`, `.pi-assistant-aborted`,
 * `.pi-tool-card-fallback`) are preserved; Tailwind utility classes are
 * replaced with semantic `pi-*` classes styled in src/ui/theme/content/.
 *
 * Intentional divergences from upstream:
 * - No usage/cost row (theme CSS always hid it; the status bar shows context).
 * - Non-custom tool renders get the `.pi-tool-card-fallback` wrapper class
 *   directly instead of via applyMessageStyleHooks() DOM post-processing.
 */

import type { AgentTool } from "@earendil-works/pi-agent-core";
import type {
  AssistantMessage as AssistantMessageType,
  ToolCall,
  ToolResultMessage,
  UserMessage as UserMessageType,
} from "@earendil-works/pi-ai";
import { html, LitElement, nothing, type TemplateResult } from "lit";
import { customElement, property } from "lit/decorators.js";

import { t } from "../../language/index.js";
import type { UserMessageWithAttachments } from "../../messages/attachments.js";
import { getToolRenderer, renderTool } from "./tool-renderer-registry.js";
import {
  getThinkingDurationMs,
  isDisclosureExpanded,
  setDisclosureExpanded,
  subscribeDisclosure,
  thinkingKey,
} from "./transcript-state.js";
import "./attachment-tile.js";
import "./markdown-block.js";
import "./thinking-block.js";

@customElement("user-message")
export class UserMessage extends LitElement {
  @property({ type: Object }) message?: UserMessageType | UserMessageWithAttachments;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
  }

  override render() {
    const message = this.message;
    if (!message) return html``;

    const content =
      typeof message.content === "string"
        ? message.content
        : message.content.find((chunk) => chunk.type === "text")?.text || "";

    const attachments =
      message.role === "user-with-attachments" && message.attachments?.length
        ? message.attachments
        : null;

    return html`
      <div class="pi-user-row">
        <div class="user-message-container">
          <markdown-block .content=${content}></markdown-block>
          ${attachments
            ? html`
                <div class="pi-user-attachments">
                  ${attachments.map(
                    (attachment) => html`<attachment-tile .attachment=${attachment}></attachment-tile>`,
                  )}
                </div>
              `
            : nothing}
        </div>
      </div>
    `;
  }
}

interface RenderItem {
  template: TemplateResult;
  group?: ToolCall;
}

const TOOL_GROUP_LABEL_KEYS: Record<string, string> = {
  fill_formula: "tool-group.fill_formula",
  write_cells: "tool-group.write_cells",
  read_range: "tool-group.read_range",
  format_cells: "tool-group.format_cells",
  conditional_format: "tool-group.conditional_format",
  view_settings: "tool-group.view_settings",
  charts: "tool-group.charts",
  execute_office_js: "tool-group.script",
  execute_wps_js: "tool-group.script",
};

@customElement("assistant-message")
export class AssistantMessage extends LitElement {
  @property({ type: Object }) message?: AssistantMessageType;
  @property({ type: Array }) tools?: AgentTool[];
  @property({ type: Object }) pendingToolCalls?: ReadonlySet<string>;
  @property({ type: Boolean }) hideToolCalls = false;
  @property({ type: Object }) toolResultsById?: Map<string, ToolResultMessage>;
  @property({ type: Boolean }) isStreaming = false;
  /**
   * The turn is still running, so a tool call without a result is about to
   * run or is running even before pi-agent-core marks it pending.
   */
  @property({ type: Boolean }) awaitingToolResults = false;

  private _unsubscribe: (() => void) | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
    // Group open/closed state lives in transcript state.
    this._unsubscribe = subscribeDisclosure(() => this.requestUpdate());
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
  }

  override render() {
    const message = this.message;
    if (!message) return html``;

    // Render content in the order it appears. Only the last block of a
    // streaming response can still be growing, so a thinking block reads
    // "Thinking…" only until the model moves on to text or a tool call.
    const lastIndex = message.content.length - 1;
    const items: RenderItem[] = [];
    for (const [index, chunk] of message.content.entries()) {
      if (chunk.type === "text" && chunk.text.trim() !== "") {
        items.push({ template: html`<markdown-block .content=${chunk.text}></markdown-block>` });
      } else if (chunk.type === "thinking" && chunk.thinking.trim() !== "") {
        const key = thinkingKey(message, index);
        items.push({ template: html`
          <thinking-block
            .content=${chunk.thinking}
            .isStreaming=${this.isStreaming && index === lastIndex}
            .disclosureKey=${key}
            .durationMs=${getThinkingDurationMs(key)}
          ></thinking-block>
        ` });
      } else if (chunk.type === "toolCall") {
        if (this.hideToolCalls) continue;

        const tool = this.tools?.find((candidate) => candidate.name === chunk.name);
        const result = this.toolResultsById?.get(chunk.id);
        const pending = !result && (
          (this.pendingToolCalls?.has(chunk.id) ?? false)
          || (this.awaitingToolResults && message.stopReason === "toolUse")
        );

        // Aborted turn with no result for this call → render as aborted.
        const aborted = message.stopReason === "aborted" && !result;
        const finished = result ? !result.isError : !aborted && !pending && !this.isStreaming;

        items.push({
          // Only finished calls with a tool card join a group; a running or
          // failed call (or a fallback JSON render) breaks the run.
          ...(finished && getToolRenderer(chunk.name) ? { group: chunk } : {}),
          template: html`
            <tool-message
              .tool=${tool}
              .toolCall=${chunk}
              .result=${result}
              .pending=${pending}
              .aborted=${aborted}
              .isStreaming=${this.isStreaming}
              .disclosureKey=${`tool:${message.timestamp}:${chunk.id}`}
            ></tool-message>
          `,
        });
      }
    }

    // Keep groups in Lit's render tree: moving its nodes after rendering
    // breaks part ownership when switching tabs or finishing a tool.
    const orderedParts: TemplateResult[] = [];
    let end = 0;
    for (const [index, item] of items.entries()) {
      if (index < end) continue;
      const leader = item.group;
      end = index + 1;
      if (leader) {
        while (items[end]?.group?.name === leader.name) end++;
      }
      const count = end - index;
      if (!leader || count < 2) {
        orderedParts.push(item.template);
        continue;
      }

      const key = `group:${message.timestamp}:${leader.id}`;
      const expanded = isDisclosureExpanded(key, count < 3);
      const labelKey = TOOL_GROUP_LABEL_KEYS[leader.name];
      orderedParts.push(html`
        <div class=${expanded ? "pi-tool-group" : "pi-tool-group pi-tool-group--collapsed"}>
          <button
            type="button"
            class="pi-tool-group__header"
            aria-expanded=${expanded ? "true" : "false"}
            @click=${() => setDisclosureExpanded(key, !expanded)}
          >
            <span class="pi-tool-group__chevron">▸</span>
            <span class="pi-tool-group__label">${labelKey
              ? t(labelKey, { count })
              : t("tool-group.other", { count, tool: leader.name })}</span>
          </button>
          ${items.slice(index, end).map((member) => member.template)}
        </div>
      `);
    }

    return html`
      <div>
        ${orderedParts.length ? html`<div class="pi-assistant-body">${orderedParts}</div>` : nothing}
        ${message.stopReason === "error" && message.errorMessage
          ? html`
              <div class="pi-assistant-error">
                <strong>${t("messages.errorLabel")}</strong> ${message.errorMessage}
              </div>
            `
          : nothing}
        ${message.stopReason === "aborted"
          ? html`<span class="pi-assistant-aborted">${t("messages.aborted")}</span>`
          : nothing}
      </div>
    `;
  }
}

@customElement("tool-message")
export class ToolMessage extends LitElement {
  @property({ type: Object }) toolCall?: ToolCall;
  @property({ type: Object }) tool?: AgentTool;
  @property({ type: Object }) result?: ToolResultMessage;
  @property({ type: Boolean }) pending = false;
  @property({ type: Boolean }) aborted = false;
  @property({ type: Boolean }) isStreaming = false;
  @property() disclosureKey = "";

  private _unsubscribe: (() => void) | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  // No inline display style: theme CSS sets `display: block`, so a collapsed
  // tool group can still hide its members.
  override connectedCallback(): void {
    super.connectedCallback();
    this._unsubscribe = subscribeDisclosure(() => this.requestUpdate());
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
  }

  override render() {
    const toolCall = this.toolCall;
    if (!toolCall) return html``;

    const toolName = this.tool?.name || toolCall.name;

    // Aborted calls render like an errored result with no content.
    const result: ToolResultMessage | undefined = this.aborted
      ? {
          role: "toolResult",
          isError: true,
          content: [],
          toolCallId: toolCall.id,
          toolName: toolCall.name,
          timestamp: Date.now(),
        }
      : this.result;

    const key = this.disclosureKey;
    const expanded = isDisclosureExpanded(key);
    const renderResult = renderTool(
      toolName,
      toolCall.arguments,
      result,
      !this.aborted && (this.isStreaming || this.pending),
      { expanded, toggle: () => setDisclosureExpanded(key, !expanded) },
    );

    // Custom renderers own their full card; fallback renders get a wrapper.
    if (renderResult.isCustom) {
      return renderResult.content;
    }

    return html`<div class="pi-tool-card-fallback">${renderResult.content}</div>`;
  }
}

declare global {
  interface HTMLElementTagNameMap {
    "user-message": UserMessage;
    "assistant-message": AssistantMessage;
    "tool-message": ToolMessage;
  }
}
