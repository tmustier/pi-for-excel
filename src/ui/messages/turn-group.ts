/**
 * <turn-group> — the record of what the agent did in one turn.
 *
 * One card per turn (docs/proposals/sidebar-redesign.md §5). The header
 * carries the turn's state: "Working" with the elapsed time while it runs,
 * "Waiting for approval" while a row asks, "Worked for 8s" once done, with
 * the count of failed steps. Each tool call is a row: a past-tense sentence
 * with the reference in mono, one meta fact, and a chevron that drills into
 * the call's input and result. A failed row turns red and shows the error
 * the model received; there is no Retry, the model's recovery is the next
 * row. A row waiting for approval asks in place. The footer is the
 * workbook-change receipt with Undo and Review.
 *
 * The group collapses itself when the answer starts; the footer stays. Open
 * state follows transcript-state (details mode, hand toggles).
 */

import { html, LitElement, nothing, type TemplateResult } from "lit";
import { customElement, property } from "lit/decorators.js";
import { AlertTriangle, Check, ChevronDown, ChevronUp, Circle, Loader, type IconNode } from "lucide";

import { t } from "../../language/index.js";
import { getInlineApprovals, type InlineApproval } from "../approval-dialog.js";
import { icon } from "../icons.js";
import { hasToolSetupCard, renderToolSetupCards } from "../tool-renderers.js";
import { describeToolRow, renderToolRowDetail, toolKind, toolKindIcon, type ToolKind, type ToolRowTense } from "../tool-rows.js";
import { renderTurnChangesReceipt } from "./turn-changes-receipt.js";
import { getThinkingDurationMs, isDisclosureExpanded, setDisclosureExpanded, subscribeDisclosure, thinkingKey } from "./transcript-state.js";
import type { TurnGroupModel, TurnNoteStep, TurnToolStep } from "./turn-model.js";
import "./markdown-block.js";
import "./thinking-block.js";

const KIND_ORDER: readonly ToolKind[] = ["sheet", "python", "web", "code", "extension"];

export function formatElapsed(ms: number): string {
  const seconds = Math.max(0, Math.round(ms / 1000));
  return seconds < 60 ? `${seconds}s` : `${Math.floor(seconds / 60)}m ${seconds % 60}s`;
}

const STATUS_ICONS: Record<TurnToolStep["state"] | "waiting", { icon: IconNode; className: string }> = {
  live: { icon: Loader, className: "pi-turn-row__status pi-turn-row__status--live" },
  done: { icon: Check, className: "pi-turn-row__status pi-turn-row__status--done" },
  failed: { icon: AlertTriangle, className: "pi-turn-row__status pi-turn-row__status--failed" },
  stopped: { icon: Circle, className: "pi-turn-row__status pi-turn-row__status--stopped" },
  waiting: { icon: Circle, className: "pi-turn-row__status pi-turn-row__status--waiting" },
};

@customElement("turn-group")
export class TurnGroup extends LitElement {
  @property({ type: Object }) model?: TurnGroupModel;

  private _unsubscribe: (() => void) | undefined;
  private _ticker: number | undefined;

  protected override createRenderRoot(): HTMLElement | DocumentFragment {
    return this;
  }

  override connectedCallback(): void {
    super.connectedCallback();
    this.style.display = "block";
    this._unsubscribe = subscribeDisclosure(() => this.requestUpdate());
  }

  override disconnectedCallback(): void {
    super.disconnectedCallback();
    this._unsubscribe?.();
    this._unsubscribe = undefined;
    this._stopTicking();
  }

  override updated(): void {
    // The elapsed time ticks once a second while the turn runs.
    if (this.model?.live) this._startTicking();
    else this._stopTicking();
  }

  private _startTicking(): void {
    if (this._ticker !== undefined) return;
    this._ticker = window.setInterval(() => this.requestUpdate(), 1000);
  }

  private _stopTicking(): void {
    if (this._ticker === undefined) return;
    window.clearInterval(this._ticker);
    this._ticker = undefined;
  }

  private _approvalFor(toolCallId: string): InlineApproval | undefined {
    return getInlineApprovals().find((approval) => approval.toolCallId === toolCallId);
  }

  override render() {
    const model = this.model;
    if (!model) return nothing;

    const toolSteps = model.steps.filter((step): step is TurnToolStep => step.type === "tool");
    const failed = toolSteps.filter((step) => step.state === "failed").length;
    const waiting = toolSteps.some((step) => this._approvalFor(step.toolCall.id) !== undefined);
    const kinds = KIND_ORDER.filter((kind) => toolSteps.some((step) => toolKind(step.toolCall.name) === kind));
    const state = waiting ? "waiting" : model.live ? "live" : "done";
    const elapsed = formatElapsed((model.endedAt ?? Date.now()) - model.startedAt);
    const expanded = isDisclosureExpanded(model.key, !model.answerStarted);

    const title = state === "waiting" ? t("turn.waiting")
      : state === "live" ? t("turn.working")
      : t("turn.workedFor", { duration: elapsed });

    return html`
      <div class="pi-turn" data-state=${state}>
        <button
          type="button"
          class="pi-turn__header"
          aria-expanded=${expanded ? "true" : "false"}
          @click=${() => setDisclosureExpanded(model.key, !expanded)}
        >
          <span class="pi-turn__stack" aria-hidden="true">
            ${kinds.map((kind) => html`<span class="pi-turn__kind">${icon(toolKindIcon(kind), "xs")}</span>`)}
          </span>
          <span class="pi-turn__title">${title}</span>
          ${state === "live" ? html`<span class="pi-turn__spinner" aria-hidden="true">${icon(Loader, "sm")}</span>` : nothing}
          ${failed > 0
            ? html`<span class="pi-turn__failed">${icon(AlertTriangle, "xs")}${t("turn.failedCount", { count: failed })}</span>`
            : state === "live" ? html`<span class="pi-turn__elapsed">${elapsed}</span>` : nothing}
          <span class="pi-turn__chevron" aria-hidden="true">${icon(expanded ? ChevronUp : ChevronDown, "sm")}</span>
        </button>
        ${expanded ? html`<div class="pi-turn__rows">${model.steps.map((step) =>
          step.type === "tool" ? this._renderToolRow(step) : this._renderNote(step))}</div>` : nothing}
        ${model.changes ? renderTurnChangesReceipt(model.changes, "footer") : nothing}
      </div>
    `;
  }

  private _renderNote(step: TurnNoteStep): TemplateResult {
    const chunk = step.message.content[step.contentIndex];
    if (chunk?.type === "thinking") {
      const key = thinkingKey(step.message, step.contentIndex);
      return html`<div class="pi-turn-row pi-turn-row--note">
        <thinking-block
          .content=${chunk.thinking}
          .isStreaming=${step.streaming}
          .disclosureKey=${key}
          .durationMs=${getThinkingDurationMs(key)}
        ></thinking-block>
      </div>`;
    }
    if (chunk?.type === "text") {
      return html`<div class="pi-turn-row pi-turn-row--note"><markdown-block .content=${chunk.text}></markdown-block></div>`;
    }
    return html``;
  }

  private _renderToolRow(step: TurnToolStep): TemplateResult {
    const approval = this._approvalFor(step.toolCall.id);
    if (approval) return this._renderApprovalRow(step, approval);

    const tense: ToolRowTense = step.state;
    const row = describeToolRow(step.toolCall.name, step.toolCall.arguments, step.result, tense, step.tool?.label);
    const status = STATUS_ICONS[step.state];
    const expanded = isDisclosureExpanded(step.key);
    const toggle = (): void => setDisclosureExpanded(step.key, !expanded);

    return html`
      <div class="pi-turn-row" data-state=${step.state} data-tool-name=${step.toolCall.name}>
        <div
          class="pi-turn-row__line"
          role="button"
          tabindex="0"
          aria-expanded=${expanded ? "true" : "false"}
          @click=${toggle}
          @keydown=${(event: KeyboardEvent) => {
            if (event.key !== "Enter" && event.key !== " ") return;
            event.preventDefault();
            toggle();
          }}
        >
          <span class=${status.className} aria-hidden="true">${icon(status.icon, "xs")}</span>
          <span class="pi-turn-row__sentence">${row.sentence}</span>
          ${row.meta ? html`<span class="pi-turn-row__meta">${row.meta}</span>` : nothing}
          <span class="pi-turn-row__chevron" aria-hidden="true">${icon(expanded ? ChevronUp : ChevronDown, "xs")}</span>
        </div>
        ${row.reason ? html`<div class="pi-turn-row__reason">${row.reason}</div>` : nothing}
        ${hasToolSetupCard(step.result) ? html`<div class="pi-turn-row__setup">${renderToolSetupCards(step.result)}</div>` : nothing}
        ${expanded ? html`<div class="pi-turn-row__detail">${renderToolRowDetail(step.toolCall.name, step.toolCall.arguments, step.result)}</div>` : nothing}
      </div>
    `;
  }

  private _renderApprovalRow(step: TurnToolStep, approval: InlineApproval): TemplateResult {
    const { request } = approval;
    const status = STATUS_ICONS.waiting;
    const consequence = [request.overwrites, request.note].filter((part): part is string => typeof part === "string" && part !== "");
    const prepare = (): void => {
      void approval.prepareAndApprove().catch(() => {
        this._preparationFailed = approval;
        this.requestUpdate();
      });
    };
    return html`
      <div class="pi-turn-row pi-turn-row--approval" data-state="waiting" data-tool-name=${step.toolCall.name} role="group" aria-label=${request.title}>
        <div class="pi-turn-row__line">
          <span class=${status.className} aria-hidden="true">${icon(status.icon, "xs")}</span>
          <span class="pi-turn-row__sentence">${request.headline ?? request.title}</span>
        </div>
        ${request.details ? html`<div class="pi-turn-row__approval-details">${request.details}</div>`
          : request.message ? html`<div class="pi-turn-row__sub pi-turn-row__sub--pre">${request.message}</div>` : nothing}
        ${consequence.length > 0 ? html`<div class="pi-turn-row__sub">${consequence.map((line) => html`<span>${line}</span>`)}</div>` : nothing}
        ${this._preparationFailed === approval ? html`<div class="pi-turn-row__reason" role="alert">${t("approval.preparationFailed")}</div>` : nothing}
        <div class="pi-turn-row__actions">
          <button type="button" class="pi-btn pi-btn--secondary" @click=${() => approval.answer(false)}>${t("approval.decline")}</button>
          <button type="button" class="pi-btn pi-btn--primary" @click=${() => approval.answer(true)}>${request.confirmLabel}</button>
        </div>
        ${request.allowForTurn && request.signal
          ? html`<button type="button" class="pi-turn-row__link" @click=${() => approval.allowForTurn()}>${t("approval.allowForTurn")}</button>`
          : nothing}
        ${request.secondaryApproval
          ? html`<button type="button" class="pi-turn-row__link" @click=${prepare}>${request.secondaryApproval.label}</button>`
          : nothing}
      </div>
    `;
  }

  private _preparationFailed: InlineApproval | null = null;
}

declare global {
  interface HTMLElementTagNameMap {
    "turn-group": TurnGroup;
  }
}
