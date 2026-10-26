import { icon } from "./icons.js";
import { html, type TemplateResult } from "lit";
import { ref } from "lit/directives/ref.js";
import { ChevronUp, ChevronsUpDown, Loader } from "lucide";

export type ToolHeaderState = "inprogress" | "complete" | "error";

interface ElementRef<T extends Element> {
  value?: T;
}

function setChevronState(chevronContainer: HTMLElement, expanded: boolean): void {
  const upIcon = chevronContainer.querySelector<HTMLElement>(".pi-tool-card__chevron-up");
  const collapsedIcon = chevronContainer.querySelector<HTMLElement>(".pi-tool-card__chevron-collapsed");
  if (!upIcon || !collapsedIcon) return;

  upIcon.classList.toggle("hidden", !expanded);
  collapsedIcon.classList.toggle("hidden", expanded);
}

function toggleContent(
  event: Event,
  contentRef: ElementRef<HTMLDivElement>,
  chevronRef: ElementRef<HTMLElement>,
): void {
  event.preventDefault();

  const content = contentRef.value;
  const chevron = chevronRef.value;
  if (!content || !chevron) return;

  const isCollapsed = content.classList.contains("pi-tool-card__body--collapsed");
  content.classList.toggle("pi-tool-card__body--collapsed", !isCollapsed);
  setChevronState(chevron, isCollapsed);
}

function renderStreamingSpinner(state: ToolHeaderState): TemplateResult {
  if (state !== "inprogress") return html``;

  return html`
    <span class="pi-tool-card__spinner" aria-hidden="true">
      ${icon(Loader, "sm")}
    </span>
  `;
}

export function renderToolCardHeader(state: ToolHeaderState, text: TemplateResult): TemplateResult {
  return html`
    <div class="pi-tool-card__toggle pi-tool-card__toggle--static">
      <span class="pi-tool-card__toggle-main">
        ${renderStreamingSpinner(state)}
        ${text}
      </span>
    </div>
  `;
}

export interface ToolCardDisclosure {
  expanded: boolean;
  toggle: () => void;
}

function renderChevron(expanded: boolean): TemplateResult {
  return html`
    <span class="pi-tool-card__chevron-up ${expanded ? "" : "hidden"}">
      ${icon(ChevronUp, "sm")}
    </span>
    <span class="pi-tool-card__chevron-collapsed ${expanded ? "hidden" : ""}">
      ${icon(ChevronsUpDown, "sm")}
    </span>
  `;
}

/**
 * Header for a card whose expanded state lives in transcript state
 * (tool calls): the caller re-renders with the new state after `toggle`.
 */
export function renderDisclosureToolCardHeader(
  state: ToolHeaderState,
  text: TemplateResult,
  disclosure: ToolCardDisclosure,
): TemplateResult {
  return html`
    <button
      type="button"
      class="pi-tool-card__toggle"
      aria-expanded=${disclosure.expanded ? "true" : "false"}
      @click=${(event: Event) => {
        event.preventDefault();
        disclosure.toggle();
      }}
    >
      <span class="pi-tool-card__toggle-main">
        ${renderStreamingSpinner(state)}
        ${text}
      </span>
      <span class="pi-tool-card__toggle-chevron" aria-hidden="true">${renderChevron(disclosure.expanded)}</span>
    </button>
  `;
}

/**
 * Header for a self-contained card (archive and summary cards) that toggles
 * its body in place and does not follow details mode.
 */
export function renderCollapsibleToolCardHeader(
  state: ToolHeaderState,
  text: TemplateResult,
  contentRef: ElementRef<HTMLDivElement>,
  chevronRef: ElementRef<HTMLElement>,
  defaultExpanded = false,
): TemplateResult {
  return html`
    <button
      type="button"
      class="pi-tool-card__toggle"
      @click=${(event: Event) => toggleContent(event, contentRef, chevronRef)}
    >
      <span class="pi-tool-card__toggle-main">
        ${renderStreamingSpinner(state)}
        ${text}
      </span>
      <span class="pi-tool-card__toggle-chevron" ${ref(chevronRef)} aria-hidden="true">${renderChevron(defaultExpanded)}</span>
    </button>
  `;
}
