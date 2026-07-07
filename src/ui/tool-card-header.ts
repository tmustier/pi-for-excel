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
      <span class="pi-tool-card__toggle-chevron" ${ref(chevronRef)} aria-hidden="true">
        <span class="pi-tool-card__chevron-up ${defaultExpanded ? "" : "hidden"}">
          ${icon(ChevronUp, "sm")}
        </span>
        <span class="pi-tool-card__chevron-collapsed ${defaultExpanded ? "hidden" : ""}">
          ${icon(ChevronsUpDown, "sm")}
        </span>
      </span>
    </button>
  `;
}
