/**
 * Pi for Excel — Header bar component.
 *
 * Renders the header with model alias (clickable to change) and status dot.
 */

import { html, type TemplateResult } from "lit";

export interface HeaderState {
  status: "ready" | "working" | "error";
  statusText?: string;
  modelAlias?: string;
  onModelClick?: () => void;
}

const STATUS_CONFIG = {
  ready: { color: "var(--primary)", label: "Ready" },
  working: { color: "oklch(0.70 0.15 85)", label: "Working…" },
  error: { color: "oklch(0.60 0.22 25)", label: "Error" },
} as const;

/**
 * Render the header bar.
 */
export function renderHeader(state: HeaderState = { status: "ready" }): TemplateResult {
  const cfg = STATUS_CONFIG[state.status];
  const label = state.statusText ?? cfg.label;
  const model = state.modelAlias || "Select model";

  return html`
    <div class="pi-header">
      <button class="pi-header__model" @click=${state.onModelClick} title="Change model">
        <span class="pi-header__mark">π</span>
        <span class="pi-header__model-name">${model}</span>
        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>
      </button>
      <div class="pi-header__status">
        <span class="pi-header__dot" style="background: ${cfg.color}; box-shadow: 0 0 6px ${cfg.color};"></span>
        <span class="pi-header__label">${label}</span>
      </div>
    </div>
  `;
}

/**
 * CSS for the header. Injected into the document once.
 */
export const headerStyles = `
  .pi-header {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 10px 14px;
    padding-right: 44px;
    /* Grey to blend with Excel's chrome — warm it slightly toward our palette */
    background: oklch(0.92 0.004 90);
    border-bottom: none;
    flex-shrink: 0;
    position: relative;
    z-index: 10;
  }

  /* Soft gradient fade from header grey → content warm-white */
  .pi-header::after {
    content: '';
    position: absolute;
    bottom: -12px;
    left: 0;
    right: 0;
    height: 12px;
    background: linear-gradient(to bottom, oklch(0.92 0.004 90), transparent);
    pointer-events: none;
    z-index: 1;
  }

  .pi-header__model {
    display: flex;
    align-items: center;
    gap: 6px;
    background: none;
    border: none;
    cursor: pointer;
    padding: 4px 8px;
    margin: -4px -8px;
    border-radius: 6px;
    transition: background 0.15s;
    color: var(--foreground);
    min-width: 0;
  }
  .pi-header__model:hover {
    background: oklch(0 0 0 / 0.04);
  }

  .pi-header__mark {
    font-family: 'DM Sans', serif;
    font-size: 18px;
    font-weight: 700;
    color: var(--primary);
    line-height: 1;
    letter-spacing: -0.02em;
    flex-shrink: 0;
  }

  .pi-header__model-name {
    font-family: var(--font-mono);
    font-size: 12px;
    font-weight: 500;
    color: var(--foreground);
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    min-width: 0;
  }

  .pi-header__model svg {
    flex-shrink: 0;
    color: var(--muted-foreground);
  }

  .pi-header__status {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 6px;
  }

  .pi-header__dot {
    width: 6px;
    height: 6px;
    border-radius: 50%;
    animation: pi-pulse-dot 2.5s ease-in-out infinite;
  }

  .pi-header__label {
    font-family: var(--font-mono);
    font-size: 10px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--muted-foreground);
  }

  @keyframes pi-pulse-dot {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.4; }
  }
`;
