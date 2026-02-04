/**
 * Pi for Excel — Header bar component.
 *
 * Renders the branded header with logo and status indicator.
 * Extracted for easy swapping / versioning.
 */

import { html, type TemplateResult } from "lit";

export interface HeaderState {
  status: "ready" | "working" | "error";
  statusText?: string;
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

  return html`
    <div class="pi-header">
      <div class="pi-header__logo">
        <span class="pi-header__mark">π</span>
        <span class="pi-header__text">for <span class="pi-header__accent">Excel</span></span>
      </div>
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
    border-bottom: 1px solid var(--border);
    background: var(--background);
    flex-shrink: 0;
    position: relative;
    z-index: 10;
  }

  /* Accent line */
  .pi-header::after {
    content: '';
    position: absolute;
    bottom: -1px;
    left: 0;
    right: 0;
    height: 1px;
    background: linear-gradient(90deg, oklch(0.50 0.14 155 / 0.5), oklch(0.55 0.10 100 / 0.2), transparent 80%);
  }

  .pi-header__logo {
    display: flex;
    align-items: center;
    gap: 8px;
  }

  .pi-header__mark {
    font-family: 'DM Sans', serif;
    font-size: 22px;
    font-weight: 700;
    color: var(--primary);
    line-height: 1;
    letter-spacing: -0.02em;
  }

  .pi-header__text {
    font-family: var(--font-mono);
    font-size: 11px;
    font-weight: 400;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: var(--muted-foreground);
  }

  .pi-header__accent {
    color: var(--foreground);
    font-weight: 500;
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
