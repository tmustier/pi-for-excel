/**
 * Pi for Excel — Loading and error state components.
 *
 * Extracted for easy swapping / versioning.
 */

import { html, type TemplateResult } from "lit";

/**
 * Render the loading spinner.
 */
export function renderLoading(): TemplateResult {
  return html`
    <div class="pi-loading">
      <div class="pi-loading__spinner">
        <div class="pi-loading__ring"></div>
        <div class="pi-loading__ring pi-loading__ring--inner"></div>
      </div>
      <span class="pi-loading__text">Initializing…</span>
    </div>
  `;
}

/**
 * Show an error message. Returns a template that can be rendered into #error.
 */
export function renderError(message: string): TemplateResult {
  return html`<div class="pi-error">${message}</div>`;
}

/**
 * CSS for loading and error states.
 */
export const loadingStyles = `
  .pi-loading {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    flex: 1;
    gap: 16px;
  }

  .pi-loading__spinner {
    position: relative;
    width: 32px;
    height: 32px;
  }

  .pi-loading__ring {
    position: absolute;
    inset: 0;
    border: 2px solid var(--border);
    border-top-color: var(--primary);
    border-radius: 50%;
    animation: pi-spin 0.8s cubic-bezier(0.4, 0, 0.2, 1) infinite;
  }

  .pi-loading__ring--inner {
    inset: 4px;
    border-top-color: transparent;
    border-right-color: oklch(0.55 0.10 100 / 0.5);
    animation-duration: 1.2s;
    animation-direction: reverse;
  }

  .pi-loading__text {
    font-family: var(--font-mono);
    font-size: 11px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--muted-foreground);
  }

  @keyframes pi-spin {
    to { transform: rotate(360deg); }
  }

  .pi-error {
    padding: 12px 14px;
    font-size: 12px;
    font-family: var(--font-mono);
    background: oklch(0.95 0.03 25);
    color: oklch(0.45 0.18 25);
    border-bottom: 1px solid oklch(0.88 0.06 25);
  }
`;
