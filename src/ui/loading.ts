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

