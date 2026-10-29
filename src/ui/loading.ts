import { html, type TemplateResult } from "lit";
import { t } from "../language/index.js";

export function renderLoading(message?: string): TemplateResult {
  return html`
    <div class="pi-loading">
      <div class="pi-loading__spinner">
        <div class="pi-loading__ring"></div>
        <div class="pi-loading__ring pi-loading__ring--inner"></div>
      </div>
      <span class="pi-loading__text">${message ?? t("loading.initializing")}</span>
    </div>
  `;
}

export function renderError(message: string): TemplateResult {
  return html`<div class="pi-error">${message}</div>`;
}

