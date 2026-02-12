/**
 * Taskpane bootstrap.
 *
 * Runs immediately when the add-in loads:
 * - renders loading UI
 * - installs global fetch + UI monkey patches
 * - waits for Office.onReady (with fallback) and then initializes the app
 */

import { render } from "lit";

import { installFetchInterceptor } from "../auth/cors-proxy.js";
import { installModelSelectorPatch } from "../compat/model-selector-patch.js";
import { renderLoading, renderError } from "../ui/loading.js";
import { getErrorMessage } from "../utils/errors.js";

import { initTaskpane } from "./init.js";

function getRequiredElement<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) {
    throw new Error(`[pi] Missing required element #${id}`);
  }
  return el as T;
}

function showFatalError(errorRoot: HTMLElement, message: string): void {
  render(renderError(message), errorRoot);
}

export function bootstrapTaskpane(): void {
  const appEl = getRequiredElement<HTMLElement>("app");
  const loadingRoot = getRequiredElement<HTMLElement>("loading-root");
  const errorRoot = getRequiredElement<HTMLElement>("error-root");

  // Initial loading UI
  render(renderLoading(), loadingRoot);

  // Global patches
  installFetchInterceptor();
  installModelSelectorPatch();

  // Office bootstrap (with fallback for local dev)
  let initialized = false;

  void Office.onReady(async (info) => {
    console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);
    try {
      initialized = true;
      await initTaskpane({ appEl, errorRoot });
    } catch (e: unknown) {
      showFatalError(errorRoot, `Failed to initialize: ${getErrorMessage(e)}`);
      console.error("[pi] Init error:", e);
    }
  });

  setTimeout(() => {
    if (initialized) return;

    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    initialized = true;
    initTaskpane({ appEl, errorRoot }).catch((e: unknown) => {
      showFatalError(errorRoot, `Failed to initialize: ${getErrorMessage(e)}`);
      console.error("[pi] Init error:", e);
    });
  }, 3000);
}
