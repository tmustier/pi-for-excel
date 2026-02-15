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
import { installProcessEnvShim } from "../compat/process-env-shim.js";
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

async function awaitWithTimeout<T>(label: string, timeoutMs: number, task: Promise<T>): Promise<T> {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;

  try {
    const timeoutPromise = new Promise<never>((_, reject) => {
      timeoutId = setTimeout(() => {
        reject(new Error(`${label} timed out after ${timeoutMs}ms`));
      }, timeoutMs);
    });

    return await Promise.race([task, timeoutPromise]);
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
  }
}

export function bootstrapTaskpane(): void {
  const appEl = getRequiredElement<HTMLElement>("app");
  const loadingRoot = getRequiredElement<HTMLElement>("loading-root");
  const errorRoot = getRequiredElement<HTMLElement>("error-root");

  // Initial loading UI
  render(renderLoading(), loadingRoot);

  // Global patches
  installProcessEnvShim();
  installFetchInterceptor();
  installModelSelectorPatch();

  // Office bootstrap (with fallback for local dev)
  let initialized = false;

  const runInit = () => {
    if (initialized) return;

    initialized = true;

    void awaitWithTimeout(
      "Taskpane initialization",
      12_000,
      initTaskpane({ appEl, errorRoot }),
    ).catch((error: unknown) => {
      loadingRoot.innerHTML = "";
      showFatalError(errorRoot, `Failed to initialize: ${getErrorMessage(error)}`);
      console.error("[pi] Init error:", error);
    });
  };

  if (typeof Office === "undefined") {
    console.warn("[pi] Office.js is unavailable — initializing without Excel");
    runInit();
    return;
  }

  void Office.onReady((info) => {
    console.log(`[pi] Office.js ready: host=${info.host}, platform=${info.platform}`);
    runInit();
  });

  setTimeout(() => {
    if (initialized) return;

    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    runInit();
  }, 3000);
}
