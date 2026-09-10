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
import { createSpreadsheetHost, detectSpreadsheetHostKind } from "../host/index.js";
import type { SpreadsheetHost } from "../host/types.js";
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
  installProcessEnvShim();
  installFetchInterceptor();
  installModelSelectorPatch();

  // Office bootstrap (with fallback for local dev)
  let initialized = false;

  const runInit = (host: SpreadsheetHost = createSpreadsheetHost(detectSpreadsheetHostKind())) => {
    if (initialized) return;

    initialized = true;

    let initComplete = false;

    const markInitComplete = () => {
      if (initComplete) return false;
      initComplete = true;
      return true;
    };

    const slowInitTimer = setTimeout(() => {
      if (initComplete) return;
      console.warn("[pi] Taskpane initialization is taking longer than expected (>12s)");
    }, 12_000);

    const hardTimeoutTimer = setTimeout(() => {
      if (!markInitComplete()) return;
      loadingRoot.innerHTML = "";
      showFatalError(
        errorRoot,
        "Failed to initialize: Taskpane initialization timed out after 60000ms",
      );
      console.error("[pi] Init error: Taskpane initialization timed out after 60000ms");
    }, 60_000);

    void initTaskpane({ appEl, errorRoot, host })
      .then(() => {
        if (!markInitComplete()) return;
        clearTimeout(slowInitTimer);
        clearTimeout(hardTimeoutTimer);
      })
      .catch((error: unknown) => {
        if (!markInitComplete()) {
          console.error("[pi] Init error after timeout:", error);
          return;
        }

        clearTimeout(slowInitTimer);
        clearTimeout(hardTimeoutTimer);
        loadingRoot.innerHTML = "";
        showFatalError(errorRoot, `Failed to initialize: ${getErrorMessage(error)}`);
        console.error("[pi] Init error:", error);
      });
  };

  const bootHost = createSpreadsheetHost(detectSpreadsheetHostKind());

  const finishHostReady = (host: SpreadsheetHost): void => {
    host.ready()
      .then((info) => {
        if (info.kind === "office") {
          console.log(`[pi] Office.js ready: host=${info.host ?? "unknown"}, platform=${info.platform ?? "unknown"}`);
        } else if (info.kind === "wps") {
          console.log("[pi] WPS host ready");
        } else {
          console.warn("[pi] Office.js is unavailable — initializing without Excel");
        }
        runInit(host);
      })
      .catch((error: unknown) => {
        console.warn("[pi] Host readiness failed — initializing in browser mode", error);
        runInit(createSpreadsheetHost("browser"));
      });
  };

  if (bootHost.kind !== "office") {
    finishHostReady(bootHost);
    return;
  }

  finishHostReady(bootHost);

  setTimeout(() => {
    if (initialized) return;

    console.warn("[pi] Office.js not ready after 3s — initializing without Excel");
    runInit(createSpreadsheetHost("browser"));
  }, 3000);
}
