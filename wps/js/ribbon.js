/* global Application, wps */

"use strict";

const DEV_TASKPANE_URL = "https://localhost:3141/src/taskpane.html";
const PROD_TASKPANE_URL = "https://pi-for-excel.vercel.app/src/taskpane.html";
const TASKPANE_ID_STORAGE_KEY = "pi.taskpane.id";

let piTaskpane = null;
let piRibbonUI = null;

function getGlobalScope() {
  if (typeof globalThis !== "undefined") return globalThis;
  if (typeof window !== "undefined") return window;
  return this;
}

function getConfiguredTaskpaneUrl() {
  const globalScope = getGlobalScope();
  const configured = globalScope && globalScope.PI_WPS_TASKPANE_URL;
  if (typeof configured === "string" && configured.trim().length > 0) {
    return configured.trim();
  }

  return DEV_TASKPANE_URL;
}

function getApplication() {
  if (typeof Application !== "undefined" && Application) return Application;
  if (typeof wps !== "undefined" && wps && typeof wps.EtApplication === "function") {
    return wps.EtApplication();
  }
  return null;
}

function getPluginStorage() {
  const app = getApplication();
  if (app && app.PluginStorage) return app.PluginStorage;
  if (typeof wps !== "undefined" && wps && wps.PluginStorage) return wps.PluginStorage;
  return null;
}

function storageGet(key) {
  const storage = getPluginStorage();
  if (!storage || typeof storage.getItem !== "function") return null;
  const value = storage.getItem(key);
  return typeof value === "string" && value.length > 0 ? value : null;
}

function storageSet(key, value) {
  const storage = getPluginStorage();
  if (!storage || typeof storage.setItem !== "function") return;
  storage.setItem(key, value);
}

function createTaskpane(url) {
  if (typeof wps !== "undefined" && wps && typeof wps.CreateTaskPane === "function") {
    return wps.CreateTaskPane(url);
  }
  if (typeof wps !== "undefined" && wps && typeof wps.CreateTaskpane === "function") {
    return wps.CreateTaskpane(url);
  }

  const app = getApplication();
  if (app && typeof app.CreateTaskpane === "function") {
    return app.CreateTaskpane(url);
  }
  if (app && typeof app.CreateTaskPane === "function") {
    return app.CreateTaskPane(url);
  }

  throw new Error("WPS TaskPane API is unavailable.");
}

function rememberTaskpaneId(pane) {
  if (!pane || pane.ID === undefined || pane.ID === null) return;
  storageSet(TASKPANE_ID_STORAGE_KEY, String(pane.ID));
}

function OnAddInLoad(ribbonUI) {
  piRibbonUI = ribbonUI || null;
  const app = getApplication();
  if (app && typeof app === "object" && !app.ribbonUI) {
    app.ribbonUI = ribbonUI;
  }
  storageGet(TASKPANE_ID_STORAGE_KEY);
  return true;
}

// WPS template skeletons spell this callback `OnAddinLoad`.
// Keep the alias as a real top-level function so either ribbon spelling works.
function OnAddinLoad(ribbonUI) {
  return OnAddInLoad(ribbonUI);
}

function OnGetEnabled() {
  return true;
}

function OnGetVisible() {
  return true;
}

function GetImage() {
  return "pi.svg";
}

function openPiTaskpane() {
  const existingTaskpaneId = storageGet(TASKPANE_ID_STORAGE_KEY);
  if (piTaskpane) {
    piTaskpane.Visible = !piTaskpane.Visible;
    return true;
  }

  const app = getApplication();
  if (existingTaskpaneId && app && typeof app.GetTaskPane === "function") {
    try {
      piTaskpane = app.GetTaskPane(existingTaskpaneId);
      if (piTaskpane) {
        piTaskpane.Visible = !piTaskpane.Visible;
        return true;
      }
    } catch {
      storageSet(TASKPANE_ID_STORAGE_KEY, "");
    }
  }

  const url = getConfiguredTaskpaneUrl();
  piTaskpane = createTaskpane(url);
  rememberTaskpaneId(piTaskpane);
  piTaskpane.Visible = true;
  return true;
}

function showTaskpaneError(error) {
  const message = error && error.message ? error.message : String(error);
  if (typeof alert === "function") {
    alert(`Could not open Pi taskpane: ${message}`);
  }
  return true;
}

function OnAction(control) {
  const id = control && (control.Id || control.id);
  if (!id || id === "PiOpenTaskpaneButton") {
    try {
      return openPiTaskpane();
    } catch (error) {
      return showTaskpaneError(error);
    }
  }
  return true;
}

function OpenPi() {
  try {
    return openPiTaskpane();
  } catch (error) {
    return showTaskpaneError(error);
  }
}

const globalScope = getGlobalScope();
if (globalScope) {
  const ribbonCallbacks = {
    OnAddInLoad,
    OnAddinLoad,
    OnGetEnabled,
    OnGetVisible,
    GetImage,
    OnAction,
    OpenPi,
  };

  // WPS' official template prefixes callbacks with `ribbon.` in ribbon.xml.
  // Keep the legacy globals as aliases for older generated manifests and local
  // smoke packages.
  globalScope.ribbon = {
    ...(globalScope.ribbon && typeof globalScope.ribbon === "object" ? globalScope.ribbon : {}),
    ...ribbonCallbacks,
  };
  globalScope.OnAddInLoad = OnAddInLoad;
  globalScope.OnAddinLoad = OnAddinLoad;
  globalScope.OnGetEnabled = OnGetEnabled;
  globalScope.OnGetVisible = OnGetVisible;
  globalScope.GetImage = GetImage;
  globalScope.OnAction = OnAction;
  globalScope.OpenPi = OpenPi;
  globalScope.PI_WPS_TASKPANE_URLS = {
    dev: DEV_TASKPANE_URL,
    prod: PROD_TASKPANE_URL,
  };
  globalScope.PI_WPS_RIBBON_STATE = function PI_WPS_RIBBON_STATE() {
    return {
      hasRibbonUI: !!piRibbonUI,
      hasTaskpane: !!piTaskpane,
      taskpaneUrl: getConfiguredTaskpaneUrl(),
    };
  };
}
