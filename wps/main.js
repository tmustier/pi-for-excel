/* global Application, wps */

(function registerPiForWpsRibbon(globalScope) {
  "use strict";

  const DEV_TASKPANE_URL = "https://localhost:3141/src/taskpane.html";
  const PROD_TASKPANE_URL = "https://pi-for-excel.vercel.app/src/taskpane.html";
  const TASKPANE_ID_STORAGE_KEY = "pi.taskpane.id";

  let taskpane = null;

  function getConfiguredTaskpaneUrl() {
    const configured = globalScope.PI_WPS_TASKPANE_URL;
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

  globalScope.OnAddInLoad = function OnAddInLoad() {
    storageGet(TASKPANE_ID_STORAGE_KEY);
  };

  // WPS template skeletons spell this callback `OnAddinLoad`; keep the
  // Office-style camel-cased alias too so either ribbon spelling works.
  globalScope.OnAddinLoad = globalScope.OnAddInLoad;

  globalScope.OpenPi = function OpenPi() {
    if (taskpane) {
      taskpane.Visible = !taskpane.Visible;
      return;
    }

    const url = getConfiguredTaskpaneUrl();
    taskpane = createTaskpane(url);
    rememberTaskpaneId(taskpane);
    taskpane.Visible = true;
  };

  globalScope.PI_WPS_TASKPANE_URLS = {
    dev: DEV_TASKPANE_URL,
    prod: PROD_TASKPANE_URL,
  };
})(typeof globalThis !== "undefined" ? globalThis : window);
