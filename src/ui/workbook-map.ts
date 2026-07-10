/**
 * Workbook map dialog — "/map" X-ray overlay.
 *
 * Shows a minimap of one sheet where every cell is colored by what it is
 * (formula / hardcoded input / text label / error), with sheet tabs,
 * hover-to-inspect, and click-to-select-in-Excel.
 */

import { t } from "../language/index.js";
import {
  WorkbookMapUnsupportedError,
  listVisibleSheets,
  scanSheetMap,
  selectSheetCell,
} from "../workbook/map-scan.js";
import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { WORKBOOK_MAP_OVERLAY_ID } from "./overlay-ids.js";
import { showToast } from "./toast.js";
import { createWorkbookMapView } from "./workbook-map-view.js";

/** Open (or toggle closed) the workbook map overlay. */
export function showWorkbookMapDialog(): void {
  if (closeOverlayById(WORKBOOK_MAP_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: WORKBOOK_MAP_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-map-dialog",
  });

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: t("map.closeLabel"),
    title: t("map.title"),
    subtitle: t("map.subtitle"),
  });

  const toolbar = document.createElement("div");
  toolbar.className = "pi-map-toolbar";
  toolbar.hidden = true;

  const tabs = document.createElement("div");
  tabs.className = "pi-map-tabs";

  const refreshButton = createOverlayButton({
    text: t("map.refresh"),
    className: "pi-overlay-btn--compact pi-map-refresh",
  });

  toolbar.append(tabs, refreshButton);

  const view = createWorkbookMapView({
    maxWidth: 300,
    maxHeight: 340,
    onCellClick: ({ sheetName, address }) => {
      void (async () => {
        try {
          await selectSheetCell(sheetName, address);
          view.setStatus(t("map.status.selected", { address }));
        } catch {
          showToast(t("map.toast.selectFailed"));
        }
      })();
    },
  });

  dialog.card.append(header, toolbar, view.root);

  let disposed = false;
  dialog.addCleanup(() => {
    disposed = true;
  });

  dialog.mount();

  let activeSheet: string | undefined;
  let scanToken = 0;
  const tabButtons = new Map<string, HTMLButtonElement>();

  const syncTabs = (): void => {
    for (const [name, button] of tabButtons) {
      const isActive = name === activeSheet;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-pressed", String(isActive));
    }
  };

  const runScan = (sheetName?: string): void => {
    scanToken += 1;
    const token = scanToken;
    view.setStatus(t("map.loading"));

    void (async () => {
      try {
        const scan = await scanSheetMap(sheetName);
        if (disposed || token !== scanToken) return;
        activeSheet = scan.sheetName;
        syncTabs();
        view.renderScan(scan);
        view.setStatus(t("map.hint"));
      } catch (error) {
        if (disposed || token !== scanToken) return;
        view.setStatus(
          error instanceof WorkbookMapUnsupportedError ? t("map.unsupportedApi") : t("map.scanFailed"),
        );
      }
    })();
  };

  refreshButton.addEventListener("click", () => {
    runScan(activeSheet);
  });

  void (async () => {
    try {
      const sheets = await listVisibleSheets();
      if (disposed) return;

      activeSheet = activeSheet ?? sheets.activeSheetName;

      if (sheets.sheetNames.length > 1) {
        for (const name of sheets.sheetNames) {
          const tab = document.createElement("button");
          tab.type = "button";
          tab.className = "pi-map-tab";
          tab.textContent = name;
          tab.addEventListener("click", () => {
            if (activeSheet === name) return;
            activeSheet = name;
            syncTabs();
            runScan(name);
          });
          tabButtons.set(name, tab);
          tabs.appendChild(tab);
        }
        toolbar.hidden = false;
        syncTabs();
      }
    } catch {
      // Sheet tabs are an enhancement; the active-sheet scan still works.
    }
  })();

  runScan();
}
