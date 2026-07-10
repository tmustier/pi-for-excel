/**
 * Workbook map view — canvas minimap + legend for one sheet scan.
 *
 * Pure DOM component: feed it a `SheetMapScan` (live from Office.js or a
 * mock in the UI gallery) and it paints the map. Colors come from the
 * `.pi-map-swatch--*` theme classes so the canvas always matches the legend
 * and follows light/dark mode.
 */

import { cellAddress } from "../excel/helpers.js";
import { t } from "../language/index.js";
import type { MapCellClass, SheetMapScan } from "../workbook/map-scan.js";
import { MAP_PAINT_ORDER } from "../workbook/map-scan.js";
import {
  MAP_GRID_MIN_CELL_SIZE,
  computeMapLayout,
  pointToCell,
  rectToPixels,
  type MapLayout,
} from "./workbook-map-layout.js";

type MapColorKey = MapCellClass | "empty" | "grid";

const LEGEND_ORDER: readonly MapCellClass[] = ["formula", "input", "label", "error"];

const LEGEND_LABEL_KEYS: Record<MapCellClass | "empty", string> = {
  formula: "map.legend.formula",
  input: "map.legend.input",
  label: "map.legend.label",
  error: "map.legend.error",
  empty: "map.legend.empty",
};

/** sRGB fallbacks for hosts where swatch styles cannot be resolved. */
const FALLBACK_COLORS: Record<MapColorKey, string> = {
  formula: "#5b7fc7",
  input: "#d9a441",
  label: "#c3cfc0",
  error: "#c0392b",
  empty: "#faf9f7",
  grid: "rgba(0, 0, 0, 0.07)",
};

export interface WorkbookMapCellTarget {
  sheetName: string;
  /** Local A1 address on that sheet, e.g. "B12". */
  address: string;
}

export interface WorkbookMapViewOptions {
  /** Canvas CSS px budget used when the view cannot measure its container. */
  maxWidth: number;
  /** Canvas CSS px height budget. */
  maxHeight: number;
  /** Invoked when the user clicks a mapped cell. */
  onCellClick?: (target: WorkbookMapCellTarget) => void;
}

export interface WorkbookMapView {
  root: HTMLDivElement;
  renderScan(scan: SheetMapScan): void;
  setStatus(text: string): void;
}

function resolveSwatchColor(swatch: HTMLElement, fallback: string): string {
  const color = getComputedStyle(swatch).backgroundColor;
  if (!color || color === "transparent" || color === "rgba(0, 0, 0, 0)") {
    return fallback;
  }
  return color;
}

export function createWorkbookMapView(options: WorkbookMapViewOptions): WorkbookMapView {
  const root = document.createElement("div");
  root.className = "pi-map-view";

  const wrap = document.createElement("div");
  wrap.className = "pi-map-canvas-wrap";

  const canvas = document.createElement("canvas");
  canvas.className = "pi-map-canvas";
  if (options.onCellClick) {
    canvas.classList.add("pi-map-canvas--clickable");
  }

  const emptyMessage = document.createElement("div");
  emptyMessage.className = "pi-overlay-empty pi-map-empty";
  emptyMessage.textContent = t("map.emptySheet");
  emptyMessage.hidden = true;

  wrap.append(canvas, emptyMessage);

  const status = document.createElement("div");
  status.className = "pi-map-status";

  const legend = document.createElement("div");
  legend.className = "pi-map-legend";

  const swatches = new Map<MapColorKey, HTMLSpanElement>();
  const legendCounts = new Map<MapCellClass | "empty", HTMLSpanElement>();

  for (const cls of [...LEGEND_ORDER, "empty" as const]) {
    const item = document.createElement("span");
    item.className = "pi-map-legend__item";

    const swatch = document.createElement("span");
    swatch.className = `pi-map-swatch pi-map-swatch--${cls}`;
    swatches.set(cls, swatch);

    const label = document.createElement("span");
    label.className = "pi-map-legend__label";
    label.textContent = t(LEGEND_LABEL_KEYS[cls]);

    const count = document.createElement("span");
    count.className = "pi-map-legend__count";
    legendCounts.set(cls, count);

    item.append(swatch, label, count);
    legend.appendChild(item);
  }

  // Hidden probe so the grid-line color is themeable like the others.
  const gridProbe = document.createElement("span");
  gridProbe.className = "pi-map-swatch pi-map-swatch--grid";
  gridProbe.hidden = true;
  legend.appendChild(gridProbe);
  swatches.set("grid", gridProbe);

  root.append(wrap, status, legend);

  let currentScan: SheetMapScan | null = null;
  let currentLayout: MapLayout | null = null;

  const resolveColors = (): Record<MapColorKey, string> => {
    const colors = { ...FALLBACK_COLORS };
    for (const [key, swatch] of swatches) {
      colors[key] = resolveSwatchColor(swatch, FALLBACK_COLORS[key]);
    }
    return colors;
  };

  const measureWidthBudget = (): number => {
    // clientWidth includes padding; keep a small allowance so the canvas
    // never forces horizontal overflow inside the padded wrap.
    const measured = wrap.clientWidth - 18;
    return measured > 40 ? measured : options.maxWidth;
  };

  const paint = (scan: SheetMapScan, layout: MapLayout): void => {
    const bounds = scan.bounds;
    if (!bounds) return;

    const dpr = Math.max(1, Math.min(3, window.devicePixelRatio || 1));
    canvas.width = Math.max(1, Math.round(layout.width * dpr));
    canvas.height = Math.max(1, Math.round(layout.height * dpr));
    canvas.style.width = `${layout.width}px`;
    canvas.style.height = `${layout.height}px`;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const colors = resolveColors();

    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = colors.empty;
    ctx.fillRect(0, 0, layout.width, layout.height);

    for (const cls of MAP_PAINT_ORDER) {
      ctx.fillStyle = colors[cls];
      for (const rect of scan.rects[cls]) {
        const px = rectToPixels(rect, bounds, layout.cellSize);
        ctx.fillRect(px.x, px.y, px.w, px.h);
      }
    }

    if (layout.cellSize >= MAP_GRID_MIN_CELL_SIZE) {
      ctx.strokeStyle = colors.grid;
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let col = 0; col <= bounds.colCount; col += 1) {
        const x = Math.round(col * layout.cellSize) + 0.5;
        ctx.moveTo(x, 0);
        ctx.lineTo(x, layout.height);
      }
      for (let row = 0; row <= bounds.rowCount; row += 1) {
        const y = Math.round(row * layout.cellSize) + 0.5;
        ctx.moveTo(0, y);
        ctx.lineTo(layout.width, y);
      }
      ctx.stroke();
    }
  };

  const renderLegendCounts = (scan: SheetMapScan): void => {
    for (const cls of [...LEGEND_ORDER, "empty" as const]) {
      const count = legendCounts.get(cls);
      if (count) count.textContent = scan.counts[cls].toLocaleString();
    }
  };

  const cellUnderPointer = (event: MouseEvent): { row: number; col: number } | null => {
    if (!currentScan?.bounds || !currentLayout) return null;
    const rect = canvas.getBoundingClientRect();
    return pointToCell(
      event.clientX - rect.left,
      event.clientY - rect.top,
      currentScan.bounds,
      currentLayout.cellSize,
    );
  };

  canvas.addEventListener("mousemove", (event) => {
    const cell = cellUnderPointer(event);
    if (!cell) return;
    status.textContent = cellAddress(cell.col, cell.row + 1);
  });

  canvas.addEventListener("mouseleave", () => {
    status.textContent = t("map.hint");
  });

  canvas.addEventListener("click", (event) => {
    const onCellClick = options.onCellClick;
    if (!onCellClick || !currentScan) return;
    const cell = cellUnderPointer(event);
    if (!cell) return;
    onCellClick({
      sheetName: currentScan.sheetName,
      address: cellAddress(cell.col, cell.row + 1),
    });
  });

  return {
    root,

    renderScan(scan: SheetMapScan): void {
      currentScan = scan;
      renderLegendCounts(scan);

      if (!scan.bounds) {
        currentLayout = null;
        canvas.hidden = true;
        emptyMessage.hidden = false;
        return;
      }

      canvas.hidden = false;
      emptyMessage.hidden = true;

      const layout = computeMapLayout(
        scan.bounds.rowCount,
        scan.bounds.colCount,
        measureWidthBudget(),
        options.maxHeight,
      );
      currentLayout = layout;
      paint(scan, layout);
    },

    setStatus(text: string): void {
      status.textContent = text;
    },
  };
}
