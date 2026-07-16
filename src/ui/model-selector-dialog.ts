/**
 * Model selector dialog — first-party replacement for pi-web-ui's
 * `ModelSelector` (docs/ui-ownership.md step 6).
 *
 * Differences from upstream (intentional):
 * - No Thinking/Vision filter pills, capability icons, or cost column — the
 *   previous theme already hid all three; they are simply not rendered.
 * - Provider catalogues come from the browser-native Pi AI Models runtime.
 *   Cached dynamic models render immediately and background refreshes update
 *   an already-open selector.
 * - Provider filtering (active credentials) and featured-model ordering are
 *   built in rather than monkey-patched (src/models/featured-models.ts).
 */

import {
  modelsAreEqual,
  type Api,
  type Model,
  type Models,
} from "@earendil-works/pi-ai";

import { t } from "../language/index.js";
import { getActiveProviders } from "../models/active-providers.js";
import {
  orderModelsForSelector,
  subsequenceScore,
  type ModelSelectorItem,
} from "../models/featured-models.js";
import { MODEL_SELECTOR_OVERLAY_ID } from "./overlay-ids.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";

export interface ModelSelectorDialogOptions {
  models: Models;
  currentModel: Model<Api> | null;
  onSelect: (model: Model<Api>) => void;
}

function formatTokenCount(count: number): string {
  if (count >= 1_000_000) return `${Math.round(count / 1_000_000)}M`;
  if (count >= 1_000) return `${Math.round(count / 1_000)}k`;
  return String(count);
}

function collectModelItems(modelsRuntime: Models): ModelSelectorItem[] {
  const items: ModelSelectorItem[] = [];
  for (const provider of modelsRuntime.getProviders()) {
    for (const model of modelsRuntime.getModels(provider.id)) {
      items.push({ provider: provider.id, id: model.id, model });
    }
  }
  return items;
}

export function openModelSelectorDialog(options: ModelSelectorDialogOptions): void {
  closeOverlayById(MODEL_SELECTOR_OVERLAY_ID);

  const dialog = createOverlayDialog({
    overlayId: MODEL_SELECTOR_OVERLAY_ID,
    cardClassName: "pi-model-selector-card",
    restoreFocusOnClose: true,
  });

  let allItems = collectModelItems(options.models);
  let searchQuery = "";
  let selectedIndex = 0;
  // Selection follows the mouse only after real pointer movement, so
  // keyboard-triggered scrolling doesn't hand selection to whatever lands
  // under the cursor.
  let navigationMode: "keyboard" | "mouse" = "mouse";
  let lastMouse = { x: 0, y: 0 };

  const { header } = createOverlayHeader({
    title: t("modelSelector.title"),
    onClose: dialog.close,
    closeLabel: t("dialog.close"),
  });
  header.classList.add("pi-model-selector-header");

  const searchInput = document.createElement("input");
  searchInput.type = "text";
  searchInput.className = "pi-overlay-input pi-model-selector-search";
  searchInput.placeholder = t("modelSelector.searchPlaceholder");
  searchInput.setAttribute("aria-label", t("modelSelector.searchPlaceholder"));
  searchInput.autocomplete = "off";
  searchInput.spellcheck = false;

  const list = document.createElement("div");
  list.className = "pi-model-selector-list";
  list.setAttribute("role", "listbox");

  const getFilteredItems = (): ModelSelectorItem[] => {
    let filtered = allItems;

    const active = getActiveProviders();
    if (active) {
      filtered = filtered.filter((item) => active.has(item.provider));
    }

    const query = searchQuery.toLowerCase().replace(/\s+/g, "");
    if (query) {
      const scored: Array<{ item: ModelSelectorItem; score: number }> = [];
      for (const item of filtered) {
        const searchText = `${item.provider} ${item.id} ${item.model.name}`.toLowerCase();
        const score = subsequenceScore(query, searchText);
        if (score > 0) {
          scored.push({ item, score });
        }
      }
      scored.sort((a, b) => b.score - a.score);

      // Preserve score order, but float the current model to the top.
      const current = options.currentModel;
      const items = scored.map((entry) => entry.item);
      if (current) {
        items.sort((a, b) => {
          const aIsCurrent = modelsAreEqual(current, a.model) ? 0 : 1;
          const bIsCurrent = modelsAreEqual(current, b.model) ? 0 : 1;
          return aIsCurrent - bIsCurrent;
        });
      }
      return items;
    }

    return orderModelsForSelector(filtered, options.currentModel);
  };

  const scrollSelectedIntoView = (): void => {
    requestAnimationFrame(() => {
      const selected = list.querySelector(".pi-model-selector-item--selected");
      selected?.scrollIntoView({ block: "nearest" });
    });
  };

  const select = (item: ModelSelectorItem): void => {
    dialog.close();
    options.onSelect(item.model);
  };

  const renderList = (): void => {
    const filtered = getFilteredItems();
    selectedIndex = Math.max(0, Math.min(selectedIndex, filtered.length - 1));

    list.replaceChildren();

    if (filtered.length === 0) {
      const empty = document.createElement("div");
      empty.className = "pi-model-selector-empty";
      empty.textContent = t("modelSelector.noResults");
      list.appendChild(empty);
      return;
    }

    filtered.forEach((item, index) => {
      const isCurrent = Boolean(
        options.currentModel && modelsAreEqual(options.currentModel, item.model),
      );

      const row = document.createElement("button");
      row.type = "button";
      row.className = "pi-model-selector-item";
      row.classList.toggle("pi-model-selector-item--selected", index === selectedIndex);
      row.setAttribute("role", "option");
      row.setAttribute("aria-selected", index === selectedIndex ? "true" : "false");

      const topRow = document.createElement("span");
      topRow.className = "pi-model-selector-item-top";

      const idEl = document.createElement("span");
      idEl.className = "pi-model-selector-item-id";
      idEl.textContent = item.id;
      topRow.appendChild(idEl);

      if (isCurrent) {
        const check = document.createElement("span");
        check.className = "pi-model-selector-item-check";
        check.textContent = "✓";
        check.setAttribute("aria-label", t("modelSelector.current"));
        topRow.appendChild(check);
      }

      const providerEl = document.createElement("span");
      providerEl.className = "pi-model-selector-item-provider";
      providerEl.textContent = item.provider;
      topRow.appendChild(providerEl);

      const metaRow = document.createElement("span");
      metaRow.className = "pi-model-selector-item-meta";
      metaRow.textContent = `${formatTokenCount(item.model.contextWindow)} ctx · ${formatTokenCount(item.model.maxTokens)} out`;

      row.append(topRow, metaRow);

      row.addEventListener("click", () => {
        select(item);
      });
      row.addEventListener("mouseenter", () => {
        if (navigationMode !== "mouse") return;
        if (selectedIndex === index) return;
        selectedIndex = index;
        renderList();
      });

      list.appendChild(row);
    });
  };

  searchInput.addEventListener("input", () => {
    searchQuery = searchInput.value;
    selectedIndex = 0;
    list.scrollTop = 0;
    renderList();
  });

  dialog.overlay.addEventListener("mousemove", (event) => {
    if (event.clientX === lastMouse.x && event.clientY === lastMouse.y) return;
    lastMouse = { x: event.clientX, y: event.clientY };
    navigationMode = "mouse";
  });

  dialog.overlay.addEventListener("keydown", (event) => {
    if (event.isComposing || event.key === "Process") return;

    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      navigationMode = "keyboard";
      const filtered = getFilteredItems();
      const delta = event.key === "ArrowDown" ? 1 : -1;
      selectedIndex = Math.max(0, Math.min(selectedIndex + delta, filtered.length - 1));
      renderList();
      scrollSelectedIntoView();
      return;
    }

    if (event.key === "Enter") {
      event.preventDefault();
      const filtered = getFilteredItems();
      const item = filtered[selectedIndex];
      if (item) {
        select(item);
      }
    }
  });

  dialog.card.append(header, searchInput, list);
  const handleModelsChanged = (): void => {
    if (!dialog.overlay.isConnected) return;
    allItems = collectModelItems(options.models);
    renderList();
  };
  document.addEventListener("pi:models-changed", handleModelsChanged);
  dialog.addCleanup(() => {
    document.removeEventListener("pi:models-changed", handleModelsChanged);
  });

  dialog.mount();
  renderList();
  searchInput.focus();
}
