/**
 * Rules editor overlay — rules (user + workbook) and conventions.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  getUserRules,
  getWorkbookRules,
  setUserRules,
  setWorkbookRules,
  USER_RULES_SOFT_LIMIT,
  WORKBOOK_RULES_SOFT_LIMIT,
} from "../../rules/store.js";
import {
  getStoredConventions,
  isBuiltinPresetName,
  normalizeConventionColor,
  resolveConventions,
  setStoredConventions,
} from "../../conventions/store.js";
import {
  DEFAULT_CURRENCY_SYMBOL,
  DEFAULT_PRESET_FORMATS,
} from "../../conventions/defaults.js";
import { buildFormatString } from "../../conventions/format-builder.js";
import type {
  NumberPreset,
  StoredConventions,
  StoredCustomPreset,
  StoredFormatPreset,
} from "../../conventions/types.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { RULES_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { formatWorkbookLabel, getWorkbookContext } from "../../workbook/context.js";

type RulesTab = "user" | "workbook" | "conventions";

const BUILTIN_PRESET_NAMES: NumberPreset[] = [
  "number",
  "integer",
  "currency",
  "percent",
  "ratio",
  "text",
];

function setActiveTab(
  tabButtons: Record<RulesTab, HTMLButtonElement>,
  activeTab: RulesTab,
): void {
  const tabs: RulesTab[] = ["user", "workbook", "conventions"];

  for (const tab of tabs) {
    const button = tabButtons[tab];
    if (!button) continue;

    const isActive = tab === activeTab;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", String(isActive));
    button.setAttribute("tabindex", isActive ? "0" : "-1");
  }
}

function formatCounterLabel(chars: number, limit: number): string {
  return `${chars.toLocaleString()} / ${limit.toLocaleString()} chars`;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const node = document.createElement(tag);
  if (className) node.className = className;
  return node;
}

function cloneStoredConventions(value: StoredConventions): StoredConventions {
  return structuredClone(value);
}

function getPresetSection(draft: StoredConventions): Partial<Record<NumberPreset, StoredFormatPreset>> {
  if (!draft.presetFormats) {
    draft.presetFormats = {};
  }
  return draft.presetFormats;
}

function getCustomPresetSection(draft: StoredConventions): Record<string, StoredCustomPreset> {
  if (!draft.customPresets) {
    draft.customPresets = {};
  }
  return draft.customPresets;
}

function getVisualDefaults(draft: StoredConventions): NonNullable<StoredConventions["visualDefaults"]> {
  if (!draft.visualDefaults) {
    draft.visualDefaults = {};
  }
  return draft.visualDefaults;
}

function getColorConventions(draft: StoredConventions): NonNullable<StoredConventions["colorConventions"]> {
  if (!draft.colorConventions) {
    draft.colorConventions = {};
  }
  return draft.colorConventions;
}

function getHeaderStyle(draft: StoredConventions): NonNullable<StoredConventions["headerStyle"]> {
  if (!draft.headerStyle) {
    draft.headerStyle = {};
  }
  return draft.headerStyle;
}

function createColorSwatch(color: string): HTMLElement {
  const swatch = el("input", "pi-conventions-color-swatch");
  swatch.type = "color";
  swatch.disabled = true;
  swatch.tabIndex = -1;
  swatch.value = normalizeConventionColor(color) ?? "#000000";
  return swatch;
}

function createColorLegend(labelText: string, color: string): HTMLElement {
  const item = el("div", "pi-conventions-color-legend");

  const label = el("span", "pi-conventions-color-legend-label");
  label.textContent = labelText;

  const value = el("span", "pi-conventions-color-legend-value");
  value.textContent = normalizeConventionColor(color) ?? color;

  item.append(createColorSwatch(color), label, value);
  return item;
}

const SVG_NS = "http://www.w3.org/2000/svg";

function createHeaderPreview(
  fillColor: string,
  fontColor: string,
  bold: boolean,
  wrapText: boolean,
): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, "svg");
  svg.setAttribute("class", "pi-conventions-header-preview");
  svg.setAttribute("viewBox", "0 0 360 52");
  svg.setAttribute("role", "img");
  svg.setAttribute("aria-label", "Header style preview");

  const normalizedFill = normalizeConventionColor(fillColor) ?? "#4472C4";
  const normalizedFont = normalizeConventionColor(fontColor) ?? "#FFFFFF";

  const labels = ["Revenue", "Cost of Goods Sold", "Margin"];

  for (const [index, label] of labels.entries()) {
    const x = index * 120;

    const rect = document.createElementNS(SVG_NS, "rect");
    rect.setAttribute("x", String(x));
    rect.setAttribute("y", "0");
    rect.setAttribute("width", "120");
    rect.setAttribute("height", "52");
    rect.setAttribute("fill", normalizedFill);
    rect.setAttribute("stroke", "rgba(0,0,0,0.10)");

    const text = document.createElementNS(SVG_NS, "text");
    text.setAttribute("x", String(x + 8));
    text.setAttribute("fill", normalizedFont);
    text.setAttribute("font-size", "12");
    text.setAttribute("font-family", "DM Sans, sans-serif");
    text.setAttribute("font-weight", bold ? "700" : "400");

    if (wrapText && label === "Cost of Goods Sold") {
      text.setAttribute("y", "20");
      const line1 = document.createElementNS(SVG_NS, "tspan");
      line1.setAttribute("x", String(x + 8));
      line1.setAttribute("dy", "0");
      line1.textContent = "Cost of Goods";
      const line2 = document.createElementNS(SVG_NS, "tspan");
      line2.setAttribute("x", String(x + 8));
      line2.setAttribute("dy", "14");
      line2.textContent = "Sold";
      text.append(line1, line2);
    } else {
      text.setAttribute("y", "30");
      text.textContent = label === "Cost of Goods Sold" ? "Cost of Goods…" : label;
    }

    svg.append(rect, text);
  }

  return svg;
}

function normalizePresetName(name: string): string {
  return name.trim();
}

function createPresetPreview(presetName: string, preset: StoredFormatPreset): { positive: string; negative: string; zero: string } {
  const params = preset.builderParams;
  if (!params) {
    return {
      positive: "Custom",
      negative: "Custom",
      zero: "Custom",
    };
  }

  const dp = params.dp ?? 2;
  const formatter = new Intl.NumberFormat("en-US", {
    minimumFractionDigits: dp,
    maximumFractionDigits: dp,
    useGrouping: params.thousandsSeparator ?? true,
  });

  const symbol = params.currencySymbol ?? (presetName === "currency" ? DEFAULT_CURRENCY_SYMBOL : "");
  const suffix = presetName === "percent" ? "%" : presetName === "ratio" ? "x" : "";

  const positiveCore = `${symbol}${formatter.format(1234.5)}${suffix}`;
  const negativeCore = params.negativeStyle === "minus"
    ? `${symbol}-${formatter.format(1234.5)}${suffix}`
    : `${symbol}(${formatter.format(1234.5)}${suffix})`;

  let zeroCore = `${symbol}0${suffix}`;
  if (params.zeroStyle === "dash") {
    zeroCore = `${symbol}--${suffix}`;
  } else if (params.zeroStyle === "single-dash") {
    zeroCore = `${symbol}-${suffix}`;
  } else if (params.zeroStyle === "blank") {
    zeroCore = "(blank)";
  }

  return {
    positive: positiveCore,
    negative: negativeCore,
    zero: zeroCore,
  };
}

function createPreviewChip(label: string, value: string): HTMLElement {
  const chip = el("div", "pi-conventions-preview-chip");
  const title = el("div", "pi-conventions-preview-chip-label");
  title.textContent = label;
  const content = el("div", "pi-conventions-preview-chip-value");
  content.textContent = value;
  chip.append(title, content);
  return chip;
}

function createLabeledInput(args: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  className?: string;
  placeholder?: string;
}): HTMLElement {
  const row = el("div", "pi-conventions-field");
  const label = el("label", "pi-conventions-label");
  label.textContent = args.label;

  const input = el("input", args.className ?? "pi-conventions-input");
  input.type = "text";
  input.value = args.value;
  if (args.placeholder) {
    input.placeholder = args.placeholder;
  }

  input.addEventListener("change", () => {
    args.onChange(input.value);
  });

  row.append(label, input);
  return row;
}

function createLabeledNumberInput(args: {
  label: string;
  value: number;
  min?: number;
  max?: number;
  onChange: (value: number) => void;
}): HTMLElement {
  const row = el("div", "pi-conventions-field");
  const label = el("label", "pi-conventions-label");
  label.textContent = args.label;

  const input = el("input", "pi-conventions-input pi-conventions-input--narrow");
  input.type = "number";
  input.value = String(args.value);
  if (args.min !== undefined) input.min = String(args.min);
  if (args.max !== undefined) input.max = String(args.max);

  input.addEventListener("change", () => {
    const parsed = Number.parseInt(input.value, 10);
    if (!Number.isNaN(parsed)) {
      args.onChange(parsed);
    }
  });

  row.append(label, input);
  return row;
}

function createToggleButton(args: {
  label: string;
  value: boolean;
  onChange: (value: boolean) => void;
}): HTMLElement {
  const row = el("div", "pi-conventions-field");
  const label = el("label", "pi-conventions-label");
  label.textContent = args.label;

  const button = el("button", "pi-conventions-toggle");
  button.type = "button";
  button.setAttribute("role", "switch");

  const applyVisual = (value: boolean): void => {
    button.classList.toggle("is-on", value);
    button.setAttribute("aria-checked", String(value));
    button.textContent = value ? "On" : "Off";
  };

  let current = args.value;
  applyVisual(current);

  button.addEventListener("click", () => {
    current = !current;
    applyVisual(current);
    args.onChange(current);
  });

  row.append(label, button);
  return row;
}

function createQuickToggleSelect(args: {
  label: string;
  currentValue: string;
  options: Array<{ value: string; label: string }>;
  onChange: (value: string) => void;
}): HTMLElement {
  const group = el("div", "pi-conventions-quick-toggle");
  const label = el("label", "pi-conventions-quick-toggle-label");
  label.textContent = args.label;

  const select = el("select", "pi-conventions-select");
  for (const option of args.options) {
    const optionNode = document.createElement("option");
    optionNode.value = option.value;
    optionNode.textContent = option.label;
    optionNode.selected = option.value === args.currentValue;
    select.appendChild(optionNode);
  }

  select.addEventListener("change", () => {
    args.onChange(select.value);
  });

  group.append(label, select);
  return group;
}

function addCustomPreset(draft: StoredConventions): void {
  const customPresets = getCustomPresetSection(draft);

  const existingNames = new Set(Object.keys(customPresets));
  let index = 1;
  let candidate = `custom-${index}`;
  while (existingNames.has(candidate)) {
    index += 1;
    candidate = `custom-${index}`;
  }

  const defaultFormat = DEFAULT_PRESET_FORMATS.number.format;
  customPresets[candidate] = {
    format: defaultFormat,
    description: "",
  };
}

function renameCustomPreset(draft: StoredConventions, from: string, to: string): void {
  if (!draft.customPresets) {
    return;
  }

  const normalizedTo = normalizePresetName(to);
  if (normalizedTo.length === 0 || normalizedTo === from) {
    return;
  }

  if (Object.prototype.hasOwnProperty.call(draft.customPresets, normalizedTo)) {
    return;
  }

  const existing = draft.customPresets[from];
  if (!existing) {
    return;
  }

  delete draft.customPresets[from];
  draft.customPresets[normalizedTo] = existing;
}

function applyQuickPresetBuilder(presetName: NumberPreset, preset: StoredFormatPreset): void {
  const params = preset.builderParams;
  if (!params || presetName === "text") {
    return;
  }

  const built = buildFormatString(
    presetName,
    params.dp,
    params.currencySymbol,
    {
      negativeStyle: params.negativeStyle ?? "parens",
      zeroStyle: params.zeroStyle ?? "dash",
      thousandsSeparator: params.thousandsSeparator ?? true,
      accountingPadding: true,
    },
  );

  preset.format = built.format;
}

function renderFormatCard(args: {
  title: string;
  presetName: string;
  preset: StoredFormatPreset;
  onChange: () => void;
  onRemove?: () => void;
  onRename?: (value: string) => void;
  description?: string;
  onDescriptionChange?: (value: string) => void;
}): HTMLElement {
  const details = el("details", "pi-conventions-format-card");
  const summary = el("summary", "pi-conventions-format-card-summary");

  const summaryLeft = el("div", "pi-conventions-format-card-left");
  const titleNode = el("div", "pi-conventions-format-card-title");
  titleNode.textContent = args.title;
  const previewNode = el("div", "pi-conventions-format-card-preview");
  previewNode.textContent = args.preset.format;
  summaryLeft.append(titleNode, previewNode);

  summary.append(summaryLeft);
  details.append(summary);

  const body = el("div", "pi-conventions-format-card-body");

  if (args.onRename) {
    body.appendChild(createLabeledInput({
      label: "Name",
      value: args.presetName,
      onChange: (value) => {
        args.onRename?.(value);
        args.onChange();
      },
    }));
  }

  if (args.onDescriptionChange) {
    body.appendChild(createLabeledInput({
      label: "Description",
      value: args.description ?? "",
      onChange: (value) => {
        args.onDescriptionChange?.(value);
        args.onChange();
      },
      placeholder: "Optional",
    }));
  }

  const formatInput = createLabeledInput({
    label: "Format",
    value: args.preset.format,
    onChange: (value) => {
      args.preset.format = value;
      args.preset.builderParams = undefined;
      args.onChange();
    },
    className: "pi-conventions-input pi-conventions-input--wide pi-conventions-input--mono",
  });
  body.append(formatInput);

  const preview = createPresetPreview(args.presetName, args.preset);
  const previewRow = el("div", "pi-conventions-preview-row");
  previewRow.append(
    createPreviewChip("Positive", preview.positive),
    createPreviewChip("Negative", preview.negative),
    createPreviewChip("Zero", preview.zero),
  );
  body.append(previewRow);

  if (args.preset.builderParams && isBuiltinPresetName(args.presetName)) {
    const builtinPresetName = args.presetName;
    const quickRow = el("div", "pi-conventions-quick-toggles");
    const params = args.preset.builderParams;

    quickRow.appendChild(createQuickToggleSelect({
      label: "dp",
      currentValue: String(params.dp ?? 2),
      options: [0, 1, 2, 3, 4].map((dp) => ({ value: String(dp), label: String(dp) })),
      onChange: (value) => {
        const parsed = Number.parseInt(value, 10);
        if (!Number.isNaN(parsed)) {
          params.dp = parsed;
          applyQuickPresetBuilder(builtinPresetName, args.preset);
          args.onChange();
        }
      },
    }));

    quickRow.appendChild(createQuickToggleSelect({
      label: "neg",
      currentValue: params.negativeStyle ?? "parens",
      options: [
        { value: "parens", label: "(1,234)" },
        { value: "minus", label: "-1,234" },
      ],
      onChange: (value) => {
        if (value === "parens" || value === "minus") {
          params.negativeStyle = value;
          applyQuickPresetBuilder(builtinPresetName, args.preset);
          args.onChange();
        }
      },
    }));

    quickRow.appendChild(createQuickToggleSelect({
      label: "zero",
      currentValue: params.zeroStyle ?? "dash",
      options: [
        { value: "dash", label: "--" },
        { value: "single-dash", label: "-" },
        { value: "zero", label: "0" },
        { value: "blank", label: "blank" },
      ],
      onChange: (value) => {
        if (value === "dash" || value === "single-dash" || value === "zero" || value === "blank") {
          params.zeroStyle = value;
          applyQuickPresetBuilder(builtinPresetName, args.preset);
          args.onChange();
        }
      },
    }));

    if (builtinPresetName === "currency") {
      quickRow.appendChild(createLabeledInput({
        label: "symbol",
        value: params.currencySymbol ?? DEFAULT_CURRENCY_SYMBOL,
        onChange: (value) => {
          params.currencySymbol = value;
          applyQuickPresetBuilder("currency", args.preset);
          args.onChange();
        },
        className: "pi-conventions-input pi-conventions-input--narrow",
      }));
    }

    body.append(quickRow);
  } else if (isBuiltinPresetName(args.presetName)) {
    const builtinPresetName = args.presetName;
    const restore = el("button", "pi-conventions-link-btn");
    restore.type = "button";
    restore.textContent = "Custom format — use quick options to reset";
    restore.addEventListener("click", () => {
      const builtDefault = DEFAULT_PRESET_FORMATS[builtinPresetName];
      args.preset.builderParams = {
        ...(builtDefault.builderParams ?? {}),
      };
      applyQuickPresetBuilder(builtinPresetName, args.preset);
      args.onChange();
    });
    body.append(restore);
  }

  if (args.onRemove) {
    const removeButton = el("button", "pi-conventions-link-btn pi-conventions-link-btn--danger");
    removeButton.type = "button";
    removeButton.textContent = "Remove preset";
    removeButton.addEventListener("click", () => {
      args.onRemove?.();
      args.onChange();
    });
    body.append(removeButton);
  }

  details.append(body);
  return details;
}

function renderConventionsEditor(
  container: HTMLElement,
  draft: StoredConventions,
  requestRerender: () => void,
): void {
  container.replaceChildren();

  const resolved = resolveConventions(draft);

  const formatsSection = el("section", "pi-conventions-section");
  const formatsTitle = el("h3", "pi-conventions-section-title");
  formatsTitle.textContent = "Number formats";
  formatsSection.append(formatsTitle);

  const presetFormats = getPresetSection(draft);

  for (const presetName of BUILTIN_PRESET_NAMES) {
    const preset = presetFormats[presetName] ?? {
      format: resolved.presetFormats[presetName].format,
      builderParams: resolved.presetFormats[presetName].builderParams,
    };

    presetFormats[presetName] = preset;

    formatsSection.appendChild(renderFormatCard({
      title: presetName,
      presetName,
      preset,
      onChange: requestRerender,
    }));
  }

  const customPresets = getCustomPresetSection(draft);

  const customNames = Object.keys(customPresets).sort((left, right) => left.localeCompare(right));
  for (const customName of customNames) {
    const custom = customPresets[customName];
    if (!custom) continue;

    formatsSection.appendChild(renderFormatCard({
      title: customName,
      presetName: customName,
      preset: custom,
      description: custom.description,
      onRename: (nextName) => renameCustomPreset(draft, customName, nextName),
      onDescriptionChange: (value) => {
        custom.description = value;
      },
      onRemove: () => {
        delete customPresets[customName];
      },
      onChange: requestRerender,
    }));
  }

  const addCustomButton = el("button", "pi-overlay-btn pi-overlay-btn--ghost");
  addCustomButton.type = "button";
  addCustomButton.textContent = "Add custom format";
  addCustomButton.addEventListener("click", () => {
    addCustomPreset(draft);
    requestRerender();
  });
  formatsSection.append(addCustomButton);

  const colorsSection = el("section", "pi-conventions-section");
  const colorsTitle = el("h3", "pi-conventions-section-title");
  colorsTitle.textContent = "Colors (font color)";
  colorsSection.append(colorsTitle);

  const colorConventions = getColorConventions(draft);
  const hardcodedValueColor = colorConventions.hardcodedValueColor
    ?? resolved.colorConventions.hardcodedValueColor;
  const crossSheetColor = colorConventions.crossSheetLinkColor
    ?? resolved.colorConventions.crossSheetLinkColor;

  const colorField = (labelText: string, current: string, update: (value: string) => void): HTMLElement => {
    const row = el("div", "pi-conventions-field");
    const label = el("label", "pi-conventions-label");
    label.textContent = labelText;

    const right = el("div", "pi-conventions-color-field");
    right.append(createColorSwatch(current));

    const input = el("input", "pi-conventions-input");
    input.type = "text";
    input.value = current;
    input.placeholder = "#RRGGBB or rgb(r,g,b)";
    input.addEventListener("change", () => {
      const normalized = normalizeConventionColor(input.value);
      if (normalized) {
        update(normalized);
      }
      requestRerender();
    });

    right.append(input);
    row.append(label, right);
    return row;
  };

  colorsSection.append(
    colorField("Hardcoded values", hardcodedValueColor, (value) => {
      colorConventions.hardcodedValueColor = value;
    }),
    colorField("Cross-sheet links", crossSheetColor, (value) => {
      colorConventions.crossSheetLinkColor = value;
    }),
  );

  const headerSection = el("section", "pi-conventions-section");
  const headerTitle = el("h3", "pi-conventions-section-title");
  headerTitle.textContent = "Header style";
  headerSection.append(headerTitle);

  const headerStyle = getHeaderStyle(draft);
  const headerFill = headerStyle.fillColor ?? resolved.headerStyle.fillColor;
  const headerFont = headerStyle.fontColor ?? resolved.headerStyle.fontColor;
  const headerBold = headerStyle.bold ?? resolved.headerStyle.bold;
  const headerWrap = headerStyle.wrapText ?? resolved.headerStyle.wrapText;

  const headerColors = el("div", "pi-conventions-header-colors");
  headerColors.append(
    createColorLegend("Fill", headerFill),
    createColorLegend("Font", headerFont),
  );
  headerSection.append(headerColors);

  const previewRow = createHeaderPreview(headerFill, headerFont, headerBold, headerWrap);
  headerSection.append(previewRow);

  headerSection.append(
    colorField("Fill color", headerFill, (value) => {
      headerStyle.fillColor = value;
    }),
    colorField("Font color", headerFont, (value) => {
      headerStyle.fontColor = value;
    }),
    createToggleButton({
      label: "Bold",
      value: headerBold,
      onChange: (value) => {
        headerStyle.bold = value;
        requestRerender();
      },
    }),
    createToggleButton({
      label: "Wrap text",
      value: headerWrap,
      onChange: (value) => {
        headerStyle.wrapText = value;
        requestRerender();
      },
    }),
  );

  const visualSection = el("section", "pi-conventions-section");
  const visualTitle = el("h3", "pi-conventions-section-title");
  visualTitle.textContent = "Default font";
  visualSection.append(visualTitle);

  const visualDefaults = getVisualDefaults(draft);
  const fontName = visualDefaults.fontName ?? resolved.visualDefaults.fontName;
  const fontSize = visualDefaults.fontSize ?? resolved.visualDefaults.fontSize;

  visualSection.append(
    createLabeledInput({
      label: "Font name",
      value: fontName,
      onChange: (value) => {
        visualDefaults.fontName = value;
        requestRerender();
      },
    }),
    createLabeledNumberInput({
      label: "Font size",
      value: fontSize,
      min: 6,
      max: 72,
      onChange: (value) => {
        visualDefaults.fontSize = value;
        requestRerender();
      },
    }),
  );

  container.append(formatsSection, colorsSection, headerSection, visualSection);
}

export async function showRulesDialog(opts?: {
  onSaved?: () => void | Promise<void>;
}): Promise<void> {
  if (closeOverlayById(RULES_OVERLAY_ID)) {
    return;
  }

  const storage = getAppStorage();
  const workbookContext = await getWorkbookContext();
  const workbookId = workbookContext.workbookId;
  const workbookLabel = formatWorkbookLabel(workbookContext);

  let userDraft = (await getUserRules(storage.settings)) ?? "";
  let workbookDraft = (await getWorkbookRules(storage.settings, workbookId)) ?? "";
  const storedConventions = await getStoredConventions(storage.settings);
  const conventionsDraft = cloneStoredConventions(storedConventions);
  let activeTab: RulesTab = "user";

  const dialog = createOverlayDialog({
    overlayId: RULES_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close rules",
    title: "Rules",
    subtitle: "Set guidance for all files, this workbook, and formatting conventions.",
  });

  const tabs = document.createElement("div");
  tabs.className = "pi-overlay-tabs";
  tabs.setAttribute("role", "tablist");

  const userTab = document.createElement("button");
  userTab.type = "button";
  userTab.textContent = "All my files";
  userTab.className = "pi-overlay-tab";
  userTab.setAttribute("role", "tab");

  const workbookTab = document.createElement("button");
  workbookTab.type = "button";
  workbookTab.textContent = "This file";
  workbookTab.className = "pi-overlay-tab";
  workbookTab.setAttribute("role", "tab");

  const conventionsTab = document.createElement("button");
  conventionsTab.type = "button";
  conventionsTab.textContent = "Formats";
  conventionsTab.className = "pi-overlay-tab";
  conventionsTab.setAttribute("role", "tab");

  tabs.append(userTab, workbookTab, conventionsTab);

  const workbookTag = document.createElement("div");
  workbookTag.className = "pi-overlay-workbook-tag";
  workbookTag.textContent = `Workbook: ${workbookLabel}`;

  const hint = document.createElement("div");
  hint.className = "pi-overlay-hint";

  const textarea = document.createElement("textarea");
  textarea.className = "pi-overlay-textarea";

  const conventionsContainer = document.createElement("div");
  conventionsContainer.className = "pi-conventions-container";

  const body = document.createElement("div");
  body.className = "pi-overlay-body";
  body.append(header, tabs, workbookTag, hint, textarea, conventionsContainer);

  const footer = document.createElement("div");
  footer.className = "pi-overlay-footer";

  const counter = document.createElement("div");
  counter.className = "pi-overlay-counter";

  const actions = document.createElement("div");
  actions.className = "pi-overlay-actions";

  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.textContent = "Cancel";
  cancelBtn.className = "pi-overlay-btn pi-overlay-btn--ghost";

  const saveBtn = document.createElement("button");
  saveBtn.type = "button";
  saveBtn.textContent = "Save";
  saveBtn.className = "pi-overlay-btn pi-overlay-btn--primary";

  actions.append(cancelBtn, saveBtn);
  footer.append(counter, actions);
  dialog.card.append(body, footer);

  const tabButtons: Record<RulesTab, HTMLButtonElement> = {
    user: userTab,
    workbook: workbookTab,
    conventions: conventionsTab,
  };

  const rerenderConventions = (): void => {
    renderConventionsEditor(conventionsContainer, conventionsDraft, rerenderConventions);
  };

  const refreshTabUi = (): void => {
    setActiveTab(tabButtons, activeTab);

    const isConventionsTab = activeTab === "conventions";
    textarea.hidden = isConventionsTab;
    conventionsContainer.hidden = !isConventionsTab;
    counter.hidden = isConventionsTab;

    if (activeTab === "user") {
      textarea.value = userDraft;
      textarea.placeholder =
        "Your preferences and habits, e.g.\n• Always use EUR for currencies\n• Format dates as dd-mmm-yyyy\n• Check circular references after writes";

      const count = userDraft.length;
      counter.textContent = formatCounterLabel(count, USER_RULES_SOFT_LIMIT);
      counter.classList.toggle("is-warning", count > USER_RULES_SOFT_LIMIT);

      hint.textContent =
        "Guidance given to Pi in all your conversations. Pi can also update these when you tell it your preferences — e.g. \"always use EUR\".";
      workbookTag.hidden = true;
      return;
    }

    if (activeTab === "workbook") {
      textarea.value = workbookDraft;
      textarea.placeholder =
        "Notes about this workbook's structure, e.g.\n• DCF model for Acme Corp, FY2025\n• Revenue assumptions in Inputs!B5:B15\n• Don't modify the Summary sheet";

      const count = workbookDraft.length;
      counter.textContent = formatCounterLabel(count, WORKBOOK_RULES_SOFT_LIMIT);
      counter.classList.toggle("is-warning", count > WORKBOOK_RULES_SOFT_LIMIT);

      hint.textContent = !workbookId
        ? "Can't identify this workbook right now — try saving the file first."
        : "Guidance given to Pi only when it reads this file.";

      workbookTag.hidden = false;
      return;
    }

    workbookTag.hidden = true;
    hint.textContent = "Set preset formats, font colors, header style, and default font.";
    rerenderConventions();
  };

  const saveActiveDraft = (): void => {
    if (activeTab === "user") {
      userDraft = textarea.value;
      return;
    }

    if (activeTab === "workbook") {
      workbookDraft = textarea.value;
    }
  };

  userTab.addEventListener("click", () => {
    saveActiveDraft();
    activeTab = "user";
    refreshTabUi();
  });

  workbookTab.addEventListener("click", () => {
    saveActiveDraft();
    activeTab = "workbook";
    refreshTabUi();
  });

  conventionsTab.addEventListener("click", () => {
    saveActiveDraft();
    activeTab = "conventions";
    refreshTabUi();
  });

  textarea.addEventListener("input", () => {
    saveActiveDraft();
    refreshTabUi();
  });

  cancelBtn.addEventListener("click", () => {
    closeOverlay();
  });

  saveBtn.addEventListener("click", () => {
    void (async () => {
      saveActiveDraft();

      await setUserRules(storage.settings, userDraft);
      if (workbookId) {
        await setWorkbookRules(storage.settings, workbookId, workbookDraft);
      }

      await setStoredConventions(storage.settings, conventionsDraft);

      document.dispatchEvent(new CustomEvent("pi:rules-updated"));
      document.dispatchEvent(new CustomEvent("pi:conventions-updated"));
      document.dispatchEvent(new CustomEvent("pi:status-update"));

      if (opts?.onSaved) {
        await opts.onSaved();
      }

      showToast("Rules saved");
      closeOverlay();
    })();
  });

  refreshTabUi();
  dialog.mount();
  textarea.focus();
}
