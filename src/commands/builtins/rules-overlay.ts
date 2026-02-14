/**
 * Rules editor overlay — rules (user + workbook) and number format conventions.
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
  setStoredConventions,
  resolveConventions,
  mergeStoredConventions,
} from "../../conventions/store.js";
import { DEFAULT_CURRENCY_SYMBOL, PRESET_DEFAULT_DP } from "../../conventions/defaults.js";
import type { StoredConventions, NumberPreset } from "../../conventions/types.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { RULES_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { formatWorkbookLabel, getWorkbookContext } from "../../workbook/context.js";

type RulesTab = "user" | "workbook" | "conventions";

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

// ── Conventions form builder ─────────────────────────────────────────

interface ConventionsFormState {
  currencySymbol: string;
  negativeStyle: "parens" | "minus";
  zeroStyle: "dash" | "zero" | "blank";
  thousandsSeparator: boolean;
  accountingPadding: boolean;
  numberDp: number;
  currencyDp: number;
  percentDp: number;
  ratioDp: number;
}

function updateConventionsFormField<K extends keyof ConventionsFormState>(
  state: ConventionsFormState,
  key: K,
  value: ConventionsFormState[K],
): void {
  state[key] = value;
}

function resolvedToFormState(stored: StoredConventions): ConventionsFormState {
  const resolved = resolveConventions(stored);
  return {
    currencySymbol: resolved.currencySymbol,
    negativeStyle: resolved.conventions.negativeStyle,
    zeroStyle: resolved.conventions.zeroStyle,
    thousandsSeparator: resolved.conventions.thousandsSeparator,
    accountingPadding: resolved.conventions.accountingPadding,
    numberDp: resolved.presetDp.number ?? PRESET_DEFAULT_DP.number ?? 2,
    currencyDp: resolved.presetDp.currency ?? PRESET_DEFAULT_DP.currency ?? 2,
    percentDp: resolved.presetDp.percent ?? PRESET_DEFAULT_DP.percent ?? 1,
    ratioDp: resolved.presetDp.ratio ?? PRESET_DEFAULT_DP.ratio ?? 1,
  };
}

function formStateToStored(form: ConventionsFormState): StoredConventions {
  const presetDp: Partial<Record<NumberPreset, number>> = {
    number: form.numberDp,
    currency: form.currencyDp,
    percent: form.percentDp,
    ratio: form.ratioDp,
  };

  const stored: StoredConventions = {
    currencySymbol: form.currencySymbol,
    negativeStyle: form.negativeStyle,
    zeroStyle: form.zeroStyle,
    thousandsSeparator: form.thousandsSeparator,
    accountingPadding: form.accountingPadding,
    presetDp,
  };

  return stored;
}

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (className) e.className = className;
  return e;
}

function createSelectField(
  label: string,
  options: Array<{ value: string; label: string }>,
  currentValue: string,
  onChange: (value: string) => void,
): HTMLElement {
  const row = el("div", "pi-conventions-field");

  const labelEl = el("label", "pi-conventions-label");
  labelEl.textContent = label;

  const select = el("select", "pi-conventions-select");
  for (const opt of options) {
    const optEl = document.createElement("option");
    optEl.value = opt.value;
    optEl.textContent = opt.label;
    if (opt.value === currentValue) optEl.selected = true;
    select.appendChild(optEl);
  }
  select.addEventListener("change", () => { onChange(select.value); });

  row.append(labelEl, select);
  return row;
}

function createTextInputField(
  label: string,
  currentValue: string,
  opts: { placeholder?: string; narrow?: boolean },
  onChange: (value: string) => void,
): HTMLElement {
  const row = el("div", "pi-conventions-field");

  const labelEl = el("label", "pi-conventions-label");
  labelEl.textContent = label;

  const cls = opts.narrow ? "pi-conventions-input pi-conventions-input--narrow" : "pi-conventions-input";
  const input = el("input", cls);
  input.type = "text";
  input.value = currentValue;
  if (opts.placeholder) input.placeholder = opts.placeholder;
  input.addEventListener("input", () => { onChange(input.value); });

  row.append(labelEl, input);
  return row;
}

function createToggleField(
  label: string,
  currentValue: boolean,
  onChange: (value: boolean) => void,
): HTMLElement {
  const row = el("div", "pi-conventions-field");

  const labelEl = el("label", "pi-conventions-label");
  labelEl.textContent = label;

  const toggle = el("button", "pi-conventions-toggle");
  toggle.type = "button";
  toggle.setAttribute("role", "switch");

  const updateVisual = (val: boolean) => {
    toggle.classList.toggle("is-on", val);
    toggle.setAttribute("aria-checked", String(val));
    toggle.textContent = val ? "On" : "Off";
  };
  updateVisual(currentValue);

  let value = currentValue;
  toggle.addEventListener("click", () => {
    value = !value;
    updateVisual(value);
    onChange(value);
  });

  row.append(labelEl, toggle);
  return row;
}

function createNumberField(
  label: string,
  currentValue: number,
  opts: { min?: number; max?: number },
  onChange: (value: number) => void,
): HTMLElement {
  const row = el("div", "pi-conventions-field");

  const labelEl = el("label", "pi-conventions-label");
  labelEl.textContent = label;

  const input = el("input", "pi-conventions-input pi-conventions-input--narrow");
  input.type = "number";
  input.value = String(currentValue);
  if (opts.min !== undefined) input.min = String(opts.min);
  if (opts.max !== undefined) input.max = String(opts.max);

  input.addEventListener("input", () => {
    const n = parseInt(input.value, 10);
    if (!Number.isNaN(n)) onChange(n);
  });

  row.append(labelEl, input);
  return row;
}

type ConventionsFormUpdater = <K extends keyof ConventionsFormState>(key: K, value: ConventionsFormState[K]) => void;

function buildConventionsForm(
  state: ConventionsFormState,
  onUpdate: ConventionsFormUpdater,
): HTMLElement {
  const form = el("div", "pi-conventions-form");

  const generalSection = el("div", "pi-conventions-section");
  const generalTitle = el("div", "pi-conventions-section-title");
  generalTitle.textContent = "General";
  generalSection.appendChild(generalTitle);

  generalSection.appendChild(
    createTextInputField("Currency symbol", state.currencySymbol, { placeholder: DEFAULT_CURRENCY_SYMBOL, narrow: true },
      (v) => { onUpdate("currencySymbol", v || DEFAULT_CURRENCY_SYMBOL); }),
  );
  generalSection.appendChild(
    createSelectField("Negatives", [
      { value: "parens", label: "(1,234) — parentheses" },
      { value: "minus", label: "-1,234 — minus sign" },
    ], state.negativeStyle, (v) => { onUpdate("negativeStyle", v as "parens" | "minus"); }),
  );
  generalSection.appendChild(
    createSelectField("Zeros", [
      { value: "dash", label: '-- — dash' },
      { value: "zero", label: "0 — literal zero" },
      { value: "blank", label: "(blank)" },
    ], state.zeroStyle, (v) => { onUpdate("zeroStyle", v as "dash" | "zero" | "blank"); }),
  );
  generalSection.appendChild(
    createToggleField("Thousands separator", state.thousandsSeparator,
      (v) => { onUpdate("thousandsSeparator", v); }),
  );
  generalSection.appendChild(
    createToggleField("Accounting padding", state.accountingPadding,
      (v) => { onUpdate("accountingPadding", v); }),
  );

  const dpSection = el("div", "pi-conventions-section");
  const dpTitle = el("div", "pi-conventions-section-title");
  dpTitle.textContent = "Decimal places";
  dpSection.appendChild(dpTitle);

  dpSection.appendChild(
    createNumberField("Number", state.numberDp, { min: 0, max: 10 },
      (v) => { onUpdate("numberDp", v); }),
  );
  dpSection.appendChild(
    createNumberField("Currency", state.currencyDp, { min: 0, max: 10 },
      (v) => { onUpdate("currencyDp", v); }),
  );
  dpSection.appendChild(
    createNumberField("Percent", state.percentDp, { min: 0, max: 10 },
      (v) => { onUpdate("percentDp", v); }),
  );
  dpSection.appendChild(
    createNumberField("Ratio", state.ratioDp, { min: 0, max: 10 },
      (v) => { onUpdate("ratioDp", v); }),
  );

  form.append(generalSection, dpSection);
  return form;
}

// ── Main overlay ─────────────────────────────────────────────────────

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
  const conventionsFormState = resolvedToFormState(storedConventions);
  let activeTab: RulesTab = "user";

  const dialog = createOverlayDialog({
    overlayId: RULES_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close rules",
    title: "Rules",
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
  conventionsTab.textContent = "Number format";
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

  let conventionsFormEl: HTMLElement | null = null;

  const refreshTabUi = () => {
    setActiveTab(tabButtons, activeTab);

    const isConventionsTab = activeTab === "conventions";

    // Toggle visibility of textarea vs conventions form
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

      if (!workbookId) {
        hint.textContent =
          "Can't identify this workbook right now — try saving the file first.";
      } else {
        hint.textContent =
          "Guidance given to Pi only when it reads this file.";
      }

      workbookTag.hidden = false;
      return;
    }

    // Conventions tab
    workbookTag.hidden = true;
    hint.textContent = "Default number formatting conventions. Pi uses these when applying styles.";

    // Build the form if not yet created
    if (!conventionsFormEl) {
      conventionsFormEl = buildConventionsForm(
        conventionsFormState,
        (key, value) => {
          updateConventionsFormField(conventionsFormState, key, value);
        },
      );
      conventionsContainer.appendChild(conventionsFormEl);
    }
  };

  const saveActiveDraft = () => {
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

      // Save rules
      await setUserRules(storage.settings, userDraft);
      if (workbookId) {
        await setWorkbookRules(storage.settings, workbookId, workbookDraft);
      }

      // Save conventions
      const currentStored = await getStoredConventions(storage.settings);
      const updates = formStateToStored(conventionsFormState);
      const merged = mergeStoredConventions(currentStored, updates);
      await setStoredConventions(storage.settings, merged);

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
