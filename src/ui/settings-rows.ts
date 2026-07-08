/**
 * Row primitives for the settings shell.
 *
 * A settings page is composed of groups (labelled cards) containing rows:
 * navigation rows (label + value preview + chevron), toggle rows, and
 * select rows. All classes live in `theme/overlays/settings-shell.css`.
 */

import { setSafeInnerHTML } from "../utils/html.js";
import { createToggle } from "./extensions-hub-components.js";

const NAV_CHEVRON_SVG = `<svg class="pi-set-row__chevron" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" aria-hidden="true"><path d="M6 4l4 4-4 4"/></svg>`;

export interface SettingsGroupResult {
  root: HTMLElement;
  list: HTMLDivElement;
}

/**
 * A labelled group card. Rows appended to `list` are separated by hairlines.
 */
export function createSettingsGroup(label?: string): SettingsGroupResult {
  const root = document.createElement("section");
  root.className = "pi-set-group";

  if (label !== undefined && label.length > 0) {
    const heading = document.createElement("h3");
    heading.className = "pi-set-group__label";
    heading.textContent = label;
    root.appendChild(heading);
  }

  const list = document.createElement("div");
  list.className = "pi-set-group__list";
  root.appendChild(list);

  return { root, list };
}

export interface NavRowOptions {
  icon?: Element;
  label: string;
  sublabel?: string;
  value?: string;
  onActivate: () => void;
}

export interface NavRowResult {
  root: HTMLButtonElement;
  setValue: (value: string) => void;
}

/**
 * Full-width navigation row: optional icon, label (+sublabel), right-aligned
 * value preview and chevron. Activating navigates to a nested page.
 */
export function createNavRow(opts: NavRowOptions): NavRowResult {
  const row = document.createElement("button");
  row.type = "button";
  row.className = "pi-set-row pi-set-row--nav";

  if (opts.icon) {
    const iconEl = document.createElement("span");
    iconEl.className = "pi-set-row__icon";
    iconEl.setAttribute("aria-hidden", "true");
    iconEl.appendChild(opts.icon);
    row.appendChild(iconEl);
  }

  const text = document.createElement("span");
  text.className = "pi-set-row__text";

  const labelEl = document.createElement("span");
  labelEl.className = "pi-set-row__label";
  labelEl.textContent = opts.label;
  text.appendChild(labelEl);

  if (opts.sublabel !== undefined && opts.sublabel.length > 0) {
    const sub = document.createElement("span");
    sub.className = "pi-set-row__sublabel";
    sub.textContent = opts.sublabel;
    text.appendChild(sub);
  }

  row.appendChild(text);

  const valueEl = document.createElement("span");
  valueEl.className = "pi-set-row__value";
  if (opts.value !== undefined) {
    valueEl.textContent = opts.value;
  }
  row.appendChild(valueEl);

  const chevronHost = document.createElement("span");
  chevronHost.className = "pi-set-row__chevron-host";
  setSafeInnerHTML(chevronHost, NAV_CHEVRON_SVG, "trusted static chevron SVG constant");
  row.appendChild(chevronHost);

  row.addEventListener("click", opts.onActivate);

  return {
    root: row,
    setValue: (value: string) => {
      valueEl.textContent = value;
    },
  };
}

export interface SettingToggleRowOptions {
  label: string;
  sublabel?: string;
  checked?: boolean;
  onChange?: (checked: boolean) => void;
}

export interface SettingToggleRowResult {
  root: HTMLDivElement;
  input: HTMLInputElement;
}

/**
 * Inline toggle row styled to sit inside a settings group.
 */
export function createSettingToggleRow(opts: SettingToggleRowOptions): SettingToggleRowResult {
  const row = document.createElement("div");
  row.className = "pi-set-row pi-set-row--control";

  const text = document.createElement("span");
  text.className = "pi-set-row__text";

  const labelEl = document.createElement("span");
  labelEl.className = "pi-set-row__label";
  labelEl.textContent = opts.label;
  text.appendChild(labelEl);

  if (opts.sublabel !== undefined && opts.sublabel.length > 0) {
    const sub = document.createElement("span");
    sub.className = "pi-set-row__sublabel";
    sub.textContent = opts.sublabel;
    text.appendChild(sub);
  }

  row.appendChild(text);

  const toggle = createToggle({
    ...(opts.checked !== undefined ? { checked: opts.checked } : {}),
    ...(opts.onChange !== undefined ? { onChange: opts.onChange } : {}),
  });
  row.appendChild(toggle.root);

  return { root: row, input: toggle.input };
}

export interface SettingSelectRowOptions {
  label: string;
  sublabel?: string;
  options: Array<{ value: string; label: string }>;
  value: string;
  onChange: (value: string) => void;
}

export interface SettingSelectRowResult {
  root: HTMLDivElement;
  select: HTMLSelectElement;
}

/**
 * Inline select row styled to sit inside a settings group.
 */
export function createSettingSelectRow(opts: SettingSelectRowOptions): SettingSelectRowResult {
  const row = document.createElement("div");
  row.className = "pi-set-row pi-set-row--control";

  const text = document.createElement("span");
  text.className = "pi-set-row__text";

  const labelEl = document.createElement("span");
  labelEl.className = "pi-set-row__label";
  labelEl.textContent = opts.label;
  text.appendChild(labelEl);

  if (opts.sublabel !== undefined && opts.sublabel.length > 0) {
    const sub = document.createElement("span");
    sub.className = "pi-set-row__sublabel";
    sub.textContent = opts.sublabel;
    text.appendChild(sub);
  }

  row.appendChild(text);

  const select = document.createElement("select");
  select.className = "pi-set-row__select";
  for (const option of opts.options) {
    const node = document.createElement("option");
    node.value = option.value;
    node.textContent = option.label;
    node.selected = option.value === opts.value;
    select.appendChild(node);
  }

  select.addEventListener("change", () => {
    opts.onChange(select.value);
  });

  row.appendChild(select);
  return { root: row, select };
}

export interface SegmentedControlOptions {
  segments: Array<{ id: string; label: string }>;
  active: string;
  ariaLabel: string;
  onChange: (id: string) => void;
}

export interface SegmentedControlResult {
  root: HTMLDivElement;
  setActive: (id: string) => void;
}

/**
 * Shared segmented control for peer views inside one page (e.g. the Rules
 * editor's three edit surfaces). Replaces the per-overlay tab systems.
 */
export function createSegmentedControl(opts: SegmentedControlOptions): SegmentedControlResult {
  const root = document.createElement("div");
  root.className = "pi-set-segmented";
  root.setAttribute("role", "tablist");
  root.setAttribute("aria-label", opts.ariaLabel);

  const buttons = new Map<string, HTMLButtonElement>();

  const setActive = (id: string): void => {
    for (const [segmentId, button] of buttons) {
      const isActive = segmentId === id;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
      button.setAttribute("tabindex", isActive ? "0" : "-1");
    }
  };

  for (const segment of opts.segments) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-set-segmented__btn";
    button.textContent = segment.label;
    button.setAttribute("role", "tab");
    button.addEventListener("click", () => {
      setActive(segment.id);
      opts.onChange(segment.id);
    });
    buttons.set(segment.id, button);
    root.appendChild(button);
  }

  setActive(opts.active);
  return { root, setActive };
}
