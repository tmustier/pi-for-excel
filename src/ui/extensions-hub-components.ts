/**
 * DOM helpers for the Extensions hub overlay components.
 *
 * Each function creates a small DOM subtree matching the CSS classes
 * in `extensions-hub.css`. They are intentionally thin — the overlay
 * tab builders compose them and wire event listeners.
 */

// ── Chevron SVG (shared across item cards) ──────────

const CHEVRON_SVG = `<svg class="pi-item-card__chevron" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M6 4l4 4-4 4"/></svg>`;

// ── Toggle switch ───────────────────────────────────

export interface ToggleOptions {
  checked?: boolean;
  onChange?: (checked: boolean) => void;
  stopPropagation?: boolean;
}

/**
 * Creates an iOS-style toggle switch.
 * Uses the shared `.pi-toggle` CSS from toggle.css.
 */
export function createToggle(opts: ToggleOptions = {}): {
  root: HTMLLabelElement;
  input: HTMLInputElement;
} {
  const label = document.createElement("label");
  label.className = "pi-toggle";

  const input = document.createElement("input");
  input.type = "checkbox";
  input.className = "pi-toggle__input";
  if (opts.checked) input.checked = true;

  const track = document.createElement("span");
  track.className = "pi-toggle__track";

  const thumb = document.createElement("span");
  thumb.className = "pi-toggle__thumb";

  label.append(input, track, thumb);

  if (opts.stopPropagation) {
    label.addEventListener("click", (e) => e.stopPropagation());
  }

  if (opts.onChange) {
    const handler = opts.onChange;
    input.addEventListener("change", () => handler(input.checked));
  }

  return { root: label, input };
}

// ── Toggle row ──────────────────────────────────────

export interface ToggleRowOptions {
  label: string;
  sublabel?: string;
  checked?: boolean;
  onChange?: (checked: boolean) => void;
}

/**
 * Creates a label + toggle row.
 */
export function createToggleRow(opts: ToggleRowOptions): {
  root: HTMLDivElement;
  input: HTMLInputElement;
} {
  const row = document.createElement("div");
  row.className = "pi-toggle-row";

  const labels = document.createElement("div");
  labels.className = "pi-toggle-row__labels";

  const labelEl = document.createElement("div");
  labelEl.className = "pi-toggle-row__label";
  labelEl.textContent = opts.label;
  labels.appendChild(labelEl);

  if (opts.sublabel) {
    const sub = document.createElement("div");
    sub.className = "pi-toggle-row__sublabel";
    sub.textContent = opts.sublabel;
    labels.appendChild(sub);
  }

  const toggle = createToggle({
    checked: opts.checked,
    onChange: opts.onChange,
  });

  row.append(labels, toggle.root);
  return { root: row, input: toggle.input };
}

// ── Section header ──────────────────────────────────

export interface SectionHeaderOptions {
  label: string;
  count?: number;
  actionLabel?: string;
  onAction?: () => void;
}

export function createSectionHeader(opts: SectionHeaderOptions): HTMLDivElement {
  const header = document.createElement("div");
  header.className = "pi-section-header";

  const label = document.createElement("span");
  label.className = "pi-section-header__label";
  label.textContent = opts.label;
  header.appendChild(label);

  if (opts.count != null) {
    const count = document.createElement("span");
    count.className = "pi-section-header__count";
    count.textContent = `${opts.count} ${opts.count === 1 ? "skill" : "skills"}`;
    header.appendChild(count);
  }

  if (opts.actionLabel && opts.onAction) {
    const action = document.createElement("button");
    action.type = "button";
    action.className = "pi-section-header__action";
    action.textContent = opts.actionLabel;
    action.addEventListener("click", opts.onAction);
    header.appendChild(action);
  }

  return header;
}

// ── Item card ───────────────────────────────────────

export type ItemCardIconColor = "green" | "blue" | "purple" | "amber";

export interface ItemCardOptions {
  icon: string;
  iconColor?: ItemCardIconColor;
  name: string;
  description?: string;
  meta?: string;
  expandable?: boolean;
  expanded?: boolean;
  badges?: Array<{ text: string; tone: "ok" | "warn" | "muted" | "info" }>;
  /** Content for the right side of the header (before chevron). */
  rightContent?: HTMLElement;
}

export interface ItemCardResult {
  root: HTMLDivElement;
  header: HTMLDivElement;
  body: HTMLDivElement;
  setExpanded: (expanded: boolean) => void;
}

/**
 * Creates an expandable item card.
 */
export function createItemCard(opts: ItemCardOptions): ItemCardResult {
  const card = document.createElement("div");
  card.className = "pi-item-card";
  if (opts.expanded) card.setAttribute("data-expanded", "");

  // Header
  const header = document.createElement("div");
  header.className = `pi-item-card__header${opts.expandable ? " pi-item-card__header--expandable" : ""}`;

  const iconEl = document.createElement("div");
  iconEl.className = `pi-item-card__icon${opts.iconColor ? ` pi-item-card__icon--${opts.iconColor}` : ""}`;
  iconEl.textContent = opts.icon;

  const text = document.createElement("div");
  text.className = "pi-item-card__text";

  const nameEl = document.createElement("div");
  nameEl.className = "pi-item-card__name";
  nameEl.textContent = opts.name;
  text.appendChild(nameEl);

  if (opts.description) {
    const desc = document.createElement("div");
    desc.className = "pi-item-card__desc";
    desc.textContent = opts.description;
    text.appendChild(desc);
  }

  if (opts.meta) {
    const meta = document.createElement("div");
    meta.className = "pi-item-card__meta";
    meta.textContent = opts.meta;
    text.appendChild(meta);
  }

  header.append(iconEl, text);

  // Right side
  const right = document.createElement("div");
  right.className = "pi-item-card__right";

  if (opts.rightContent) {
    right.appendChild(opts.rightContent);
  }

  if (opts.badges) {
    for (const badge of opts.badges) {
      const badgeEl = document.createElement("span");
      badgeEl.className = `pi-overlay-badge pi-overlay-badge--${badge.tone}`;
      badgeEl.textContent = badge.text;
      right.appendChild(badgeEl);
    }
  }

  if (opts.expandable) {
    const chevron = document.createElement("span");
    chevron.innerHTML = CHEVRON_SVG;
    // The SVG is the first child, extract it
    const svg = chevron.firstElementChild;
    if (svg) right.appendChild(svg);
  }

  header.appendChild(right);

  // Body
  const body = document.createElement("div");
  body.className = "pi-item-card__body";

  // Expand/collapse
  if (opts.expandable) {
    header.addEventListener("click", () => {
      card.toggleAttribute("data-expanded");
    });
  }

  card.append(header, body);

  return {
    root: card,
    header,
    body,
    setExpanded: (expanded: boolean) => {
      if (expanded) {
        card.setAttribute("data-expanded", "");
      } else {
        card.removeAttribute("data-expanded");
      }
    },
  };
}

// ── Config row (inside item card body) ──────────────

export function createConfigRow(label: string, content: HTMLElement): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-item-card__config-row";

  const labelEl = document.createElement("span");
  labelEl.className = "pi-item-card__config-label";
  labelEl.textContent = label;

  row.append(labelEl, content);
  return row;
}

export function createConfigValue(text: string): HTMLSpanElement {
  const el = document.createElement("span");
  el.className = "pi-item-card__config-value";
  el.textContent = text;
  el.title = text;
  return el;
}

export function createConfigInput(opts: {
  value?: string;
  placeholder?: string;
  type?: string;
  onChange?: (value: string) => void;
}): HTMLInputElement {
  const input = document.createElement("input");
  input.className = "pi-item-card__config-input";
  input.type = opts.type ?? "text";
  if (opts.value) input.value = opts.value;
  if (opts.placeholder) input.placeholder = opts.placeholder;
  if (opts.onChange) {
    const handler = opts.onChange;
    input.addEventListener("change", () => handler(input.value));
  }
  return input;
}

// ── Callout ─────────────────────────────────────────

export function createCallout(
  tone: "info" | "warn" | "success",
  icon: string,
  message: string,
  opts?: { compact?: boolean },
): HTMLDivElement {
  const callout = document.createElement("div");
  callout.className = `pi-callout pi-callout--${tone}${opts?.compact ? " pi-callout--compact" : ""}`;

  const iconEl = document.createElement("span");
  iconEl.className = "pi-callout__icon";
  iconEl.textContent = icon;

  const body = document.createElement("div");
  body.className = "pi-callout__body";
  body.textContent = message;

  callout.append(iconEl, body);
  return callout;
}

// ── Add form ────────────────────────────────────────

export function createAddForm(): HTMLDivElement {
  const form = document.createElement("div");
  form.className = "pi-add-form";
  return form;
}

export function createAddFormRow(): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-add-form__row";
  return row;
}

export function createAddFormInput(placeholder: string): HTMLInputElement {
  const input = document.createElement("input");
  input.className = "pi-add-form__input";
  input.placeholder = placeholder;
  return input;
}

// ── Empty inline ────────────────────────────────────

export function createEmptyInline(icon: string, text: string): HTMLDivElement {
  const el = document.createElement("div");
  el.className = "pi-empty-inline";

  const iconEl = document.createElement("span");
  iconEl.className = "pi-empty-inline__icon";
  iconEl.textContent = icon;

  const textEl = document.createElement("span");
  textEl.className = "pi-empty-inline__text";
  textEl.textContent = text;

  el.append(iconEl, textEl);
  return el;
}

// ── Action buttons row ──────────────────────────────

export function createActionsRow(...buttons: HTMLElement[]): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-overlay-actions";
  row.append(...buttons);
  return row;
}

export function createButton(
  label: string,
  opts?: { primary?: boolean; danger?: boolean; compact?: boolean; onClick?: () => void },
): HTMLButtonElement {
  const btn = document.createElement("button");
  btn.type = "button";
  const classes = ["pi-overlay-btn"];
  if (opts?.primary) classes.push("pi-overlay-btn--primary");
  if (opts?.danger) classes.push("pi-overlay-btn--danger");
  if (opts?.compact) classes.push("pi-overlay-btn--compact");
  btn.className = classes.join(" ");
  btn.textContent = label;
  if (opts?.onClick) btn.addEventListener("click", opts.onClick);
  return btn;
}
