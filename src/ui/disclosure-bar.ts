/**
 * Disclosure bar — non-blocking banner shown after first provider connect.
 *
 * Informs the user about Pi's external capabilities (web search, extensions,
 * MCP, skills) and lets them acknowledge or customize before using the agent.
 */

const ACKNOWLEDGED_KEY = "pi.onboarding.disclosure.acknowledged";

function isAcknowledged(): boolean {
  try {
    return localStorage.getItem(ACKNOWLEDGED_KEY) === "1";
  } catch {
    return false;
  }
}

function setAcknowledged(): void {
  try {
    localStorage.setItem(ACKNOWLEDGED_KEY, "1");
  } catch {
    // ignore — private mode / storage unavailable
  }
}

export interface DisclosureBarOptions {
  /** Number of configured providers (bar only shows when ≥1). */
  providerCount: number;
  /** Callback to open Settings overlay. If provided, "Change anytime in Settings" becomes a link. */
  onOpenSettings?: () => void;
}

/**
 * Create and return the disclosure bar element, or `null` if already dismissed.
 *
 * The caller is responsible for inserting the element into the DOM.
 * The returned element removes itself when the user dismisses it.
 */
export function createDisclosureBar(options: DisclosureBarOptions): HTMLElement | null {
  if (isAcknowledged() || options.providerCount < 1) {
    return null;
  }

  const bar = document.createElement("div");
  bar.className = "pi-disclosure-bar";

  const text = document.createElement("div");
  text.className = "pi-disclosure-bar__text";
  text.textContent = "Pi can search the web, use extensions, and connect to external services.";
  bar.appendChild(text);

  // --- Expandable picker (hidden by default) ---
  const picker = document.createElement("div");
  picker.className = "pi-disclosure-picker";
  bar.appendChild(picker);

  const toggleRows: { label: string; sublabel: string }[] = [
    { label: "Web search & page fetch", sublabel: "Search engines and read web pages" },
    { label: "Extensions & plugins", sublabel: "Sidebar tools and custom commands" },
    { label: "External services (MCP)", sublabel: "Connect to tool servers you configure" },
    { label: "Skills", sublabel: "Instruction documents the AI follows" },
  ];

  for (const row of toggleRows) {
    const toggleRow = document.createElement("label");
    toggleRow.className = "pi-toggle-row";

    const labelWrap = document.createElement("div");
    labelWrap.className = "pi-toggle-row__labels";

    const labelEl = document.createElement("span");
    labelEl.className = "pi-toggle-row__label";
    labelEl.textContent = row.label;
    labelWrap.appendChild(labelEl);

    const sublabel = document.createElement("span");
    sublabel.className = "pi-toggle-row__sublabel";
    sublabel.textContent = row.sublabel;
    labelWrap.appendChild(sublabel);

    toggleRow.appendChild(labelWrap);

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.className = "pi-toggle__input";

    const track = document.createElement("span");
    track.className = "pi-toggle__track";

    const toggle = document.createElement("span");
    toggle.className = "pi-toggle";
    toggle.appendChild(checkbox);
    toggle.appendChild(track);

    toggleRow.appendChild(toggle);
    picker.appendChild(toggleRow);
  }

  // --- Actions row ---
  const actions = document.createElement("div");
  actions.className = "pi-disclosure-bar__actions";
  bar.appendChild(actions);

  const dismiss = () => {
    setAcknowledged();
    bar.remove();
  };

  const gotItBtn = document.createElement("button");
  gotItBtn.className = "pi-overlay-btn pi-overlay-btn--primary pi-overlay-btn--compact";
  gotItBtn.textContent = "Got it";
  gotItBtn.addEventListener("click", dismiss);
  actions.appendChild(gotItBtn);

  const customizeBtn = document.createElement("button");
  customizeBtn.className = "pi-disclosure-bar__link";
  customizeBtn.textContent = "Customize";
  actions.appendChild(customizeBtn);

  let hint: HTMLElement;
  if (options.onOpenSettings) {
    const link = document.createElement("button");
    link.type = "button";
    link.className = "pi-disclosure-bar__settings-link";
    link.textContent = "Change anytime in Settings";
    link.addEventListener("click", () => {
      dismiss();
      options.onOpenSettings?.();
    });
    hint = link;
  } else {
    const span = document.createElement("span");
    span.className = "pi-disclosure-bar__muted";
    span.textContent = "· Change anytime in Settings";
    hint = span;
  }
  actions.appendChild(hint);

  customizeBtn.addEventListener("click", () => {
    const isVisible = picker.classList.toggle("is-visible");
    if (isVisible) {
      gotItBtn.textContent = "Done";
      customizeBtn.style.display = "none";
      hint.style.display = "none";
    } else {
      gotItBtn.textContent = "Got it";
      customizeBtn.style.display = "";
      hint.style.display = "";
    }
  });

  return bar;
}
