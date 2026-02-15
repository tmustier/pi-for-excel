/**
 * Skills catalog overlay.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";
import { getFilesWorkspace } from "../../files/workspace.js";
import { mergeAgentSkillDefinitions, listAgentSkills } from "../../skills/catalog.js";
import {
  filterAgentSkillsByEnabledState,
  loadDisabledSkillNamesFromSettings,
  setSkillEnabledInSettings,
  type SkillActivationMutableSettingsStore,
} from "../../skills/activation-store.js";
import {
  loadExternalAgentSkillsFromWorkspace,
  removeExternalAgentSkillFromWorkspace,
  upsertExternalAgentSkillInWorkspace,
} from "../../skills/external-store.js";
import { dispatchSkillsChanged } from "../../skills/events.js";
import type { AgentSkillDefinition } from "../../skills/types.js";
import {
  closeOverlayById,
  createOverlayBadge,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { SKILLS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";

interface SkillsSnapshot {
  bundled: AgentSkillDefinition[];
  external: AgentSkillDefinition[];
  active: AgentSkillDefinition[];
  disabledNames: Set<string>;
  externalLoadError: string | null;
  activationLoadError: string | null;
}

function normalizeSkillName(name: string): string {
  return name.trim().toLowerCase();
}

function formatSkillCount(count: number): string {
  return `${count} skill${count === 1 ? "" : "s"}`;
}

function createSkillItem(args: {
  skill: AgentSkillDefinition;
  active: boolean;
  enabled: boolean;
  shadowed: boolean;
  removable: boolean;
  toggleable: boolean;
  busy: boolean;
  onRemove?: () => void;
  onToggleEnabled?: () => void;
}): HTMLElement {
  const item = document.createElement("div");
  item.className = "pi-overlay-surface pi-skills-item";

  const top = document.createElement("div");
  top.className = "pi-skills-item__top";

  const name = document.createElement("strong");
  name.className = "pi-skills-item__name";
  name.textContent = args.skill.name;

  const badges = document.createElement("div");
  badges.className = "pi-overlay-badges";
  badges.appendChild(createOverlayBadge(args.skill.sourceKind, "muted"));
  badges.appendChild(createOverlayBadge(args.active ? "active" : "inactive", args.active ? "ok" : "muted"));
  if (!args.enabled) {
    badges.appendChild(createOverlayBadge("disabled", "warn"));
  }
  if (args.shadowed) {
    badges.appendChild(createOverlayBadge("shadowed", "warn"));
  }

  top.append(name, badges);

  const description = document.createElement("div");
  description.className = "pi-skills-item__description";
  description.textContent = args.skill.description;

  const meta = document.createElement("div");
  meta.className = "pi-skills-item__meta";

  const location = document.createElement("code");
  location.className = "pi-skills-item__location";
  location.textContent = args.skill.location;
  meta.appendChild(location);

  if (args.skill.compatibility) {
    const compatibility = document.createElement("span");
    compatibility.className = "pi-skills-item__compatibility";
    compatibility.textContent = `Compatibility: ${args.skill.compatibility}`;
    meta.appendChild(compatibility);
  }

  item.append(top, description, meta);

  const actions = document.createElement("div");
  actions.className = "pi-overlay-actions pi-overlay-actions--inline";

  if (args.toggleable && args.onToggleEnabled) {
    const toggleButton = createOverlayButton({
      text: args.enabled ? "Disable" : "Enable",
      className: "pi-overlay-btn--compact",
    });
    toggleButton.disabled = args.busy;
    toggleButton.addEventListener("click", () => {
      args.onToggleEnabled?.();
    });

    actions.appendChild(toggleButton);
  }

  if (args.removable && args.onRemove) {
    const removeButton = createOverlayButton({
      text: "Remove",
      className: "pi-overlay-btn--compact pi-overlay-btn--danger",
    });
    removeButton.disabled = args.busy;
    removeButton.addEventListener("click", () => {
      args.onRemove?.();
    });

    actions.appendChild(removeButton);
  }

  if (actions.childElementCount > 0) {
    item.appendChild(actions);
  }

  return item;
}

function renderSkillList(args: {
  container: HTMLElement;
  skills: readonly AgentSkillDefinition[];
  activeNames: ReadonlySet<string>;
  disabledNames: ReadonlySet<string>;
  shadowedNames?: ReadonlySet<string>;
  emptyMessage: string;
  removable?: boolean;
  toggleable?: boolean;
  busy: boolean;
  onRemove?: (skillName: string) => void;
  onToggleEnabled?: (skillName: string) => void;
}): void {
  args.container.replaceChildren();

  if (args.skills.length === 0) {
    const empty = document.createElement("div");
    empty.className = "pi-overlay-empty";
    empty.textContent = args.emptyMessage;
    args.container.appendChild(empty);
    return;
  }

  for (const skill of args.skills) {
    const normalizedName = normalizeSkillName(skill.name);
    const shadowed = args.shadowedNames?.has(normalizedName) ?? false;

    args.container.appendChild(createSkillItem({
      skill,
      active: args.activeNames.has(normalizedName),
      enabled: !args.disabledNames.has(normalizedName),
      shadowed,
      removable: args.removable === true,
      toggleable: args.toggleable === true && !shadowed,
      busy: args.busy,
      onRemove: args.onRemove ? () => args.onRemove?.(skill.name) : undefined,
      onToggleEnabled: args.onToggleEnabled ? () => args.onToggleEnabled?.(skill.name) : undefined,
    }));
  }
}

async function buildSnapshot(settings: SkillActivationMutableSettingsStore): Promise<SkillsSnapshot> {
  const bundled = listAgentSkills();

  let external: AgentSkillDefinition[] = [];
  let externalLoadError: string | null = null;

  try {
    external = await loadExternalAgentSkillsFromWorkspace(getFilesWorkspace());
  } catch (error: unknown) {
    externalLoadError = error instanceof Error ? error.message : "Unknown error";
    console.warn("[skills] Failed to load external skills for UI:", error);
  }

  let disabledNames = new Set<string>();
  let activationLoadError: string | null = null;

  try {
    disabledNames = await loadDisabledSkillNamesFromSettings(settings);
  } catch (error: unknown) {
    activationLoadError = error instanceof Error ? error.message : "Unknown error";
    console.warn("[skills] Failed to load skill activation state for UI:", error);
  }

  const discoverable = mergeAgentSkillDefinitions(bundled, external);

  const active = filterAgentSkillsByEnabledState({
    skills: discoverable,
    disabledSkillNames: disabledNames,
  });

  return {
    bundled,
    external,
    active,
    disabledNames,
    externalLoadError,
    activationLoadError,
  };
}

export function showSkillsDialog(): void {
  if (closeOverlayById(SKILLS_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: SKILLS_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-skills-dialog",
  });

  const settings = getAppStorage().settings;

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: "Close skills",
    title: "Skills",
    subtitle: "Browse bundled and external SKILL.md workflows used by the agent.",
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body";

  const summaryCard = document.createElement("div");
  summaryCard.className = "pi-overlay-surface";

  const summaryTitle = document.createElement("h3");
  summaryTitle.className = "pi-overlay-section-title";
  summaryTitle.textContent = "Prompt injection status";

  const summaryText = document.createElement("p");
  summaryText.className = "pi-skills-summary";
  summaryText.textContent = "Loading skills…";

  const defaultSummaryHint = "Enabled skills are injected when relevant; disable a skill here to keep it out of the prompt.";
  const summaryHint = document.createElement("p");
  summaryHint.className = "pi-overlay-hint";
  summaryHint.textContent = defaultSummaryHint;

  summaryCard.append(summaryTitle, summaryText, summaryHint);

  const activeSection = document.createElement("section");
  activeSection.className = "pi-overlay-section";
  activeSection.appendChild(createOverlaySectionTitle("Active skills"));

  const activeList = document.createElement("div");
  activeList.className = "pi-overlay-list";
  activeSection.appendChild(activeList);

  const bundledSection = document.createElement("section");
  bundledSection.className = "pi-overlay-section";
  bundledSection.appendChild(createOverlaySectionTitle("Bundled skills"));

  const bundledList = document.createElement("div");
  bundledList.className = "pi-overlay-list";
  bundledSection.appendChild(bundledList);

  const externalSection = document.createElement("section");
  externalSection.className = "pi-overlay-section";
  externalSection.appendChild(createOverlaySectionTitle("External skills"));

  const installCard = document.createElement("div");
  installCard.className = "pi-overlay-surface";

  const installTitle = document.createElement("p");
  installTitle.className = "pi-skills-install-title";
  installTitle.textContent = "Add or update external skill (paste full SKILL.md):";

  const installInput = document.createElement("textarea");
  installInput.className = "pi-skills-install-input";
  installInput.setAttribute("aria-label", "External skill markdown");
  installInput.placeholder = "---\nname: my-skill\ndescription: What this skill does\n---\n\n# Instructions\n...";

  const installActions = document.createElement("div");
  installActions.className = "pi-overlay-actions pi-overlay-actions--inline";

  const installButton = createOverlayButton({
    text: "Install / update",
    className: "pi-overlay-btn--primary",
  });

  const clearButton = createOverlayButton({
    text: "Clear",
    className: "pi-overlay-btn--compact",
  });

  installActions.append(installButton, clearButton);

  const installHint = document.createElement("p");
  installHint.className = "pi-overlay-hint";
  installHint.textContent = "External skills are local to this profile. If a bundled skill has the same name, the bundled one stays active.";

  installCard.append(installTitle, installInput, installActions, installHint);

  const externalList = document.createElement("div");
  externalList.className = "pi-overlay-list";

  externalSection.append(installCard, externalList);

  body.append(summaryCard, activeSection, bundledSection, externalSection);
  dialog.card.append(header, body);

  let snapshot: SkillsSnapshot | null = null;
  let busy = false;

  const setBusy = (next: boolean): void => {
    busy = next;
    installInput.disabled = next;
    installButton.disabled = next;
    clearButton.disabled = next;

    if (snapshot) {
      render(snapshot);
    }
  };

  const toggleSkillEnabled = (skillName: string, disabledNames: ReadonlySet<string>): void => {
    if (busy) return;

    const normalizedName = normalizeSkillName(skillName);
    const currentlyEnabled = !disabledNames.has(normalizedName);

    void (async () => {
      setBusy(true);
      try {
        const result = await setSkillEnabledInSettings({
          settings,
          name: skillName,
          enabled: !currentlyEnabled,
        });

        if (result.changed) {
          showToast(`${result.enabled ? "Enabled" : "Disabled"} skill: ${skillName}`);
          dispatchSkillsChanged({ reason: "activation" });
        } else {
          showToast(`Skill already ${result.enabled ? "enabled" : "disabled"}: ${skillName}`);
        }

        snapshot = await buildSnapshot(settings);
        render(snapshot);
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Failed to update skill activation: ${message}`);
      } finally {
        setBusy(false);
      }
    })();
  };

  const render = (current: SkillsSnapshot): void => {
    const activeNames = new Set(current.active.map((skill) => normalizeSkillName(skill.name)));
    const bundledNames = new Set(current.bundled.map((skill) => normalizeSkillName(skill.name)));
    const shadowedExternalNames = new Set(
      current.external
        .map((skill) => normalizeSkillName(skill.name))
        .filter((name) => bundledNames.has(name)),
    );

    summaryText.textContent = `Prompt currently includes ${formatSkillCount(current.active.length)} (bundled + discoverable external).`;

    summaryHint.textContent = defaultSummaryHint;
    summaryHint.classList.remove("pi-overlay-text-warning");

    const warningMessages: string[] = [];

    if (current.externalLoadError) {
      warningMessages.push(`External skills could not be loaded (${current.externalLoadError}).`);
    }

    if (current.activationLoadError) {
      warningMessages.push(`Skill activation state could not be loaded (${current.activationLoadError}).`);
    }

    if (warningMessages.length > 0) {
      summaryHint.textContent = `${warningMessages.join(" ")} Showing default behavior.`;
      summaryHint.classList.add("pi-overlay-text-warning");
    }

    renderSkillList({
      container: activeList,
      skills: current.active,
      activeNames,
      disabledNames: current.disabledNames,
      emptyMessage: "No active skills.",
      busy,
    });

    renderSkillList({
      container: bundledList,
      skills: current.bundled,
      activeNames,
      disabledNames: current.disabledNames,
      emptyMessage: "No bundled skills are available in this build.",
      toggleable: true,
      busy,
      onToggleEnabled: (skillName: string) => {
        toggleSkillEnabled(skillName, current.disabledNames);
      },
    });

    const externalEmptyMessage = "No external skills are configured. Add one using the form above.";

    renderSkillList({
      container: externalList,
      skills: current.external,
      activeNames,
      disabledNames: current.disabledNames,
      shadowedNames: shadowedExternalNames,
      emptyMessage: externalEmptyMessage,
      removable: true,
      toggleable: true,
      busy,
      onToggleEnabled: (skillName: string) => {
        toggleSkillEnabled(skillName, current.disabledNames);
      },
      onRemove: (skillName: string) => {
        if (busy) return;

        void (async () => {
          setBusy(true);
          try {
            const removed = await removeExternalAgentSkillFromWorkspace({
              workspace: getFilesWorkspace(),
              name: skillName,
            });

            if (!removed) {
              showToast(`External skill not found: ${skillName}`);
            } else {
              showToast(`Removed external skill: ${skillName}`);
              dispatchSkillsChanged({ reason: "catalog" });
            }

            snapshot = await buildSnapshot(settings);
            render(snapshot);
          } catch (error: unknown) {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Failed to remove external skill: ${message}`);
          } finally {
            setBusy(false);
          }
        })();
      },
    });
  };

  installButton.addEventListener("click", () => {
    if (busy) return;

    const markdown = installInput.value.trim();
    if (markdown.length === 0) {
      showToast("Paste a SKILL.md document first.");
      return;
    }

    void (async () => {
      setBusy(true);
      try {
        const result = await upsertExternalAgentSkillInWorkspace({
          workspace: getFilesWorkspace(),
          markdown,
        });

        showToast(`Saved external skill: ${result.name}`);
        installInput.value = "";
        dispatchSkillsChanged({ reason: "catalog" });

        snapshot = await buildSnapshot(settings);
        render(snapshot);
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Failed to save external skill: ${message}`);
      } finally {
        setBusy(false);
      }
    })();
  });

  clearButton.addEventListener("click", () => {
    if (busy) return;
    installInput.value = "";
    installInput.focus();
  });

  void (async () => {
    try {
      snapshot = await buildSnapshot(settings);
      render(snapshot);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Unknown error";
      summaryText.textContent = `Failed to load skills: ${message}`;

      activeList.replaceChildren();
      bundledList.replaceChildren();
      externalList.replaceChildren();
    }
  })();

  dialog.mount();
}
