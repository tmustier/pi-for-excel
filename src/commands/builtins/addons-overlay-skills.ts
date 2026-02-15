import { getFilesWorkspace } from "../../files/workspace.js";
import {
  filterAgentSkillsByEnabledState,
  loadDisabledSkillNamesFromSettings,
} from "../../skills/activation-store.js";
import {
  listAgentSkills,
  mergeAgentSkillDefinitions,
} from "../../skills/catalog.js";
import { loadExternalAgentSkillsFromWorkspace } from "../../skills/external-store.js";
import {
  createOverlayBadge,
  createOverlayButton,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import type { AgentSkillDefinition } from "../../skills/types.js";
import type {
  AddonsDialogActions,
  AddonsSettingsStore,
  SkillsSnapshot,
} from "./addons-overlay-types.js";

export async function buildSkillsSnapshot(settings: AddonsSettingsStore): Promise<SkillsSnapshot> {
  const bundled = listAgentSkills();

  let external: AgentSkillDefinition[] = [];
  let externalLoadError: string | null = null;
  try {
    external = await loadExternalAgentSkillsFromWorkspace(getFilesWorkspace());
  } catch (error: unknown) {
    externalLoadError = error instanceof Error ? error.message : "Unknown error";
  }

  let disabledNames = new Set<string>();
  let activationLoadError: string | null = null;
  try {
    disabledNames = await loadDisabledSkillNamesFromSettings(settings);
  } catch (error: unknown) {
    activationLoadError = error instanceof Error ? error.message : "Unknown error";
  }

  const skills = mergeAgentSkillDefinitions(bundled, external);

  const activeSkills = filterAgentSkillsByEnabledState({
    skills,
    disabledSkillNames: disabledNames,
  });

  return {
    skills,
    activeNames: new Set(activeSkills.map((skill) => skill.name.trim().toLowerCase())),
    disabledNames,
    externalLoadError,
    activationLoadError,
  };
}

export function renderSkillsSection(args: {
  container: HTMLElement;
  actions: AddonsDialogActions;
  snapshot: SkillsSnapshot;
}): void {
  args.container.replaceChildren();

  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-addons-section";
  section.dataset.addonsSection = "skills";
  section.appendChild(createOverlaySectionTitle("Skills"));

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = "Read-only view of bundled and external agent skills.";
  section.appendChild(hint);

  const list = document.createElement("div");
  list.className = "pi-overlay-list";

  if (args.snapshot.skills.length === 0) {
    const empty = document.createElement("div");
    empty.className = "pi-overlay-empty";
    empty.textContent = "No skills available.";
    list.appendChild(empty);
  } else {
    for (const skill of args.snapshot.skills) {
      const normalizedName = skill.name.trim().toLowerCase();
      const active = args.snapshot.activeNames.has(normalizedName);
      const disabled = args.snapshot.disabledNames.has(normalizedName);

      const card = document.createElement("div");
      card.className = "pi-overlay-surface pi-addons-skill";

      const top = document.createElement("div");
      top.className = "pi-addons-skill__top";

      const name = document.createElement("strong");
      name.className = "pi-addons-skill__name";
      name.textContent = skill.name;

      const badges = document.createElement("div");
      badges.className = "pi-overlay-badges";
      badges.appendChild(createOverlayBadge(skill.sourceKind, "muted"));
      badges.appendChild(createOverlayBadge(active ? "active" : "inactive", active ? "ok" : "muted"));
      if (disabled) {
        badges.appendChild(createOverlayBadge("disabled", "warn"));
      }

      top.append(name, badges);

      const description = document.createElement("div");
      description.className = "pi-addons-skill__description";
      description.textContent = skill.description;

      const location = document.createElement("code");
      location.className = "pi-addons-skill__location";
      location.textContent = skill.location;

      card.append(top, description, location);
      list.appendChild(card);
    }
  }

  section.appendChild(list);

  if (args.snapshot.externalLoadError) {
    const warning = document.createElement("p");
    warning.className = "pi-overlay-hint pi-overlay-text-warning";
    warning.textContent = `External skills load failed: ${args.snapshot.externalLoadError}`;
    section.appendChild(warning);
  }

  if (args.snapshot.activationLoadError) {
    const warning = document.createElement("p");
    warning.className = "pi-overlay-hint pi-overlay-text-warning";
    warning.textContent = `Skill activation state unavailable: ${args.snapshot.activationLoadError}`;
    section.appendChild(warning);
  }

  const actionsRow = document.createElement("div");
  actionsRow.className = "pi-overlay-actions";

  const openButton = createOverlayButton({ text: "Open full Skills manager…" });
  openButton.addEventListener("click", () => {
    args.actions.openSkillsManager();
  });

  actionsRow.appendChild(openButton);
  section.appendChild(actionsRow);

  args.container.appendChild(section);
}
