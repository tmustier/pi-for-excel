/**
 * Skills catalog overlay.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";
import { isExperimentalFeatureEnabled } from "../../experiments/flags.js";
import { mergeAgentSkillDefinitions, listAgentSkills } from "../../skills/catalog.js";
import { loadExternalAgentSkillsFromSettings } from "../../skills/external-store.js";
import type { AgentSkillDefinition } from "../../skills/types.js";
import {
  closeOverlayById,
  createOverlayBadge,
  createOverlayDialog,
  createOverlayHeader,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { SKILLS_OVERLAY_ID } from "../../ui/overlay-ids.js";

interface SkillsSnapshot {
  bundled: AgentSkillDefinition[];
  external: AgentSkillDefinition[];
  active: AgentSkillDefinition[];
  externalDiscoveryEnabled: boolean;
  externalLoadError: string | null;
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
  shadowed: boolean;
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
  return item;
}

function renderSkillList(args: {
  container: HTMLElement;
  skills: readonly AgentSkillDefinition[];
  activeNames: ReadonlySet<string>;
  shadowedNames?: ReadonlySet<string>;
  emptyMessage: string;
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
    args.container.appendChild(createSkillItem({
      skill,
      active: args.activeNames.has(normalizedName),
      shadowed: args.shadowedNames?.has(normalizedName) ?? false,
    }));
  }
}

async function buildSnapshot(): Promise<SkillsSnapshot> {
  const settings = getAppStorage().settings;
  const bundled = listAgentSkills();

  let external: AgentSkillDefinition[] = [];
  let externalLoadError: string | null = null;

  try {
    external = await loadExternalAgentSkillsFromSettings(settings);
  } catch (error: unknown) {
    externalLoadError = error instanceof Error ? error.message : "Unknown error";
    console.warn("[skills] Failed to load external skills for UI:", error);
  }

  const externalDiscoveryEnabled = isExperimentalFeatureEnabled("external_skills_discovery");

  const active = externalDiscoveryEnabled
    ? mergeAgentSkillDefinitions(bundled, external)
    : bundled;

  return {
    bundled,
    external,
    active,
    externalDiscoveryEnabled,
    externalLoadError,
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

  const summaryHint = document.createElement("p");
  summaryHint.className = "pi-overlay-hint";
  summaryHint.textContent = "Skills are injected when relevant; the agent can read full SKILL.md content with the skills tool.";

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

  const externalList = document.createElement("div");
  externalList.className = "pi-overlay-list";
  externalSection.appendChild(externalList);

  body.append(summaryCard, activeSection, bundledSection, externalSection);
  dialog.card.append(header, body);

  void (async () => {
    try {
      const snapshot = await buildSnapshot();

      const activeNames = new Set(snapshot.active.map((skill) => normalizeSkillName(skill.name)));
      const bundledNames = new Set(snapshot.bundled.map((skill) => normalizeSkillName(skill.name)));
      const shadowedExternalNames = new Set(
        snapshot.external
          .map((skill) => normalizeSkillName(skill.name))
          .filter((name) => bundledNames.has(name)),
      );

      if (snapshot.externalDiscoveryEnabled) {
        summaryText.textContent = `Prompt currently includes ${formatSkillCount(snapshot.active.length)} (bundled + discoverable external).`;
      } else {
        summaryText.textContent = `Prompt currently includes ${formatSkillCount(snapshot.active.length)} (bundled only). External discovery is disabled.`;
      }

      if (snapshot.externalLoadError) {
        summaryHint.textContent = `External skills could not be loaded (${snapshot.externalLoadError}). Showing bundled skills only.`;
        summaryHint.classList.add("pi-overlay-text-warning");
      }

      renderSkillList({
        container: activeList,
        skills: snapshot.active,
        activeNames,
        emptyMessage: "No active skills.",
      });

      renderSkillList({
        container: bundledList,
        skills: snapshot.bundled,
        activeNames,
        emptyMessage: "No bundled skills are available in this build.",
      });

      const externalEmptyMessage = snapshot.externalDiscoveryEnabled
        ? "No external skills are configured."
        : "No external skills are configured. Enable /experimental on external_skills_discovery after adding one to activate it.";

      renderSkillList({
        container: externalList,
        skills: snapshot.external,
        activeNames,
        shadowedNames: shadowedExternalNames,
        emptyMessage: externalEmptyMessage,
      });
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
