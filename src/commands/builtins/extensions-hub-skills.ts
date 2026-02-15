/**
 * Extensions hub — Skills tab.
 *
 * Bundled skills (read-only cards), external skills (expandable with remove),
 * and install form for pasting SKILL.md content.
 */

import { getFilesWorkspace } from "../../files/workspace.js";
import { listAgentSkills, mergeAgentSkillDefinitions } from "../../skills/catalog.js";
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
import { showToast } from "../../ui/toast.js";
import {
  createSectionHeader,
  createItemCard,
  createConfigRow,
  createConfigValue,
  createAddForm,
  createEmptyInline,
  createButton,
  createActionsRow,
  createToggle,
} from "../../ui/extensions-hub-components.js";

// ── Types ───────────────────────────────────────────

interface SkillsSnapshot {
  bundled: AgentSkillDefinition[];
  external: AgentSkillDefinition[];
  active: AgentSkillDefinition[];
  disabledNames: Set<string>;
  externalLoadError: string | null;
  activationLoadError: string | null;
}

// ── Snapshot builder ────────────────────────────────

async function buildSnapshot(settings: SkillActivationMutableSettingsStore): Promise<SkillsSnapshot> {
  const bundled = listAgentSkills();
  let external: AgentSkillDefinition[] = [];
  let externalLoadError: string | null = null;
  let disabledNames = new Set<string>();
  let activationLoadError: string | null = null;

  try {
    external = await loadExternalAgentSkillsFromWorkspace(getFilesWorkspace());
  } catch (err: unknown) {
    externalLoadError = err instanceof Error ? err.message : "Unknown error";
  }

  try {
    disabledNames = await loadDisabledSkillNamesFromSettings(settings);
  } catch (err: unknown) {
    activationLoadError = err instanceof Error ? err.message : "Unknown error";
  }

  const all = mergeAgentSkillDefinitions(bundled, external);
  const active = filterAgentSkillsByEnabledState({ skills: all, disabledSkillNames: disabledNames });

  return { bundled, external, active, disabledNames, externalLoadError, activationLoadError };
}

function normalizeSkillName(name: string): string {
  return name.trim().toLowerCase();
}

// ── Main render ─────────────────────────────────────

export async function renderSkillsTab(args: {
  container: HTMLElement;
  settings: SkillActivationMutableSettingsStore;
  isBusy: () => boolean;
  runMutation: (action: () => Promise<void>, reason: "toggle" | "scope" | "external-toggle" | "config", msg?: string) => Promise<void>;
}): Promise<void> {
  const { container, settings, isBusy, runMutation } = args;

  let snapshot: SkillsSnapshot;
  try {
    snapshot = await buildSnapshot(settings);
  } catch (err: unknown) {
    container.replaceChildren();
    const msg = document.createElement("p");
    msg.className = "pi-overlay-hint";
    msg.textContent = `Failed to load skills: ${err instanceof Error ? err.message : String(err)}`;
    container.appendChild(msg);
    return;
  }

  const bundledNames = new Set(snapshot.bundled.map((s) => normalizeSkillName(s.name)));
  const disabledNames = snapshot.disabledNames;
  const activeBundledCount = snapshot.active.filter((skill) => skill.sourceKind === "bundled").length;
  const activeExternalCount = snapshot.active.filter((skill) => skill.sourceKind === "external").length;

  const isSkillEnabled = (skillName: string): boolean => !disabledNames.has(normalizeSkillName(skillName));

  const toggleSkill = (skillName: string, enabled: boolean): void => {
    if (isBusy()) return;

    void runMutation(async () => {
      const result = await setSkillEnabledInSettings({
        settings,
        name: skillName,
        enabled,
      });
      if (result.changed) {
        dispatchSkillsChanged({ reason: "activation" });
      }
    }, "toggle", `${enabled ? "Enabled" : "Disabled"} skill: ${skillName}`);
  };

  container.replaceChildren();

  const statusLine = document.createElement("p");
  statusLine.className = "pi-overlay-hint";
  statusLine.textContent = `${snapshot.active.length} skills active (${activeBundledCount} bundled, ${activeExternalCount} external)`;
  container.appendChild(statusLine);

  // ── Bundled section ───────────────────────────
  container.appendChild(createSectionHeader({
    label: "Bundled skills",
    count: snapshot.bundled.length,
  }));

  if (snapshot.bundled.length === 0) {
    container.appendChild(createEmptyInline("📋", "No bundled skills in this build."));
  } else {
    const list = document.createElement("div");
    list.className = "pi-hub-stack";

    for (const skill of snapshot.bundled) {
      list.appendChild(renderBundledSkillCard({
        skill,
        enabled: isSkillEnabled(skill.name),
        busy: isBusy(),
        onToggle: (enabled) => {
          toggleSkill(skill.name, enabled);
        },
      }));
    }
    container.appendChild(list);
  }

  // ── External section ──────────────────────────
  container.appendChild(createSectionHeader({
    label: "External skills",
    count: snapshot.external.length,
  }));

  if (snapshot.external.length === 0) {
    container.appendChild(createEmptyInline("📋", "No external skills installed.\nPaste a SKILL.md below to add one."));
  } else {
    const list = document.createElement("div");
    list.className = "pi-hub-stack";

    for (const skill of snapshot.external) {
      const shadowed = bundledNames.has(normalizeSkillName(skill.name));
      list.appendChild(renderExternalSkillCard({
        skill,
        enabled: isSkillEnabled(skill.name),
        shadowed,
        busy: isBusy(),
        onToggle: (enabled) => {
          toggleSkill(skill.name, enabled);
        },
        onRemove: () => {
          if (isBusy()) return;
          void runMutation(async () => {
            const removed = await removeExternalAgentSkillFromWorkspace({
              workspace: getFilesWorkspace(),
              name: skill.name,
            });
            if (removed) {
              dispatchSkillsChanged({ reason: "catalog" });
            } else {
              throw new Error(`External skill not found: ${skill.name}`);
            }
          }, "config", `Removed external skill: ${skill.name}`);
        },
      }));
    }
    container.appendChild(list);
  }

  // Warnings
  if (snapshot.externalLoadError) {
    const warn = document.createElement("p");
    warn.className = "pi-overlay-hint pi-hub-warn-text";
    warn.textContent = `External skills load failed: ${snapshot.externalLoadError}`;
    container.appendChild(warn);
  }

  if (snapshot.activationLoadError) {
    const warn = document.createElement("p");
    warn.className = "pi-overlay-hint pi-hub-warn-text";
    warn.textContent = `Skill activation state unavailable: ${snapshot.activationLoadError}`;
    container.appendChild(warn);
  }

  // ── Install section ───────────────────────────
  container.appendChild(createSectionHeader({ label: "Install skill" }));

  const installForm = createAddForm();

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = "Paste a SKILL.md document below to install an external skill.";
  installForm.appendChild(hint);

  const textarea = document.createElement("textarea");
  textarea.className = "pi-overlay-input pi-hub-textarea";
  textarea.placeholder = "---\nname: my-skill\ndescription: What this skill does\n---\n\nInstructions for the agent...";
  installForm.appendChild(textarea);

  const installActions = document.createElement("div");
  installActions.className = "pi-hub-actions-end";
  installActions.appendChild(createButton("Install skill", {
    primary: true,
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      const md = textarea.value.trim();
      if (!md) { showToast("Paste a SKILL.md document first."); return; }

      void runMutation(async () => {
        const result = await upsertExternalAgentSkillInWorkspace({
          workspace: getFilesWorkspace(),
          markdown: md,
        });
        showToast(`Saved external skill: ${result.name}`);
        textarea.value = "";
        dispatchSkillsChanged({ reason: "catalog" });
      }, "config");
    },
  }));
  installForm.appendChild(installActions);
  container.appendChild(installForm);

  // Footer hint
  const footer = document.createElement("p");
  footer.className = "pi-overlay-hint";
  footer.textContent = "Skills are instruction documents the AI reads on-demand to learn new workflows. They don't run code — they teach.";
  container.appendChild(footer);
}

// ── Bundled skill card ──────────────────────────────

function renderBundledSkillCard(args: {
  skill: AgentSkillDefinition;
  enabled: boolean;
  busy: boolean;
  onToggle: (enabled: boolean) => void;
}): HTMLElement {
  const toggle = createToggle({
    checked: args.enabled,
    onChange: args.onToggle,
    stopPropagation: true,
  });
  toggle.input.disabled = args.busy;

  const card = createItemCard({
    icon: "📋",
    iconColor: "amber",
    name: args.skill.name,
    description: args.skill.description,
    badges: [{ text: "Bundled", tone: "muted" }],
    rightContent: toggle.root,
  });

  return card.root;
}

// ── External skill card (expandable) ────────────────

function renderExternalSkillCard(args: {
  skill: AgentSkillDefinition;
  enabled: boolean;
  shadowed: boolean;
  busy: boolean;
  onToggle: (enabled: boolean) => void;
  onRemove: () => void;
}): HTMLElement {
  const badges: Array<{ text: string; tone: "ok" | "warn" | "muted" | "info" }> = [
    { text: "External", tone: "info" },
  ];
  if (args.shadowed) {
    badges.push({ text: "Shadowed", tone: "warn" });
  }

  const toggle = createToggle({
    checked: args.enabled,
    onChange: args.onToggle,
    stopPropagation: true,
  });
  toggle.input.disabled = args.busy || args.shadowed;

  const card = createItemCard({
    icon: "📋",
    iconColor: "amber",
    name: args.skill.name,
    description: args.skill.description,
    expandable: true,
    badges,
    rightContent: toggle.root,
  });

  // Location
  card.body.appendChild(createConfigRow("Location", createConfigValue(args.skill.location)));

  // Remove button
  card.body.appendChild(createActionsRow(
    createButton("Remove", {
      danger: true,
      compact: true,
      onClick: args.onRemove,
    }),
  ));

  return card.root;
}
