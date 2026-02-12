/**
 * Skill-related UI/runtime refresh events.
 */

export const PI_SKILLS_CHANGED_EVENT = "pi:skills-changed";

export interface SkillsChangedDetail {
  reason: "toggle" | "scope" | "external-toggle" | "config";
}

export function dispatchSkillsChanged(detail: SkillsChangedDetail): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(
    new CustomEvent<SkillsChangedDetail>(PI_SKILLS_CHANGED_EVENT, { detail }),
  );
}
