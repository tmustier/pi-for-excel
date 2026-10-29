export const PI_SKILLS_CHANGED_EVENT = "pi:skills-changed";

export interface SkillsChangedDetail {
  reason: "catalog" | "activation";
}

export function dispatchSkillsChanged(detail: SkillsChangedDetail): void {
  if (typeof document === "undefined") {
    return;
  }

  document.dispatchEvent(
    new CustomEvent<SkillsChangedDetail>(PI_SKILLS_CHANGED_EVENT, { detail }),
  );
}
