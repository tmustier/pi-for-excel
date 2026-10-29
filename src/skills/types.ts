export type AgentSkillSourceKind = "bundled" | "external";

export interface AgentSkillDefinition {
  name: string;
  description: string;
  compatibility?: string;
  location: string;
  sourceKind: AgentSkillSourceKind;
  markdown: string;
  body: string;
}

export interface AgentSkillPromptEntry {
  name: string;
  description: string;
  location: string;
}
