/**
 * skills — list/read bundled Agent Skills (SKILL.md), with optional external discovery.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";

import { filterAgentSkillsByEnabledState } from "../skills/activation-store.js";
import type { AgentSkillDefinition } from "../skills/types.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import type {
  SkillsErrorDetails,
  SkillsListDetails,
  SkillsReadDetails,
  SkillsToolDetails,
} from "./tool-details.js";

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("list"),
    Type.Literal("read"),
  ], {
    description: "list = show available skills, read = return full SKILL.md for a named skill.",
  }),
  name: Type.Optional(Type.String({
    description: "Skill name (required when action=read).",
  })),
  refresh: Type.Optional(Type.Boolean({
    description: "When true (read only), bypass the session cache and reload from current skill sources.",
  })),
});

type Params = Static<typeof schema>;

export interface SkillsToolCatalog {
  list: () => AgentSkillDefinition[] | Promise<AgentSkillDefinition[]>;
}

let defaultCatalog: SkillsToolCatalog | null = null;

async function getDefaultCatalog(): Promise<SkillsToolCatalog> {
  if (defaultCatalog) {
    return defaultCatalog;
  }

  const catalogModule = await import("../skills/catalog.js");
  defaultCatalog = {
    list: catalogModule.listAgentSkills,
  };
  return defaultCatalog;
}

function defaultIsExternalDiscoveryEnabled(): boolean {
  return true;
}

async function defaultLoadExternalSkills(): Promise<AgentSkillDefinition[]> {
  const [{ getFilesWorkspace }, { loadExternalAgentSkillsFromWorkspace }] = await Promise.all([
    import("../files/workspace.js"),
    import("../skills/external-store.js"),
  ]);

  return loadExternalAgentSkillsFromWorkspace(getFilesWorkspace());
}

async function defaultLoadDisabledSkillNames(): Promise<Set<string>> {
  if (typeof window === "undefined") {
    return new Set();
  }

  try {
    const [{ getAppStorage }, { loadDisabledSkillNamesFromSettings }] = await Promise.all([
      import("@mariozechner/pi-web-ui/dist/storage/app-storage.js"),
      import("../skills/activation-store.js"),
    ]);

    return loadDisabledSkillNamesFromSettings(getAppStorage().settings);
  } catch (error: unknown) {
    console.warn("[skills] Failed to load skill activation state for tool:", error);
    return new Set();
  }
}

export interface SkillsToolDependencies {
  getSessionId?: () => string | null;
  readCache?: SkillReadCache;
  catalog?: SkillsToolCatalog;
  isExternalDiscoveryEnabled?: () => boolean | Promise<boolean>;
  loadExternalSkills?: () => Promise<AgentSkillDefinition[]>;
  loadDisabledSkillNames?: () => Promise<Set<string>>;
}

function renderSkillListMarkdown(args: {
  skills: readonly AgentSkillDefinition[];
  externalDiscoveryEnabled: boolean;
}): string {
  if (args.skills.length === 0) {
    return args.externalDiscoveryEnabled
      ? "No Agent Skills found. External discovery is enabled, but no valid skills are configured."
      : "No Agent Skills are bundled in this build.";
  }

  const lines: string[] = [
    `Available Agent Skills (${args.skills.length}):`,
    "",
  ];

  for (const skill of args.skills) {
    lines.push(
      `- \`${skill.name}\` — ${skill.description} _(source: ${skill.sourceKind}, location: ${skill.location})_`,
    );
  }

  lines.push("");
  lines.push("Use action=read with a skill name to load the full SKILL.md instructions.");
  return lines.join("\n");
}

function renderReadError(name: string, skills: readonly AgentSkillDefinition[]): string {
  const available = skills.map((skill) => `\`${skill.name}\``).join(", ");
  return [
    `Skill not found: \`${name}\`.`,
    available.length > 0 ? `Available: ${available}.` : "No skills are available.",
  ].join(" ");
}

function normalizeSessionId(value: string | null | undefined): string | null {
  if (typeof value !== "string") {
    return null;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function mergeSkillsByName(args: {
  preferred: readonly AgentSkillDefinition[];
  fallback: readonly AgentSkillDefinition[];
}): AgentSkillDefinition[] {
  const byName = new Map<string, AgentSkillDefinition>();

  for (const skill of args.preferred) {
    byName.set(skill.name.toLowerCase(), skill);
  }

  for (const skill of args.fallback) {
    const key = skill.name.toLowerCase();
    if (!byName.has(key)) {
      byName.set(key, skill);
    }
  }

  return Array.from(byName.values()).sort((left, right) => left.name.localeCompare(right.name));
}

function buildSkillsListDetails(args: {
  skills: readonly AgentSkillDefinition[];
  externalDiscoveryEnabled: boolean;
}): SkillsListDetails {
  return {
    kind: "skills_list",
    count: args.skills.length,
    names: args.skills.map((skill) => skill.name),
    entries: args.skills.map((skill) => ({
      name: skill.name,
      sourceKind: skill.sourceKind,
      location: skill.location,
    })),
    externalDiscoveryEnabled: args.externalDiscoveryEnabled,
  };
}

function buildSkillsErrorDetails(args: {
  message: string;
  externalDiscoveryEnabled: boolean;
  requestedName?: string;
  availableNames?: string[];
}): SkillsErrorDetails {
  return {
    kind: "skills_error",
    action: "read",
    message: args.message,
    requestedName: args.requestedName,
    availableNames: args.availableNames,
    externalDiscoveryEnabled: args.externalDiscoveryEnabled,
  };
}

function buildSkillsReadDetails(args: {
  skill: AgentSkillDefinition;
  cacheHit: boolean;
  refreshed: boolean;
  sessionScoped: boolean;
  readCount?: number;
}): SkillsReadDetails {
  return {
    kind: "skills_read",
    skillName: args.skill.name,
    sourceKind: args.skill.sourceKind,
    location: args.skill.location,
    cacheHit: args.cacheHit,
    refreshed: args.refreshed,
    sessionScoped: args.sessionScoped,
    readCount: args.readCount,
  };
}

async function resolveCatalogSkills(catalog: SkillsToolCatalog): Promise<AgentSkillDefinition[]> {
  const loaded = await catalog.list();
  return [...loaded].sort((left, right) => left.name.localeCompare(right.name));
}

async function resolveAllSkills(args: {
  catalog: SkillsToolCatalog;
  externalDiscoveryEnabled: boolean;
  loadExternalSkills: () => Promise<AgentSkillDefinition[]>;
  loadDisabledSkillNames: () => Promise<Set<string>>;
}): Promise<AgentSkillDefinition[]> {
  const bundled = await resolveCatalogSkills(args.catalog);

  const merged = args.externalDiscoveryEnabled
    ? mergeSkillsByName({
      preferred: bundled,
      fallback: await args.loadExternalSkills(),
    })
    : bundled;

  const disabledSkillNames = await args.loadDisabledSkillNames();
  return filterAgentSkillsByEnabledState({
    skills: merged,
    disabledSkillNames,
  });
}

export function createSkillsTool(
  dependencies: SkillsToolDependencies = {},
): AgentTool<typeof schema, SkillsToolDetails> {
  const getSessionId = dependencies.getSessionId;
  const readCache = dependencies.readCache;

  return {
    name: "skills",
    label: "Skills",
    description:
      "List and read bundled Agent Skills (SKILL.md). "
      + "Use this to load detailed, task-specific workflows on demand.",
    parameters: schema,
    execute: async (_toolCallId: string, params: Params): Promise<AgentToolResult<SkillsToolDetails>> => {
      const catalog = dependencies.catalog ?? await getDefaultCatalog();
      const isExternalDiscoveryEnabled = dependencies.isExternalDiscoveryEnabled ?? defaultIsExternalDiscoveryEnabled;
      const loadExternalSkills = dependencies.loadExternalSkills ?? defaultLoadExternalSkills;
      const loadDisabledSkillNames = dependencies.loadDisabledSkillNames ?? defaultLoadDisabledSkillNames;

      const externalDiscoveryEnabled = await isExternalDiscoveryEnabled();
      const skills = await resolveAllSkills({
        catalog,
        externalDiscoveryEnabled,
        loadExternalSkills,
        loadDisabledSkillNames,
      });

      if (params.action === "list") {
        return {
          content: [{
            type: "text",
            text: renderSkillListMarkdown({ skills, externalDiscoveryEnabled }),
          }],
          details: buildSkillsListDetails({ skills, externalDiscoveryEnabled }),
        };
      }

      const requestedName = params.name?.trim() ?? "";
      if (requestedName.length === 0) {
        const message = "Error: name is required when action=read.";
        return {
          content: [{ type: "text", text: message }],
          details: buildSkillsErrorDetails({
            message,
            externalDiscoveryEnabled,
            availableNames: skills.map((skill) => skill.name),
          }),
        };
      }

      const refresh = params.refresh === true;
      const sessionId = normalizeSessionId(getSessionId?.());
      const sessionScoped = sessionId !== null && readCache !== undefined;

      if (!refresh && sessionId && readCache) {
        const cached = readCache.get(sessionId, requestedName);
        if (cached) {
          const cachedSkillStillAvailable = skills.some(
            (entry) => entry.name.toLowerCase() === cached.skillName.toLowerCase(),
          );

          if (cachedSkillStillAvailable) {
            const cachedSkill: AgentSkillDefinition = {
              name: cached.skillName,
              description: "",
              location: cached.location,
              sourceKind: cached.sourceKind,
              markdown: cached.markdown,
              body: cached.markdown,
            };

            return {
              content: [{ type: "text", text: cached.markdown }],
              details: buildSkillsReadDetails({
                skill: cachedSkill,
                cacheHit: true,
                refreshed: false,
                sessionScoped,
                readCount: cached.readCount,
              }),
            };
          }
        }
      }

      const skill = skills.find((entry) => entry.name.toLowerCase() === requestedName.toLowerCase()) ?? null;
      if (!skill) {
        const message = renderReadError(requestedName, skills);
        return {
          content: [{ type: "text", text: message }],
          details: buildSkillsErrorDetails({
            message,
            externalDiscoveryEnabled,
            requestedName,
            availableNames: skills.map((entry) => entry.name),
          }),
        };
      }

      const cachedEntry = sessionId && readCache
        ? readCache.set(sessionId, {
          skillName: skill.name,
          sourceKind: skill.sourceKind,
          location: skill.location,
          markdown: skill.markdown,
        })
        : null;

      return {
        content: [{ type: "text", text: skill.markdown }],
        details: buildSkillsReadDetails({
          skill,
          cacheHit: false,
          refreshed: refresh,
          sessionScoped,
          readCount: cachedEntry?.readCount,
        }),
      };
    },
  };
}
