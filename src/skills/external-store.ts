/**
 * External Agent Skills discovery store.
 *
 * Managed external installs:
 * - skills/external/<name>/SKILL.md
 *
 * Workspace-discovered skills (auto-discovery):
 * - skills/<name>/SKILL.md
 */

import { getFilesWorkspace, type FilesWorkspace } from "../files/workspace.js";
import type { WorkspaceFileEntry } from "../files/types.js";
import type {
  AgentSkillDefinition,
  AgentSkillSourceKind,
} from "./types.js";
import { parseSkillDocument } from "./frontmatter.js";

const WORKSPACE_SKILLS_ROOT_PATH = "skills";
const MANAGED_EXTERNAL_SKILLS_ROOT_PATH = "skills/external";
const SKILL_FILENAME = "SKILL.md";
const MAX_EXTERNAL_SKILL_MARKDOWN_CHARS = 1_000_000;

export type ExternalSkillWorkspace = Pick<
  FilesWorkspace,
  "listFiles" | "readFile" | "writeTextFile" | "deleteFile"
>;

function isManagedExternalSkillFile(file: WorkspaceFileEntry): boolean {
  if (file.sourceKind !== "workspace") {
    return false;
  }

  const parts = file.path.split("/");
  return (
    parts.length === 4
    && parts[0] === "skills"
    && parts[1] === "external"
    && parts[2] !== ""
    && parts[3] === SKILL_FILENAME
  );
}

function isWorkspaceDiscoveredSkillFile(file: WorkspaceFileEntry): boolean {
  if (file.sourceKind !== "workspace") {
    return false;
  }

  const parts = file.path.split("/");
  return (
    parts.length === 3
    && parts[0] === WORKSPACE_SKILLS_ROOT_PATH
    && parts[1] !== ""
    && parts[1] !== "external"
    && parts[2] === SKILL_FILENAME
  );
}

function normalizeSkillName(name: string): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    throw new Error("Skill name cannot be empty.");
  }

  if (trimmed.includes("/") || trimmed.includes("\\")) {
    throw new Error("Skill name cannot contain path separators.");
  }

  if (trimmed === "." || trimmed === "..") {
    throw new Error("Skill name cannot be '.' or '..'.");
  }

  return trimmed;
}

function getExternalSkillPath(name: string): string {
  const normalized = normalizeSkillName(name);
  return `${MANAGED_EXTERNAL_SKILLS_ROOT_PATH}/${normalized}/${SKILL_FILENAME}`;
}

function buildExternalSkillDefinition(args: {
  location: string;
  markdown: string;
  sourceKind: AgentSkillSourceKind;
}): AgentSkillDefinition | null {
  const parsed = parseSkillDocument(args.markdown);
  if (!parsed) {
    return null;
  }

  return {
    name: parsed.frontmatter.name,
    description: parsed.frontmatter.description,
    compatibility: parsed.frontmatter.compatibility,
    location: args.location,
    sourceKind: args.sourceKind,
    markdown: args.markdown,
    body: parsed.body,
  };
}

function dedupeSkillsByName(skills: readonly AgentSkillDefinition[], duplicateLabel: string): AgentSkillDefinition[] {
  const byName = new Map<string, AgentSkillDefinition>();

  for (const skill of skills) {
    const normalizedName = skill.name.toLowerCase();
    if (byName.has(normalizedName)) {
      console.warn(`[skills] Duplicate ${duplicateLabel} skill ignored: ${skill.name} (${skill.location})`);
      continue;
    }

    byName.set(normalizedName, skill);
  }

  return Array.from(byName.values()).sort((left, right) => left.name.localeCompare(right.name));
}

async function loadSkillDefinitionsFromFiles(args: {
  workspace: ExternalSkillWorkspace;
  files: readonly WorkspaceFileEntry[];
  sourceKind: AgentSkillSourceKind;
  tooLargeWarningLabel: string;
  invalidWarningLabel: string;
}): Promise<AgentSkillDefinition[]> {
  const loaded: AgentSkillDefinition[] = [];

  for (const file of args.files) {
    let readResult: Awaited<ReturnType<ExternalSkillWorkspace["readFile"]>>;

    try {
      readResult = await args.workspace.readFile(file.path, {
        mode: "text",
        maxChars: MAX_EXTERNAL_SKILL_MARKDOWN_CHARS,
      });
    } catch (error: unknown) {
      console.warn(`[skills] Failed reading ${args.invalidWarningLabel}: ${file.path}`, error);
      continue;
    }

    if (typeof readResult.text !== "string") {
      console.warn(`[skills] ${args.invalidWarningLabel} is not readable text: ${file.path}`);
      continue;
    }

    if (readResult.truncated) {
      console.warn(`[skills] ${args.tooLargeWarningLabel} is too large to load fully: ${file.path}`);
      continue;
    }

    const skill = buildExternalSkillDefinition({
      location: file.path,
      markdown: readResult.text,
      sourceKind: args.sourceKind,
    });

    if (!skill) {
      console.warn(`[skills] Invalid SKILL.md frontmatter (${args.invalidWarningLabel}): ${file.path}`);
      continue;
    }

    loaded.push(skill);
  }

  return loaded;
}

function sortSkillFiles(files: readonly WorkspaceFileEntry[]): WorkspaceFileEntry[] {
  return [...files].sort((left, right) => left.path.localeCompare(right.path));
}

async function loadAllManagedExternalAgentSkillDefinitions(args: {
  workspace: ExternalSkillWorkspace;
  files?: readonly WorkspaceFileEntry[];
}): Promise<AgentSkillDefinition[]> {
  const files = args.files ?? await args.workspace.listFiles();
  const externalFiles = sortSkillFiles(files.filter((file) => isManagedExternalSkillFile(file)));

  return loadSkillDefinitionsFromFiles({
    workspace: args.workspace,
    files: externalFiles,
    sourceKind: "external",
    tooLargeWarningLabel: "managed external skill file",
    invalidWarningLabel: "managed external skill file",
  });
}

async function loadAllWorkspaceDiscoveredAgentSkillDefinitions(args: {
  workspace: ExternalSkillWorkspace;
  files?: readonly WorkspaceFileEntry[];
}): Promise<AgentSkillDefinition[]> {
  const files = args.files ?? await args.workspace.listFiles();
  const workspaceSkillFiles = sortSkillFiles(files.filter((file) => isWorkspaceDiscoveredSkillFile(file)));

  return loadSkillDefinitionsFromFiles({
    workspace: args.workspace,
    files: workspaceSkillFiles,
    sourceKind: "external",
    tooLargeWarningLabel: "workspace-discovered skill file",
    invalidWarningLabel: "workspace-discovered skill file",
  });
}

/**
 * Loads managed external skills from the canonical Files workspace location:
 * `skills/external/<name>/SKILL.md`.
 */
export async function loadExternalAgentSkillsFromWorkspace(
  workspace: ExternalSkillWorkspace,
): Promise<AgentSkillDefinition[]> {
  const loaded = await loadAllManagedExternalAgentSkillDefinitions({ workspace });
  return dedupeSkillsByName(loaded, "managed external");
}

/**
 * Loads auto-discovered workspace skills from:
 * `skills/<name>/SKILL.md` (excluding `skills/external/*`).
 */
export async function loadWorkspaceAgentSkillsFromWorkspace(
  workspace: ExternalSkillWorkspace,
): Promise<AgentSkillDefinition[]> {
  const loaded = await loadAllWorkspaceDiscoveredAgentSkillDefinitions({ workspace });
  return dedupeSkillsByName(loaded, "workspace-discovered");
}

/**
 * Loads all discoverable non-bundled skills.
 *
 * Precedence on name collisions:
 * 1) managed external (`skills/external/<name>/SKILL.md`)
 * 2) workspace-discovered (`skills/<name>/SKILL.md`)
 */
export async function loadDiscoverableAgentSkillsFromWorkspace(
  workspace: ExternalSkillWorkspace,
): Promise<AgentSkillDefinition[]> {
  const files = await workspace.listFiles();

  const [managedExternal, workspaceDiscovered] = await Promise.all([
    loadAllManagedExternalAgentSkillDefinitions({ workspace, files }).then((loaded) => {
      return dedupeSkillsByName(loaded, "managed external");
    }),
    loadAllWorkspaceDiscoveredAgentSkillDefinitions({ workspace, files }).then((loaded) => {
      return dedupeSkillsByName(loaded, "workspace-discovered");
    }),
  ]);

  return dedupeSkillsByName([...managedExternal, ...workspaceDiscovered], "discoverable");
}

export async function loadExternalAgentSkills(): Promise<AgentSkillDefinition[]> {
  return loadExternalAgentSkillsFromWorkspace(getFilesWorkspace());
}

export async function loadWorkspaceAgentSkills(): Promise<AgentSkillDefinition[]> {
  return loadWorkspaceAgentSkillsFromWorkspace(getFilesWorkspace());
}

export async function loadDiscoverableAgentSkills(): Promise<AgentSkillDefinition[]> {
  return loadDiscoverableAgentSkillsFromWorkspace(getFilesWorkspace());
}

export interface UpsertExternalAgentSkillResult {
  name: string;
  location: string;
}

export async function upsertExternalAgentSkillInWorkspace(args: {
  workspace: ExternalSkillWorkspace;
  markdown: string;
  expectedName?: string;
}): Promise<UpsertExternalAgentSkillResult> {
  const parsed = parseSkillDocument(args.markdown);
  if (!parsed) {
    throw new Error("Invalid SKILL.md document: expected frontmatter with name and description.");
  }

  if (args.expectedName !== undefined) {
    const normalizedExpected = normalizeSkillName(args.expectedName);
    if (parsed.frontmatter.name.toLowerCase() !== normalizedExpected.toLowerCase()) {
      throw new Error(
        `Skill name mismatch: expected "${normalizedExpected}" but markdown declares "${parsed.frontmatter.name}".`,
      );
    }
  }

  const location = getExternalSkillPath(parsed.frontmatter.name);
  await args.workspace.writeTextFile(location, args.markdown, "text/markdown");

  const normalizedName = parsed.frontmatter.name.toLowerCase();
  const duplicates = (await loadAllManagedExternalAgentSkillDefinitions({ workspace: args.workspace })).filter((skill) => {
    return skill.name.toLowerCase() === normalizedName && skill.location !== location;
  });

  for (const duplicate of duplicates) {
    await args.workspace.deleteFile(duplicate.location);
  }

  return {
    name: parsed.frontmatter.name,
    location,
  };
}

export async function upsertExternalAgentSkill(args: {
  markdown: string;
  expectedName?: string;
}): Promise<UpsertExternalAgentSkillResult> {
  return upsertExternalAgentSkillInWorkspace({
    workspace: getFilesWorkspace(),
    markdown: args.markdown,
    expectedName: args.expectedName,
  });
}

export async function removeExternalAgentSkillFromWorkspace(args: {
  workspace: ExternalSkillWorkspace;
  name: string;
}): Promise<boolean> {
  const normalizedName = normalizeSkillName(args.name).toLowerCase();
  const matches = (await loadAllManagedExternalAgentSkillDefinitions({ workspace: args.workspace })).filter((skill) => {
    return skill.name.toLowerCase() === normalizedName;
  });

  if (matches.length === 0) {
    return false;
  }

  const uniqueLocations = Array.from(new Set(matches.map((skill) => skill.location)));
  for (const location of uniqueLocations) {
    await args.workspace.deleteFile(location);
  }

  return true;
}

export async function removeExternalAgentSkill(args: {
  name: string;
}): Promise<boolean> {
  return removeExternalAgentSkillFromWorkspace({
    workspace: getFilesWorkspace(),
    name: args.name,
  });
}
