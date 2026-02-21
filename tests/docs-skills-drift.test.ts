import assert from "node:assert/strict";
import { readFile, readdir, stat } from "node:fs/promises";
import { test } from "node:test";

import { listBuiltinWorkspaceDocs } from "../src/files/builtin-docs.ts";
import { listAgentSkills } from "../src/skills/catalog.ts";

const EXCLUDED_DOC_PREFIXES: readonly string[] = [
  "docs/archive/",
  "docs/release-smoke-runs/",
];

function toPosixPath(value: string): string {
  return value.replaceAll("\\", "/");
}

async function collectMarkdownFiles(dirUrl: URL, repoPrefix: string): Promise<string[]> {
  const entries = await readdir(dirUrl, { withFileTypes: true });
  const discovered: string[] = [];

  for (const entry of entries) {
    if (entry.isDirectory()) {
      const nested = await collectMarkdownFiles(new URL(`${entry.name}/`, dirUrl), `${repoPrefix}/${entry.name}`);
      discovered.push(...nested);
      continue;
    }

    if (!entry.isFile() || !entry.name.endsWith(".md")) {
      continue;
    }

    discovered.push(`${repoPrefix}/${entry.name}`);
  }

  discovered.sort((left, right) => left.localeCompare(right));
  return discovered;
}

function isBundledDocPath(path: string): boolean {
  if (path === "README.md") {
    return true;
  }

  if (!path.startsWith("docs/")) {
    return false;
  }

  return !EXCLUDED_DOC_PREFIXES.some((prefix) => path.startsWith(prefix));
}

async function readBundledDocPathsFromFilesystem(): Promise<string[]> {
  const docs = await collectMarkdownFiles(new URL("../docs/", import.meta.url), "docs");
  const filteredDocs = docs.filter((path) => isBundledDocPath(path));

  return ["README.md", ...filteredDocs]
    .sort((left, right) => left.localeCompare(right));
}

async function readBundledSkillPathsFromFilesystem(): Promise<string[]> {
  const skillsRoot = new URL("../skills/", import.meta.url);
  const entries = await readdir(skillsRoot, { withFileTypes: true });
  const discovered: string[] = [];

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }

    const skillPath = `skills/${entry.name}/SKILL.md`;
    const skillUrl = new URL(`../${skillPath}`, import.meta.url);

    try {
      const skillStat = await stat(skillUrl);
      if (skillStat.isFile()) {
        discovered.push(skillPath);
      }
    } catch {
      // ignore folders without a top-level SKILL.md
    }
  }

  discovered.sort((left, right) => left.localeCompare(right));
  return discovered;
}

async function readRepoText(relativePath: string): Promise<string> {
  return readFile(new URL(`../${relativePath}`, import.meta.url), "utf8");
}

void test("builtin docs bundle stays in sync with non-archive docs", async () => {
  const expectedDocPaths = await readBundledDocPathsFromFilesystem();
  const expectedWorkspacePaths = expectedDocPaths.map((path) => `assistant-docs/${path}`);

  const actualWorkspacePaths = listBuiltinWorkspaceDocs()
    .map((entry) => toPosixPath(entry.path))
    .sort((left, right) => left.localeCompare(right));

  assert.deepEqual(actualWorkspacePaths, expectedWorkspacePaths);
});

void test("bundled skills catalog stays in sync with skills/*/SKILL.md", async () => {
  const expectedSkillPaths = await readBundledSkillPathsFromFilesystem();
  const actualSkillPaths = listAgentSkills()
    .map((entry) => toPosixPath(entry.location))
    .sort((left, right) => left.localeCompare(right));

  assert.deepEqual(actualSkillPaths, expectedSkillPaths);
});

void test("public docs avoid retired /integrations command references", async () => {
  const docPaths = await readBundledDocPathsFromFilesystem();
  const skillPaths = await readBundledSkillPathsFromFilesystem();
  const filesToCheck = [...docPaths, ...skillPaths];

  const staleCommandRefs: string[] = [];

  for (const filePath of filesToCheck) {
    const source = await readRepoText(filePath);
    if (source.includes("`/integrations`")) {
      staleCommandRefs.push(filePath);
    }
  }

  assert.deepEqual(staleCommandRefs, []);
});

void test("web-search docs keep Jina as the default provider", async () => {
  const webSearchSkill = await readRepoText("skills/web-search/SKILL.md");

  assert.match(webSearchSkill, /Jina \(default\)/i);
  assert.doesNotMatch(webSearchSkill, /Serper\.dev \(default\)/i);
});
