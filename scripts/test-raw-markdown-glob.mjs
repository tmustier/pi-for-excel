import { readFileSync, readdirSync } from "node:fs";
import minimatch from "minimatch";
import path from "node:path";
import { fileURLToPath } from "node:url";

function toPosixPath(value) {
  return value.split(path.sep).join("/");
}

function readSearchRoot(absolutePattern) {
  const normalized = toPosixPath(absolutePattern);
  const segments = normalized.split("/");
  const wildcardIndex = segments.findIndex((segment) => segment.includes("*"));

  if (wildcardIndex === -1) {
    return null;
  }

  const rootSegments = segments.slice(0, wildcardIndex);
  if (rootSegments.length === 0) {
    return "/";
  }

  const root = rootSegments.join("/");
  return root.length === 0 ? "/" : root;
}

function walkFiles(rootDir) {
  const discovered = [];
  const stack = [rootDir];

  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }

    const entries = readdirSync(current, { withFileTypes: true });
    for (const entry of entries) {
      const absolutePath = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(absolutePath);
        continue;
      }

      if (entry.isFile()) {
        discovered.push(absolutePath);
      }
    }
  }

  discovered.sort((left, right) => left.localeCompare(right));
  return discovered;
}

function resolveGlobMatches(importerUrl, pattern) {
  const importerFilePath = fileURLToPath(importerUrl);
  const importerDir = path.dirname(importerFilePath);
  const absolutePattern = path.resolve(importerDir, pattern);
  const searchRoot = readSearchRoot(absolutePattern);

  if (searchRoot === null) {
    return [absolutePattern];
  }

  const patternPosix = toPosixPath(absolutePattern);
  const candidates = walkFiles(searchRoot);
  return candidates.filter((candidate) => minimatch(toPosixPath(candidate), patternPosix, { dot: true }));
}

function toGlobResultKey(importerUrl, absolutePath) {
  const importerFilePath = fileURLToPath(importerUrl);
  const importerDir = path.dirname(importerFilePath);
  const relative = toPosixPath(path.relative(importerDir, absolutePath));
  if (relative.startsWith(".")) {
    return relative;
  }

  return `./${relative}`;
}

export function createRawMarkdownGlobLoader() {
  return (pattern, importerUrl) => {
    const matches = resolveGlobMatches(importerUrl, pattern);
    const result = {};

    for (const match of matches) {
      const key = toGlobResultKey(importerUrl, match);
      result[key] = readFileSync(match, "utf8");
    }

    return result;
  };
}
