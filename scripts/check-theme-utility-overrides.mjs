import { promises as fs } from "node:fs";
import path from "node:path";

const ROOT = process.cwd();
const THEME_ROOT = path.join(ROOT, "src", "ui", "theme");
const UNSTABLE_FILE = path.join(THEME_ROOT, "unstable-overrides.css");

const UTILITY_CLASS_PATTERN = /\.(?:p[trblxy]?|m[trblxy]?|gap|space-[xy]?|text|bg|border|rounded|shadow|flex|grid|inline|block|hidden|absolute|relative|fixed|sticky|top|right|bottom|left|min-|max-|w-|h-|overflow|cursor|hover|transition|duration|animate)-[a-z0-9_:[\]()./%-]+/i;

async function collectCssFiles(dirPath) {
  const entries = await fs.readdir(dirPath, { withFileTypes: true });
  const files = [];

  for (const entry of entries) {
    const fullPath = path.join(dirPath, entry.name);

    if (entry.isDirectory()) {
      files.push(...await collectCssFiles(fullPath));
      continue;
    }

    if (entry.isFile() && entry.name.endsWith(".css")) {
      files.push(fullPath);
    }
  }

  return files;
}

function relative(filePath) {
  return path.relative(ROOT, filePath);
}

function stripInlineBlockComments(line) {
  return line.replace(/\/\*.*?\*\//g, "");
}

function findViolations(source, filePath) {
  const lines = source.split("\n");
  const violations = [];
  let insideComment = false;

  for (let i = 0; i < lines.length; i += 1) {
    let line = lines[i];

    if (insideComment) {
      const endIndex = line.indexOf("*/");
      if (endIndex === -1) continue;
      line = line.slice(endIndex + 2);
      insideComment = false;
    }

    while (true) {
      const start = line.indexOf("/*");
      if (start === -1) break;

      const end = line.indexOf("*/", start + 2);
      if (end === -1) {
        line = line.slice(0, start);
        insideComment = true;
        break;
      }

      line = `${line.slice(0, start)}${line.slice(end + 2)}`;
    }

    if (insideComment) continue;

    line = stripInlineBlockComments(line);

    if (!UTILITY_CLASS_PATTERN.test(line)) continue;

    violations.push({
      filePath,
      lineNo: i + 1,
      snippet: line.trim(),
    });
  }

  return violations;
}

async function main() {
  const files = (await collectCssFiles(THEME_ROOT))
    .filter((filePath) => path.resolve(filePath) !== path.resolve(UNSTABLE_FILE))
    .sort();

  const violations = [];

  for (const filePath of files) {
    const source = await fs.readFile(filePath, "utf8");
    violations.push(...findViolations(source, filePath));
  }

  if (violations.length > 0) {
    console.error("\n✗ Utility-class selectors detected outside theme/unstable-overrides.css\n");

    for (const violation of violations) {
      console.error(`- ${relative(violation.filePath)}:${violation.lineNo}`);
      console.error(`  ${violation.snippet}`);
    }

    console.error("\nEither move the selector to src/ui/theme/unstable-overrides.css or add stable semantic hooks.\n");
    process.exitCode = 1;
    return;
  }

  console.log(`✓ Theme utility-selector guard passed (${files.length} files).`);
}

void main();
