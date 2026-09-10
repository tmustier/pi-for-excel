#!/usr/bin/env node
import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";
import ts from "typescript";

const root = process.cwd();
const offenders = [];

function collectFiles(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name === "dist") continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.isFile() && full.endsWith(".ts") && !full.endsWith(".d.ts")) out.push(full);
  }
  return out;
}

function position(sourceFile, node) {
  const pos = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
  return `${pos.line + 1}:${pos.character + 1}`;
}

for (const file of collectFiles(join(root, "src"))) {
  const text = readFileSync(file, "utf8");
  const sourceFile = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);
  const rel = relative(root, file);
  const exportedNames = new Set(
    sourceFile.statements.flatMap((statement) => {
      if (ts.isExportAssignment(statement) && ts.isIdentifier(statement.expression)) {
        return [statement.expression.text];
      }
      if (!ts.isExportDeclaration(statement) || statement.moduleSpecifier || !statement.exportClause) return [];
      if (!ts.isNamedExports(statement.exportClause)) return [];
      return statement.exportClause.elements.map((element) => (element.propertyName ?? element.name).text);
    }),
  );

  for (const statement of sourceFile.statements) {
    if (ts.isExportAssignment(statement)) {
      const expression = statement.expression;
      if ((ts.isArrowFunction(expression) || ts.isFunctionExpression(expression)) && !expression.type) {
        offenders.push(`${rel}:${position(sourceFile, expression)} <default>`);
      }
    }
    const directlyExported = statement.modifiers?.some(
      (modifier) => modifier.kind === ts.SyntaxKind.ExportKeyword,
    ) ?? false;

    if (ts.isFunctionDeclaration(statement) && statement.body && !statement.type) {
      const name = statement.name?.text ?? "<default>";
      if (directlyExported || exportedNames.has(name)) {
        offenders.push(`${rel}:${position(sourceFile, statement)} function ${name}`);
      }
      continue;
    }

    if (!ts.isVariableStatement(statement)) continue;
    for (const declaration of statement.declarationList.declarations) {
      if (!ts.isIdentifier(declaration.name) || declaration.type) continue;
      if (!directlyExported && !exportedNames.has(declaration.name.text)) continue;
      const initializer = declaration.initializer;
      if (!initializer) continue;
      if ((ts.isArrowFunction(initializer) || ts.isFunctionExpression(initializer)) && !initializer.type) {
        offenders.push(`${rel}:${position(sourceFile, declaration)} ${declaration.name.text}`);
      }
    }
  }
}

if (offenders.length > 0) {
  console.error("Top-level return type check failed.");
  console.error("Declare return types for exported top-level functions so future agents can read public contracts without re-inference.");
  for (const offender of offenders) console.error(`- ${offender}`);
  process.exit(1);
}

console.log("✓ Top-level return type check passed.");
