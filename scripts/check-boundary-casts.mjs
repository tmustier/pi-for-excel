#!/usr/bin/env node
import { readdirSync, readFileSync } from "node:fs";
import { join, relative } from "node:path";
import ts from "typescript";

const root = process.cwd();
const searchRoots = ["src", "tests"];
const allowedBoundaryTypes = new Set(["DynamicValue"]);
const offenders = [];

function collectFiles(dir, out = []) {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === "node_modules" || entry.name === "dist") continue;
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectFiles(full, out);
    else if (entry.isFile() && full.endsWith(".ts")) out.push(full);
  }
  return out;
}

function unwrapExpression(node) {
  let current = node;
  while (ts.isParenthesizedExpression(current)) current = current.expression;
  return current;
}

function isJsonParseCall(node) {
  const expr = unwrapExpression(node);
  return ts.isCallExpression(expr) &&
    ts.isPropertyAccessExpression(expr.expression) &&
    expr.expression.expression.getText() === "JSON" &&
    expr.expression.name.text === "parse";
}

function isResponseJsonCall(node) {
  const expr = unwrapExpression(node);
  const call = ts.isAwaitExpression(expr) ? unwrapExpression(expr.expression) : expr;
  return ts.isCallExpression(call) &&
    ts.isPropertyAccessExpression(call.expression) &&
    call.expression.name.text === "json";
}

function typeText(sourceFile, node) {
  return node.type.getText(sourceFile).replace(/\s+/g, " ").trim();
}

function addOffender(sourceFile, file, node, kind, castType) {
  const pos = sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile));
  offenders.push(`${relative(root, file)}:${pos.line + 1}:${pos.character + 1} ${kind} cast to ${castType}`);
}

for (const dir of searchRoots) {
  for (const file of collectFiles(join(root, dir))) {
    const text = readFileSync(file, "utf8");
    const sourceFile = ts.createSourceFile(file, text, ts.ScriptTarget.Latest, true, ts.ScriptKind.TS);

    function visit(node) {
      if (ts.isAsExpression(node)) {
        const castType = typeText(sourceFile, node);
        if (isJsonParseCall(node.expression) && !allowedBoundaryTypes.has(castType)) {
          addOffender(sourceFile, file, node, "JSON.parse", castType);
        }
        if (isResponseJsonCall(node.expression) && !allowedBoundaryTypes.has(castType)) {
          addOffender(sourceFile, file, node, "Response.json", castType);
        }
      }
      ts.forEachChild(node, visit);
    }

    visit(sourceFile);
  }
}

if (offenders.length > 0) {
  console.error("Boundary cast check failed.");
  console.error("Decoded JSON/fetch payloads must first land as DynamicValue, then be parsed/refined by a concrete boundary parser.");
  console.error("Offenders:");
  for (const offender of offenders) console.error(`- ${offender}`);
  process.exit(1);
}

console.log("✓ Boundary cast check passed.");
