import { cellAddress, parseCell, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import type { DepNodeDetail, TraceDependenciesMode } from "./tool-details.js";

export interface ParsedCellAddress {
  sheet: string;
  col: number;
  row: number;
}

export interface ParsedFormulaReference {
  sheet: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
  startAddress: string;
}

const FORMULA_REF_PATTERN = /(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/gu;

export function normalizeTraceMode(mode: string | undefined): TraceDependenciesMode {
  return mode === "dependents" ? "dependents" : "precedents";
}

function normalizeSheetKey(sheetName: string): string {
  return sheetName.trim().toLowerCase();
}

function stripQuotedStringLiterals(formula: string): string {
  return formula.replace(/"(?:[^"]|"")*"/gu, "");
}

export function parseQualifiedCellAddress(cellRef: string, defaultSheet: string): ParsedCellAddress | null {
  try {
    const parsed = parseRangeRef(cellRef);
    const sheetName = parsed.sheet ?? defaultSheet;
    if (!sheetName) return null;

    const firstArea = parsed.address.split(",")[0]?.trim();
    if (!firstArea) return null;

    const firstCell = firstArea.split(":")[0]?.replace(/\$/gu, "").trim();
    if (!firstCell) return null;

    const { col, row } = parseCell(firstCell);
    return { sheet: sheetName, col, row };
  } catch {
    return null;
  }
}

export function normalizeTraversalAddress(address: string, defaultSheet: string): string | null {
  const parsed = parseQualifiedCellAddress(address, defaultSheet);
  if (!parsed) return null;
  return qualifiedAddress(parsed.sheet, cellAddress(parsed.col, parsed.row));
}

export function extractFormulaReferences(formula: string, currentSheet: string): ParsedFormulaReference[] {
  const references: ParsedFormulaReference[] = [];
  const seen = new Set<string>();
  const searchFormula = stripQuotedStringLiterals(formula);

  for (const match of searchFormula.matchAll(FORMULA_REF_PATTERN)) {
    const token = match[0];
    if (!token) continue;

    let sheetName = currentSheet;
    let addressPart = token;

    if (token.includes("!")) {
      const parsed = parseRangeRef(token);
      if (parsed.sheet) {
        sheetName = parsed.sheet;
      }
      addressPart = parsed.address;
    }

    const normalized = addressPart.replace(/\$/gu, "");
    const [rawStart, rawEnd] = normalized.split(":");
    if (!rawStart) continue;

    const startToken = rawStart.trim();
    const endToken = (rawEnd ?? rawStart).trim();
    if (!startToken || !endToken) continue;

    try {
      const start = parseCell(startToken);
      const end = parseCell(endToken);

      const startCol = Math.min(start.col, end.col);
      const endCol = Math.max(start.col, end.col);
      const startRow = Math.min(start.row, end.row);
      const endRow = Math.max(start.row, end.row);

      const key = [normalizeSheetKey(sheetName), startCol, startRow, endCol, endRow].join("|");
      if (seen.has(key)) continue;
      seen.add(key);

      references.push({
        sheet: sheetName,
        startCol,
        startRow,
        endCol,
        endRow,
        startAddress: qualifiedAddress(sheetName, cellAddress(startCol, startRow)),
      });
    } catch {
      // Skip malformed reference tokens.
    }
  }

  return references;
}

export function parsedReferencesContainTarget(
  references: readonly ParsedFormulaReference[],
  target: ParsedCellAddress,
): boolean {
  const targetSheetKey = normalizeSheetKey(target.sheet);

  return references.some((ref) => {
    if (normalizeSheetKey(ref.sheet) !== targetSheetKey) return false;

    return (
      target.col >= ref.startCol &&
      target.col <= ref.endCol &&
      target.row >= ref.startRow &&
      target.row <= ref.endRow
    );
  });
}

export function formulaReferencesTargetCell(
  formula: string,
  currentSheet: string,
  targetAddress: string,
): boolean {
  const target = parseQualifiedCellAddress(targetAddress, currentSheet);
  if (!target) return false;

  const references = extractFormulaReferences(formula, currentSheet);
  return parsedReferencesContainTarget(references, target);
}

export function summarizeTraceTree(root: DepNodeDetail): { nodeCount: number; edgeCount: number } {
  let nodeCount = 0;
  let edgeCount = 0;
  const stack: DepNodeDetail[] = [root];

  while (stack.length > 0) {
    const node = stack.pop();
    if (!node) continue;

    nodeCount += 1;
    edgeCount += node.precedents.length;
    for (const child of node.precedents) {
      stack.push(child);
    }
  }

  return { nodeCount, edgeCount };
}
