/**
 * trace_dependencies — Return formula lineage for a cell.
 *
 * Supports tracing both:
 * - precedents (what feeds this cell)
 * - dependents (what this cell feeds)
 *
 * Uses Office.js direct APIs when available and falls back to formula parsing.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  cellAddress,
  excelRun,
  getDirectDependentsSafe,
  getDirectPrecedentsSafe,
  getRange,
  parseCell,
  parseRangeRef,
  qualifiedAddress,
} from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import type {
  DepNodeDetail,
  TraceDependenciesDetails,
  TraceDependenciesMode,
  TraceDependencySource,
} from "./tool-details.js";
import {
  extractFormulaReferences,
  normalizeTraceMode,
  normalizeTraversalAddress,
  parseQualifiedCellAddress,
  parsedReferencesContainTarget,
  summarizeTraceTree,
  type ParsedFormulaReference,
} from "./trace-dependencies-logic.js";

const schema = Type.Object({
  cell: Type.String({
    description: 'Cell to trace, e.g. "D10", "Sheet2!F5". Must be a single cell, not a range.',
  }),
  mode: Type.Optional(
    Type.Union([
      Type.Literal("precedents"),
      Type.Literal("dependents"),
    ], {
      description: "Trace direction: precedents (upstream) or dependents (downstream). Default: precedents.",
    }),
  ),
  depth: Type.Optional(
    Type.Number({
      description: "How many levels of dependencies to trace. Default: 2. Max: 5.",
    }),
  ),
});

type Params = Static<typeof schema>;

interface FormulaDependentCandidate {
  dependentAddress: string;
  references: ParsedFormulaReference[];
}

interface TraceBuildState {
  usedApi: boolean;
  usedFormulaScan: boolean;
  truncated: boolean;
  dependentFormulaCandidates: FormulaDependentCandidate[] | null;
}

const MAX_DEPTH = 5;
const MAX_PRECEDENT_FALLBACK_REFS = 20;
const MAX_CHILDREN_PER_NODE = 80;
const MAX_DEPENDENT_SCAN_FORMULA_CELLS = 50_000;

function resolveTraceSource(state: TraceBuildState): TraceDependencySource {
  if (state.usedApi && state.usedFormulaScan) return "mixed";
  if (state.usedApi) return "api";
  if (state.usedFormulaScan) return "formula_scan";
  return "none";
}

async function loadLeafNode(
  context: Excel.RequestContext,
  cellRef: string,
): Promise<DepNodeDetail> {
  const { sheet, range } = getRange(context, cellRef);
  range.load("values,formulas,address,numberFormat");
  sheet.load("name");
  await context.sync();

  const rawFmt: unknown = range.numberFormat[0][0];
  const rawFormula: unknown = range.formulas[0][0];

  return {
    address: qualifiedAddress(sheet.name, range.address),
    value: range.values[0][0],
    numberFormat: typeof rawFmt === "string" && rawFmt !== "" ? rawFmt : undefined,
    formula: typeof rawFormula === "string" && rawFormula.startsWith("=") ? rawFormula : undefined,
    precedents: [],
  };
}

function estimateAddressCellCount(address: string): number {
  const parsed = parseRangeRef(address);
  const areas = parsed.address.split(",").map((part) => part.trim()).filter((part) => part.length > 0);

  let total = 0;
  for (const area of areas) {
    const [rawStart, rawEnd] = area.split(":");
    if (!rawStart) continue;

    const startToken = rawStart.replace(/\$/gu, "").trim();
    const endToken = (rawEnd ?? rawStart).replace(/\$/gu, "").trim();
    if (!startToken || !endToken) continue;

    try {
      const start = parseCell(startToken);
      const end = parseCell(endToken);
      const width = Math.abs(end.col - start.col) + 1;
      const height = Math.abs(end.row - start.row) + 1;
      total += width * height;
    } catch {
      // Skip malformed area segment.
    }
  }

  return total;
}

async function buildDependentFormulaCandidates(
  context: Excel.RequestContext,
  state: TraceBuildState,
): Promise<FormulaDependentCandidate[]> {
  const worksheets = context.workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();

  const rangesToInspect: Array<{ sheetName: string; range: Excel.Range }> = [];
  for (const worksheet of worksheets.items) {
    const usedRange = worksheet.getUsedRangeOrNullObject();
    usedRange.load("isNullObject,address");
    rangesToInspect.push({ sheetName: worksheet.name, range: usedRange });
  }
  await context.sync();

  let remainingCellBudget = MAX_DEPENDENT_SCAN_FORMULA_CELLS;
  const loadedRanges: Array<{
    sheetName: string;
    range: Excel.Range;
    startCol: number;
    startRow: number;
  }> = [];

  for (const candidate of rangesToInspect) {
    if (candidate.range.isNullObject) continue;

    const cellCount = estimateAddressCellCount(candidate.range.address);
    if (cellCount <= 0) continue;

    if (cellCount > remainingCellBudget) {
      state.truncated = true;
      continue;
    }

    const parsedRange = parseRangeRef(candidate.range.address);
    const startToken = parsedRange.address.split(":")[0]?.replace(/\$/gu, "").trim();
    if (!startToken) continue;

    let startCellRef: { col: number; row: number };
    try {
      startCellRef = parseCell(startToken);
    } catch {
      continue;
    }

    candidate.range.load("formulas");
    loadedRanges.push({
      sheetName: candidate.sheetName,
      range: candidate.range,
      startCol: startCellRef.col,
      startRow: startCellRef.row,
    });
    remainingCellBudget -= cellCount;
  }

  if (loadedRanges.length > 0) {
    await context.sync();
  }

  const formulaCandidates: FormulaDependentCandidate[] = [];

  for (const loaded of loadedRanges) {
    const formulasGrid: unknown = loaded.range.formulas;
    if (!Array.isArray(formulasGrid)) continue;

    for (const [rowIndex, rowValue] of formulasGrid.entries()) {
      if (!Array.isArray(rowValue)) continue;

      for (const [colIndex, formulaValue] of rowValue.entries()) {
        if (typeof formulaValue !== "string" || !formulaValue.startsWith("=")) {
          continue;
        }

        const references = extractFormulaReferences(formulaValue, loaded.sheetName);
        if (references.length === 0) continue;

        formulaCandidates.push({
          dependentAddress: qualifiedAddress(
            loaded.sheetName,
            cellAddress(loaded.startCol + colIndex, loaded.startRow + rowIndex),
          ),
          references,
        });
      }
    }
  }

  return formulaCandidates;
}

async function findDirectDependentsByFormulaScan(
  context: Excel.RequestContext,
  targetAddress: string,
  state: TraceBuildState,
): Promise<string[]> {
  const target = parseQualifiedCellAddress(targetAddress, "");
  if (!target) return [];

  if (!state.dependentFormulaCandidates) {
    state.dependentFormulaCandidates = await buildDependentFormulaCandidates(context, state);
  }

  const dependents = new Set<string>();
  for (const candidate of state.dependentFormulaCandidates) {
    if (!parsedReferencesContainTarget(candidate.references, target)) {
      continue;
    }

    dependents.add(candidate.dependentAddress);
    if (dependents.size >= MAX_CHILDREN_PER_NODE) {
      state.truncated = true;
      break;
    }
  }

  return [...dependents];
}

async function resolveChildAddresses(
  context: Excel.RequestContext,
  range: Excel.Range,
  sheetName: string,
  fullAddress: string,
  formula: string | undefined,
  mode: TraceDependenciesMode,
  state: TraceBuildState,
): Promise<string[]> {
  const children = new Set<string>();

  const addChild = (address: string | null): void => {
    if (!address) return;
    if (children.size >= MAX_CHILDREN_PER_NODE) {
      state.truncated = true;
      return;
    }
    children.add(address);
  };

  if (mode === "precedents") {
    const precedents = await getDirectPrecedentsSafe(context, range);
    if (precedents !== null) {
      state.usedApi = true;
      for (const group of precedents) {
        for (const address of group) {
          addChild(normalizeTraversalAddress(address, sheetName));
          if (children.size >= MAX_CHILDREN_PER_NODE) {
            return [...children];
          }
        }
      }
      return [...children];
    }

    if (!formula) {
      return [];
    }

    state.usedFormulaScan = true;
    const refs = extractFormulaReferences(formula, sheetName);
    if (refs.length > MAX_PRECEDENT_FALLBACK_REFS) {
      state.truncated = true;
    }

    for (const ref of refs.slice(0, MAX_PRECEDENT_FALLBACK_REFS)) {
      addChild(ref.startAddress);
    }

    return [...children];
  }

  const dependents = await getDirectDependentsSafe(context, range);
  if (dependents !== null) {
    state.usedApi = true;
    for (const group of dependents) {
      for (const address of group) {
        addChild(normalizeTraversalAddress(address, sheetName));
        if (children.size >= MAX_CHILDREN_PER_NODE) {
          return [...children];
        }
      }
    }
    return [...children];
  }

  state.usedFormulaScan = true;
  const scannedDependents = await findDirectDependentsByFormulaScan(context, fullAddress, state);
  for (const address of scannedDependents) {
    addChild(address);
  }

  return [...children];
}

async function traceCell(
  context: Excel.RequestContext,
  cellRef: string,
  maxDepth: number,
  currentDepth: number,
  visited: Set<string>,
  mode: TraceDependenciesMode,
  state: TraceBuildState,
): Promise<DepNodeDetail | null> {
  const { sheet, range } = getRange(context, cellRef);
  range.load("values,formulas,address,numberFormat");
  sheet.load("name");
  await context.sync();

  const fullAddr = qualifiedAddress(sheet.name, range.address);
  const rawFmt: unknown = range.numberFormat[0][0];
  const numberFormat = typeof rawFmt === "string" && rawFmt !== "" ? rawFmt : undefined;

  if (visited.has(fullAddr)) {
    return {
      address: fullAddr,
      value: range.values[0][0],
      numberFormat,
      formula: "(circular reference — already visited)",
      precedents: [],
    };
  }
  visited.add(fullAddr);

  const rawFormula: unknown = range.formulas[0][0];
  const value: unknown = range.values[0][0];
  const formula = typeof rawFormula === "string" && rawFormula.startsWith("=")
    ? rawFormula
    : undefined;

  // Precedent tracing requires a formula at the current node.
  if (mode === "precedents" && !formula) {
    return null;
  }

  const node: DepNodeDetail = {
    address: fullAddr,
    value,
    numberFormat,
    formula,
    precedents: [],
  };

  if (currentDepth >= maxDepth) return node;

  const childAddresses = await resolveChildAddresses(
    context,
    range,
    sheet.name,
    fullAddr,
    formula,
    mode,
    state,
  );

  for (const childAddress of childAddresses) {
    const child = await traceCell(
      context,
      childAddress,
      maxDepth,
      currentDepth + 1,
      visited,
      mode,
      state,
    );

    if (child) {
      node.precedents.push(child);
      continue;
    }

    const leaf = await loadLeafNode(context, childAddress);
    node.precedents.push(leaf);
  }

  return node;
}

function renderTree(node: DepNodeDetail, lines: string[], prefix: string, isLast: boolean): void {
  const connector = isLast ? "└── " : "├── ";
  const rawVal = node.value;
  const valueStr = rawVal !== "" && rawVal !== null && rawVal !== undefined
    ? ` = ${typeof rawVal === "string" || typeof rawVal === "number" || typeof rawVal === "boolean" ? String(rawVal) : JSON.stringify(rawVal)}`
    : "";
  const formulaStr = node.formula ? ` (${node.formula})` : "";

  lines.push(`${prefix}${connector}**${node.address}**${valueStr}${formulaStr}`);

  const childPrefix = prefix + (isLast ? "    " : "│   ");
  for (const [index, child] of node.precedents.entries()) {
    renderTree(child, lines, childPrefix, index === node.precedents.length - 1);
  }
}

export function createTraceDependenciesTool(): AgentTool<typeof schema> {
  return {
    name: "trace_dependencies",
    label: "Trace Dependencies",
    description:
      "Trace formula lineage for a cell. Supports precedents (inputs) and dependents " +
      "(downstream impact), recursively up to the specified depth.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<TraceDependenciesDetails | undefined>> => {
      try {
        if (params.cell.includes(":")) {
          return {
            content: [{ type: "text", text: "Error: trace_dependencies expects a single cell, not a range." }],
            details: undefined,
          };
        }

        const mode = normalizeTraceMode(params.mode);
        const maxDepth = Math.min(params.depth ?? 2, MAX_DEPTH);
        const traceState: TraceBuildState = {
          usedApi: false,
          usedFormulaScan: false,
          truncated: false,
          dependentFormulaCandidates: null,
        };

        const tree = await excelRun(async (context) => {
          return traceCell(context, params.cell, maxDepth, 0, new Set(), mode, traceState);
        });

        if (!tree) {
          return {
            content: [{ type: "text", text: `${params.cell} has no formula — it's a direct value or empty.` }],
            details: undefined,
          };
        }

        const heading = mode === "dependents" ? "Dependents" : "Precedents";
        const lines: string[] = [`**${heading} tree for ${tree.address}:**`, ""];
        renderTree(tree, lines, "", true);

        if (mode === "dependents" && tree.precedents.length === 0) {
          lines.push("", "_No direct dependents found._");
        }

        if (traceState.truncated) {
          lines.push("", "_Trace output was truncated to keep the result responsive._");
        }

        const summary = summarizeTraceTree(tree);

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: {
            kind: "trace_dependencies",
            root: tree,
            mode,
            maxDepth,
            nodeCount: summary.nodeCount,
            edgeCount: summary.edgeCount,
            source: resolveTraceSource(traceState),
            truncated: traceState.truncated,
          },
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error tracing dependencies: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}
