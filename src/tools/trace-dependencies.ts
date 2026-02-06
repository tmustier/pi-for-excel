/**
 * trace_dependencies — Return the formula dependency tree for a cell.
 *
 * This is Claude's biggest gap: it traces formulas manually, one cell at a time,
 * requiring dozens of tool calls for deep trees. We use getDirectPrecedents()
 * to walk the tree in a single tool call.
 *
 * Falls back to formula string parsing if getDirectPrecedents() fails.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, getRange, qualifiedAddress, getDirectPrecedentsSafe } from "../excel/helpers.js";

const schema = Type.Object({
  cell: Type.String({
    description: 'Cell to trace, e.g. "D10", "Sheet2!F5". Must be a single cell, not a range.',
  }),
  depth: Type.Optional(
    Type.Number({
      description: "How many levels of dependencies to trace. Default: 2. Max: 5.",
    }),
  ),
});

type Params = Static<typeof schema>;

interface DepNode {
  address: string;
  value: any;
  formula?: string;
  precedents: DepNode[];
}

export function createTraceDependenciesTool(): AgentTool<typeof schema> {
  return {
    name: "trace_dependencies",
    label: "Trace Dependencies",
    description:
      "Trace the formula dependency tree for a cell. Shows what cells feed into " +
      "the target cell, recursively up to the specified depth. " +
      "Useful for understanding how a value is calculated and finding the root inputs.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        if (params.cell.includes(":")) {
          return {
            content: [{ type: "text", text: "Error: trace_dependencies expects a single cell, not a range." }],
            details: undefined,
          };
        }

        const maxDepth = Math.min(params.depth || 2, 5);

        const tree = await excelRun(async (context) => {
          return await traceCell(context, params.cell, maxDepth, 0, new Set());
        });

        if (!tree) {
          return {
            content: [{ type: "text", text: `${params.cell} has no formula — it's a direct value or empty.` }],
            details: undefined,
          };
        }

        const lines: string[] = [];
        lines.push(`**Dependency tree for ${tree.address}:**`);
        lines.push("");
        renderTree(tree, lines, "", true);

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error tracing dependencies: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}

async function traceCell(
  context: any,
  cellRef: string,
  maxDepth: number,
  currentDepth: number,
  visited: Set<string>,
): Promise<DepNode | null> {
  const { sheet, range } = getRange(context, cellRef);
  range.load("values,formulas,address");
  sheet.load("name");
  await context.sync();

  const fullAddr = qualifiedAddress(sheet.name, range.address);

  // Avoid cycles
  if (visited.has(fullAddr)) {
    return {
      address: fullAddr,
      value: range.values[0][0],
      formula: "(circular reference — already visited)",
      precedents: [],
    };
  }
  visited.add(fullAddr);

  const formula = range.formulas[0][0];
  const value = range.values[0][0];

  // Not a formula — leaf node
  if (typeof formula !== "string" || !formula.startsWith("=")) {
    return null;
  }

  const node: DepNode = {
    address: fullAddr,
    value,
    formula,
    precedents: [],
  };

  // Don't recurse if we've hit the depth limit
  if (currentDepth >= maxDepth) return node;

  // Try getDirectPrecedents API
  const precedentAddrs = await getDirectPrecedentsSafe(context, range);

  if (precedentAddrs && precedentAddrs.length > 0) {
    // precedentAddrs is an array of address arrays
    for (const addrGroup of precedentAddrs) {
      for (const addr of addrGroup) {
        // Each addr could be "Sheet1!A1:A10" — handle ranges by taking first cell
        const singleCell = addr.includes(":") ? addr.split(":")[0] : addr;
        const child = await traceCell(context, singleCell, maxDepth, currentDepth + 1, visited);
        if (child) {
          node.precedents.push(child);
        } else {
          // Leaf value — still show it
          const { sheet: childSheet, range: childRange } = getRange(context, singleCell);
          childRange.load("values,address");
          childSheet.load("name");
          await context.sync();
          node.precedents.push({
            address: qualifiedAddress(childSheet.name, childRange.address),
            value: childRange.values[0][0],
            precedents: [],
          });
        }
      }
    }
  } else {
    // Fallback: parse formula string for cell references
    const refs = parseFormulaRefs(formula, sheet.name);
    for (const ref of refs.slice(0, 10)) { // limit to 10 refs per cell
      try {
        const child = await traceCell(context, ref, maxDepth, currentDepth + 1, visited);
        if (child) {
          node.precedents.push(child);
        } else {
          const { sheet: childSheet, range: childRange } = getRange(context, ref);
          childRange.load("values,address");
          childSheet.load("name");
          await context.sync();
          node.precedents.push({
            address: qualifiedAddress(childSheet.name, childRange.address),
            value: childRange.values[0][0],
            precedents: [],
          });
        }
      } catch {
        // Skip invalid references
      }
    }
  }

  return node;
}

/** Extract cell references from a formula string */
function parseFormulaRefs(formula: string, currentSheet: string): string[] {
  // Match patterns like: Sheet1!A1, 'Sheet Name'!B2, A1, $A$1, A1:B5
  const refPattern = /(?:'[^']+'|[A-Za-z_]\w*)!\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/g;

  const refs: string[] = [];
  const seen = new Set<string>();
  let match;

  while ((match = refPattern.exec(formula)) !== null) {
    let ref = match[0];
    // Remove $ signs
    ref = ref.replace(/\$/g, "");
    // Add current sheet if no sheet specified and it's a cell ref (not a function name)
    if (!ref.includes("!") && /^[A-Z]+\d+/.test(ref)) {
      ref = `${currentSheet}!${ref}`;
    }
    // Take just the start of a range
    if (ref.includes(":")) {
      ref = ref.split(":")[0];
    }
    const key = ref.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      refs.push(ref);
    }
  }

  return refs;
}

/** Render the dependency tree as an indented text tree */
function renderTree(node: DepNode, lines: string[], prefix: string, isLast: boolean): void {
  const connector = isLast ? "└── " : "├── ";
  const valueStr = node.value !== "" && node.value !== null && node.value !== undefined
    ? ` = ${node.value}`
    : "";
  const formulaStr = node.formula ? ` (${node.formula})` : "";

  lines.push(`${prefix}${connector}**${node.address}**${valueStr}${formulaStr}`);

  const childPrefix = prefix + (isLast ? "    " : "│   ");
  for (let i = 0; i < node.precedents.length; i++) {
    renderTree(node.precedents[i], lines, childPrefix, i === node.precedents.length - 1);
  }
}
