/**
 * search_workbook — Search for text, values, or formulas across the workbook.
 *
 * Supports substring and formula search modes.
 * Returns matching cells with their sheet, address, value, and formula.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { excelRun, qualifiedAddress, parseCell, colToLetter } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  query: Type.String({
    description: 'Search term. For formula search, use references like "Sheet1!" to find cross-sheet links.',
  }),
  search_formulas: Type.Optional(
    Type.Boolean({
      description:
        "If true, search in formula text instead of values. " +
        'Useful for finding cross-sheet references (e.g. query "Inputs!" to find all cells referencing Inputs sheet).',
    }),
  ),
  use_regex: Type.Optional(
    Type.Boolean({
      description: "If true, treat the query as a regular expression (case-insensitive).",
    }),
  ),
  offset: Type.Optional(
    Type.Number({
      description: "Skip the first N matches (pagination). Default: 0.",
    }),
  ),
  sheet: Type.Optional(
    Type.String({
      description: "Restrict search to this sheet. If omitted, searches all sheets.",
    }),
  ),
  max_results: Type.Optional(
    Type.Number({
      description: "Maximum number of results to return. Default: 20.",
    }),
  ),
  context_rows: Type.Optional(
    Type.Number({
      description:
        "Number of rows above and below each match to include as context. Default: 0 (no context). " +
        "Use 2-5 when searching for labels to see surrounding structure.",
    }),
  ),
});

type Params = Static<typeof schema>;

interface SearchMatch {
  sheet: string;
  address: string;
  value: DynamicValue;
  formula?: string;
  context?: string;
}

export function createSearchWorkbookTool(): AgentTool<typeof schema> {
  return {
    name: "search_workbook",
    label: "Search Workbook",
    description:
      "Search for text, values, or formulas across the workbook. " +
      "Returns matching cells with sheet name, address, value, and formula. " +
      "Use this to find specific data, locate cells by label, or trace cross-sheet references. " +
      "Set context_rows to see surrounding data for each match (useful for finding labeled cells and understanding their position).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const maxResults = Math.max(params.max_results || 20, 1);
        const offset = Math.max(params.offset || 0, 0);
        const searchFormulas = params.search_formulas || false;
        const useRegex = params.use_regex || false;
        const contextRows = Math.min(Math.max(params.context_rows ?? 0, 0), 10);
        const query = params.query;
        const queryLower = query.toLowerCase();

        let regex: RegExp | undefined;
        if (useRegex) {
          try {
            regex = new RegExp(query, "i");
          } catch (e) {
            return {
              content: [{ type: "text", text: `Invalid regex "${query}": ${getErrorMessage(e)}` }],
              details: undefined,
            };
          }
        }

        const result = await excelRun(async (context) => {
          const allMatches: SearchMatch[] = [];
          let totalMatches = 0;
          let hasMore = false;
          const sheets = context.workbook.worksheets;
          sheets.load("items/name,items/visibility");
          await context.sync();

          const targetSheets = params.sheet
            ? sheets.items.filter((s) => s.name === params.sheet)
            : sheets.items.filter((s) => s.visibility === "Visible");

          outer: for (const sheet of targetSheets) {
            const used = sheet.getUsedRangeOrNullObject();
            used.load("values,formulas,address");
            await context.sync();

            if (used.isNullObject) continue;

            const values = used.values;
            const formulas = used.formulas;

            // Parse start address for cell computation
            const addr = used.address;
            const bangIndex = addr.indexOf("!");
            const cellPart = bangIndex >= 0 ? addr.slice(bangIndex + 1) : addr;
            const colonIndex = cellPart.indexOf(":");
            const startCell = colonIndex >= 0 ? cellPart.slice(0, colonIndex) : cellPart;
            let start;
            try {
              start = parseCell(startCell);
            } catch {
              continue;
            }

            for (let r = 0; r < values.length; r++) {
              const valueRow = values[r] ?? [];
              const formulaRow = formulas[r] ?? [];
              for (let c = 0; c < valueRow.length; c++) {
                const value: DynamicValue = valueRow[c];
                const formula: DynamicValue = formulaRow[c];

                let match = false;
                if (searchFormulas) {
                  if (typeof formula !== "string" || formula.length === 0) continue;
                  const target = formula;
                  match = regex ? regex.test(target) : target.toLowerCase().includes(queryLower);
                } else {
                  if (value === null || value === undefined || value === "") continue;
                  const target = typeof value === "string" ? value : typeof value === "number" || typeof value === "boolean" ? String(value) : JSON.stringify(value);
                  match = regex ? regex.test(target) : target.toLowerCase().includes(queryLower);
                }

                if (match) {
                  totalMatches += 1;
                  if (totalMatches <= offset) continue;

                  const cellAddr = `${colToLetter(start.col + c)}${start.row + r}`;

                  const formulaText = typeof formula === "string" && formula.startsWith("=") ? formula : undefined;
                  const matchEntry: SearchMatch = {
                    sheet: sheet.name,
                    address: cellAddr,
                    value,
                    ...(formulaText !== undefined ? { formula: formulaText } : {}),
                  };
                  allMatches.push(matchEntry);

                  if (contextRows > 0) {
                    const rStart = Math.max(0, r - contextRows);
                    const rEnd = Math.min(values.length - 1, r + contextRows);
                    const colRadius = 10;
                    const cStart = Math.max(0, c - colRadius);
                    const firstValueRow = values[0] ?? [];
                    const cEnd = Math.min(firstValueRow.length - 1, c + colRadius);

                    const ctxLines: string[] = [];
                    const hdr: string[] = [""];
                    for (let ci = cStart; ci <= cEnd; ci++) {
                      hdr.push(colToLetter(start.col + ci));
                    }
                    ctxLines.push("| " + hdr.join(" | ") + " |");
                    ctxLines.push("|" + hdr.map(() => "---").join("|") + "|");

                    for (let ri = rStart; ri <= rEnd; ri++) {
                      const cells: string[] = [String(start.row + ri)];
                      const contextRow = values[ri] ?? [];
                      for (let ci = cStart; ci <= cEnd; ci++) {
                        const v: DynamicValue = contextRow[ci];
                        let s = v === null || v === undefined || v === "" ? "" : typeof v === "string" ? v : typeof v === "number" || typeof v === "boolean" ? String(v) : JSON.stringify(v);
                        if (s.length > 20) s = s.substring(0, 20) + "…";
                        s = s.replace(/\|/g, "\\|");
                        cells.push(s);
                      }
                      const marker = ri === r ? " ◀" : "";
                      ctxLines.push("| " + cells.join(" | ") + " |" + marker);
                    }

                    matchEntry.context = ctxLines
                      .map((l) => "  " + l)
                      .join("\n");
                  }

                  if (allMatches.length >= maxResults) {
                    hasMore = true;
                    break outer;
                  }
                }
              }
            }
          }
          return { matches: allMatches, hasMore, totalMatches };
        });

        const { matches, hasMore, totalMatches } = result;

        if (matches.length === 0) {
          const scope = params.sheet ? `in "${params.sheet}"` : "in any sheet";
          const mode = searchFormulas ? "formulas" : "values";
          const offsetNote = offset > 0 && totalMatches > 0
            ? ` after offset ${offset} (total matches: ${totalMatches})`
            : "";
          return {
            content: [{ type: "text", text: `No matches for "${params.query}" ${scope}${offsetNote} (searched ${mode}).` }],
            details: undefined,
          };
        }

        const lines: string[] = [];
        const limitNote = hasMore ? " (limit reached)" : "";
        const offsetNote = offset > 0 ? ` (offset ${offset})` : "";
        lines.push(`**${matches.length} match(es)** for "${params.query}"${limitNote}${offsetNote}:`);
        lines.push("");

        for (const m of matches) {
          const addr = qualifiedAddress(m.sheet, m.address);
          const val = typeof m.value === "string" && m.value.length > 60
            ? m.value.substring(0, 60) + "…"
            : String(m.value);
          const formulaStr = m.formula ? ` ← ${m.formula}` : "";
          lines.push(`- **${addr}**: ${val}${formulaStr}`);
          if (m.context) {
            lines.push(m.context);
          }
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error searching: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}
