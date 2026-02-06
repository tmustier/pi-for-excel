/**
 * search_workbook — Search for text, values, or formulas across the workbook.
 *
 * Supports substring and formula search modes.
 * Returns matching cells with their sheet, address, value, and formula.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
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
});

type Params = Static<typeof schema>;

interface SearchMatch {
  sheet: string;
  address: string;
  value: unknown;
  formula?: string;
}

export function createSearchWorkbookTool(): AgentTool<typeof schema> {
  return {
    name: "search_workbook",
    label: "Search Workbook",
    description:
      "Search for text, values, or formulas across the workbook. " +
      "Returns matching cells with sheet name, address, value, and formula. " +
      "Use this to find specific data, locate cells by label, or trace cross-sheet references.",
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
        const query = params.query;
        const queryLower = query.toLowerCase();

        let regex: RegExp | undefined;
        if (useRegex) {
          try {
            regex = new RegExp(query, "i");
          } catch (e: unknown) {
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
            const cellPart = addr.includes("!") ? addr.split("!")[1] : addr;
            const startCell = cellPart.split(":")[0];
            let start;
            try {
              start = parseCell(startCell);
            } catch {
              continue;
            }

            for (let r = 0; r < values.length; r++) {
              for (let c = 0; c < values[r].length; c++) {
                const value = values[r][c];
                const formula = formulas[r][c];

                let match = false;
                if (searchFormulas) {
                  if (typeof formula !== "string" || formula.length === 0) continue;
                  const target = formula;
                  match = regex ? regex.test(target) : target.toLowerCase().includes(queryLower);
                } else {
                  if (value === null || value === undefined || value === "") continue;
                  const target = String(value);
                  match = regex ? regex.test(target) : target.toLowerCase().includes(queryLower);
                }

                if (match) {
                  totalMatches += 1;
                  if (totalMatches <= offset) continue;

                  const cellAddr = `${colToLetter(start.col + c)}${start.row + r}`;

                  allMatches.push({
                    sheet: sheet.name,
                    address: cellAddr,
                    value,
                    formula: typeof formula === "string" && formula.startsWith("=") ? formula : undefined,
                  });

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
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details: undefined,
        };
      } catch (e: unknown) {
        return {
          content: [{ type: "text", text: `Error searching: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}
