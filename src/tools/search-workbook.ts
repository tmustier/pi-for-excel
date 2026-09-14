/**
 * search_workbook — Search for text, values, or formulas across the workbook.
 *
 * Supports substring and formula search modes.
 * Returns matching cells with their sheet, address, value, and formula.
 */

import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import {
  colToLetter,
  excelRun,
  parseCell,
  parseRangeRef,
  qualifiedAddress,
} from "../excel/helpers.js";
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
type SearchCellValue = string | number | boolean | null;
type LoadedSearchCell = SearchCellValue | undefined;

const searchCellMatrixSchema = Type.Refine(
  Type.Unsafe<SearchCellValue[][]>({}),
  (value) => Array.isArray(value) && value.every(Array.isArray),
  () => "search cell matrix must be an array of rows",
);

interface SearchMatch {
  sheet: string;
  address: string;
  value: LoadedSearchCell;
  formula?: string;
  context?: string;
}

interface SearchRangeOptions {
  sheetName: string;
  values: SearchCellValue[][];
  formulas: SearchCellValue[][];
  start: { col: number; row: number };
  searchFormulas: boolean;
  queryLower: string;
  regex?: RegExp;
  contextRows: number;
  offset: number;
  matchesBeforeRange: number;
  resultCapacity: number;
}

interface SearchRangeResult {
  matches: SearchMatch[];
  totalMatches: number;
  hasMore: boolean;
}

function searchableValue(value: LoadedSearchCell): string | null {
  if (value === null || value === undefined || value === "") return null;
  return typeof value === "string" ? value : String(value);
}

function contextCellText(value: LoadedSearchCell): string {
  const text = searchableValue(value) ?? "";
  const bounded = text.length > 20 ? `${text.substring(0, 20)}…` : text;
  return bounded.replace(/\\/g, "\\\\").replace(/\|/g, "\\|");
}

function buildContextPreview(
  values: SearchCellValue[][],
  start: { col: number; row: number },
  matchRow: number,
  matchColumn: number,
  contextRows: number,
): string {
  const rowStart = Math.max(0, matchRow - contextRows);
  const rowEnd = Math.min(values.length - 1, matchRow + contextRows);
  const columnStart = Math.max(0, matchColumn - 10);
  const firstValueRow = values[0] ?? [];
  const columnEnd = Math.min(firstValueRow.length - 1, matchColumn + 10);

  const lines: string[] = [];
  const header = [""];
  for (let column = columnStart; column <= columnEnd; column += 1) {
    header.push(colToLetter(start.col + column));
  }
  lines.push(`| ${header.join(" | ")} |`);
  lines.push(`|${header.map(() => "---").join("|")}|`);

  for (let row = rowStart; row <= rowEnd; row += 1) {
    const cells = [String(start.row + row)];
    const contextRow = values[row] ?? [];
    for (let column = columnStart; column <= columnEnd; column += 1) {
      cells.push(contextCellText(contextRow[column]));
    }
    const marker = row === matchRow ? " ◀" : "";
    lines.push(`| ${cells.join(" | ")} |${marker}`);
  }

  return lines.map((line) => `  ${line}`).join("\n");
}

function searchLoadedRange(options: SearchRangeOptions): SearchRangeResult {
  const matches: SearchMatch[] = [];
  let totalMatches = options.matchesBeforeRange;

  for (let row = 0; row < options.values.length; row += 1) {
    const valueRow = options.values[row] ?? [];
    const formulaRow = options.formulas[row] ?? [];

    for (let column = 0; column < valueRow.length; column += 1) {
      const value = valueRow[column];
      const formula = formulaRow[column];
      const target = options.searchFormulas
        ? (typeof formula === "string" && formula.startsWith("=") ? formula : null)
        : searchableValue(value);
      if (target === null) continue;
      const matchesQuery = options.regex
        ? options.regex.test(target)
        : target.toLowerCase().includes(options.queryLower);
      if (!matchesQuery) continue;

      totalMatches += 1;
      if (totalMatches <= options.offset) continue;
      if (matches.length >= options.resultCapacity) {
        return { matches, totalMatches, hasMore: true };
      }

      const formulaText = typeof formula === "string" && formula.startsWith("=") ? formula : undefined;
      const match: SearchMatch = {
        sheet: options.sheetName,
        address: `${colToLetter(options.start.col + column)}${options.start.row + row}`,
        value,
        ...(formulaText !== undefined ? { formula: formulaText } : {}),
      };

      if (options.contextRows > 0) {
        match.context = buildContextPreview(
          options.values,
          options.start,
          row,
          column,
          options.contextRows,
        );
      }

      matches.push(match);
    }
  }

  return { matches, totalMatches, hasMore: false };
}

function renderSearchMatches(matches: SearchMatch[], query: string, hasMore: boolean, offset: number): string {
  const lines: string[] = [];
  const limitNote = hasMore ? " (limit reached)" : "";
  const offsetNote = offset > 0 ? ` (offset ${offset})` : "";
  lines.push(`**${matches.length} match(es)** for "${query}"${limitNote}${offsetNote}:`);
  lines.push("");

  for (const match of matches) {
    const address = qualifiedAddress(match.sheet, match.address);
    const value = typeof match.value === "string" && match.value.length > 60
      ? `${match.value.substring(0, 60)}…`
      : String(match.value);
    const formula = match.formula ? ` ← ${match.formula}` : "";
    lines.push(`- **${address}**: ${value}${formula}`);
    if (match.context) lines.push(match.context);
  }

  return lines.join("\n");
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

          for (const sheet of targetSheets) {
            const used = sheet.getUsedRangeOrNullObject();
            used.load("values,formulas,address");
            await context.sync();

            if (used.isNullObject) continue;

            const values = Value.Parse(searchCellMatrixSchema, used.values);
            const formulas = Value.Parse(searchCellMatrixSchema, used.formulas);

            const cellPart = parseRangeRef(used.address).address;
            const colonIndex = cellPart.indexOf(":");
            const startCell = colonIndex >= 0 ? cellPart.slice(0, colonIndex) : cellPart;
            let start;
            try {
              start = parseCell(startCell);
            } catch {
              continue;
            }

            const rangeResult = searchLoadedRange({
              sheetName: sheet.name,
              values,
              formulas,
              start,
              searchFormulas,
              queryLower,
              ...(regex !== undefined ? { regex } : {}),
              contextRows,
              offset,
              matchesBeforeRange: totalMatches,
              resultCapacity: maxResults - allMatches.length,
            });
            allMatches.push(...rangeResult.matches);
            totalMatches = rangeResult.totalMatches;
            if (rangeResult.hasMore) {
              hasMore = true;
              break;
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

        return {
          content: [{ type: "text", text: renderSearchMatches(matches, params.query, hasMore, offset) }],
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
