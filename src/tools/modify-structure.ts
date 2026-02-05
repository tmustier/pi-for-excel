/**
 * modify_structure — Insert/delete rows, columns, and sheets.
 *
 * Combines Claude's modify_sheet_structure + modify_workbook_structure
 * into a single tool.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun } from "../excel/helpers.js";

// Helper for string enum (TypeBox doesn't have a built-in StringEnum)
function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((v) => Type.Literal(v)),
    opts,
  );
}

const schema = Type.Object({
  action: StringEnum(
    [
      "insert_rows",
      "delete_rows",
      "insert_columns",
      "delete_columns",
      "add_sheet",
      "delete_sheet",
      "rename_sheet",
      "duplicate_sheet",
      "hide_sheet",
      "unhide_sheet",
    ],
    { description: "The structural modification to perform." },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Target sheet name. Required for sheet operations and row/column operations on a specific sheet. " +
        "If omitted for row/column ops, uses the active sheet.",
    }),
  ),
  position: Type.Optional(
    Type.Number({
      description:
        "For insert_rows/delete_rows: the 1-indexed row number. " +
        "For insert_columns/delete_columns: the 1-indexed column number. " +
        "For add_sheet: the 0-indexed position to insert the new sheet.",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description: "Number of rows or columns to insert/delete. Default: 1.",
    }),
  ),
  new_name: Type.Optional(
    Type.String({
      description: 'New name for rename_sheet or add_sheet. Also used for duplicate_sheet target name.',
    }),
  ),
});

type Params = Static<typeof schema>;

export function createModifyStructureTool(): AgentTool<typeof schema> {
  return {
    name: "modify_structure",
    label: "Modify Structure",
    description:
      "Modify the workbook structure: insert/delete rows and columns, " +
      "add/delete/rename/duplicate/hide/unhide sheets. " +
      "Be careful with deletions — there is no undo.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await excelRun(async (context: any) => {
          const action = params.action;
          const count = params.count || 1;

          const getSheet = () => {
            if (params.sheet) {
              return context.workbook.worksheets.getItem(params.sheet);
            }
            return context.workbook.worksheets.getActiveWorksheet();
          };

          switch (action) {
            case "insert_rows": {
              if (!params.position) throw new Error("position is required for insert_rows");
              const sheet = getSheet();
              const range = sheet.getRange(`${params.position}:${params.position + count - 1}`);
              range.insert("Down");
              await context.sync();
              sheet.load("name");
              await context.sync();
              return `Inserted ${count} row(s) at row ${params.position} in "${sheet.name}".`;
            }

            case "delete_rows": {
              if (!params.position) throw new Error("position is required for delete_rows");
              const sheet = getSheet();
              const range = sheet.getRange(`${params.position}:${params.position + count - 1}`);
              range.delete("Up");
              await context.sync();
              sheet.load("name");
              await context.sync();
              return `Deleted ${count} row(s) starting at row ${params.position} in "${sheet.name}".`;
            }

            case "insert_columns": {
              if (!params.position) throw new Error("position is required for insert_columns");
              const sheet = getSheet();
              // Convert column number to letter
              let col = params.position - 1; // 0-indexed
              let letter = "";
              while (col >= 0) {
                letter = String.fromCharCode((col % 26) + 65) + letter;
                col = Math.floor(col / 26) - 1;
              }
              const range = sheet.getRange(`${letter}:${letter}`);
              for (let i = 0; i < count; i++) {
                range.insert("Right");
              }
              await context.sync();
              sheet.load("name");
              await context.sync();
              return `Inserted ${count} column(s) at column ${params.position} (${letter}) in "${sheet.name}".`;
            }

            case "delete_columns": {
              if (!params.position) throw new Error("position is required for delete_columns");
              const sheet = getSheet();
              let col = params.position - 1;
              let startLetter = "";
              let temp = col;
              while (temp >= 0) {
                startLetter = String.fromCharCode((temp % 26) + 65) + startLetter;
                temp = Math.floor(temp / 26) - 1;
              }
              let endLetter = "";
              temp = col + count - 1;
              while (temp >= 0) {
                endLetter = String.fromCharCode((temp % 26) + 65) + endLetter;
                temp = Math.floor(temp / 26) - 1;
              }
              const range = sheet.getRange(`${startLetter}:${endLetter}`);
              range.delete("Left");
              await context.sync();
              sheet.load("name");
              await context.sync();
              return `Deleted ${count} column(s) starting at column ${params.position} (${startLetter}) in "${sheet.name}".`;
            }

            case "add_sheet": {
              const name = params.new_name || `Sheet${Date.now()}`;
              const newSheet = context.workbook.worksheets.add(name);
              if (params.position !== undefined) {
                newSheet.position = params.position;
              }
              await context.sync();
              return `Added sheet "${name}".`;
            }

            case "delete_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for delete_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.delete();
              await context.sync();
              return `Deleted sheet "${params.sheet}".`;
            }

            case "rename_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for rename_sheet");
              if (!params.new_name) throw new Error("new_name is required for rename_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.name = params.new_name;
              await context.sync();
              return `Renamed sheet "${params.sheet}" to "${params.new_name}".`;
            }

            case "duplicate_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for duplicate_sheet");
              const source = context.workbook.worksheets.getItem(params.sheet);
              const copy = source.copy("End");
              await context.sync();
              if (params.new_name) {
                copy.load("name");
                await context.sync();
                copy.name = params.new_name;
                await context.sync();
                return `Duplicated "${params.sheet}" as "${params.new_name}".`;
              }
              copy.load("name");
              await context.sync();
              return `Duplicated "${params.sheet}" as "${copy.name}".`;
            }

            case "hide_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for hide_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.visibility = "Hidden";
              await context.sync();
              return `Hidden sheet "${params.sheet}".`;
            }

            case "unhide_sheet": {
              if (!params.sheet) throw new Error("sheet name is required for unhide_sheet");
              const sheet = context.workbook.worksheets.getItem(params.sheet);
              sheet.visibility = "Visible";
              await context.sync();
              return `Unhidden sheet "${params.sheet}".`;
            }

            default:
              throw new Error(`Unknown action: ${action}`);
          }
        });

        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e: any) {
        return {
          content: [{ type: "text", text: `Error: ${e.message}` }],
          details: undefined,
        };
      }
    },
  };
}
