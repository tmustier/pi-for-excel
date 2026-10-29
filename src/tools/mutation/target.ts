import { Type, type Static } from "typebox";
import { parseRangeRef } from "../../excel/helpers.js";

export const mutationTargetSchema = Type.Object({
  action: Type.Optional(Type.String()),
  sheet: Type.Optional(Type.String()),
  range: Type.Optional(Type.String()),
  start_cell: Type.Optional(Type.String()),
  output_start_cell: Type.Optional(Type.String()),
  source_range: Type.Optional(Type.String()),
  position: Type.Optional(Type.Union([Type.String(), Type.Number()])),
});

export function requireExplicitMutationTarget(tool: string, params: Static<typeof mutationTargetSchema>): void {
  const requireRange = (ref: string): void => {
    if (!parseRangeRef(ref).sheet?.trim()) {
      throw new Error("Name the sheet, e.g. Sheet1!B2. Mutating tools do not use the active sheet.");
    }
  };
  switch (tool) {
    case "write_cells":
      if (params.start_cell !== undefined) requireRange(params.start_cell);
      break;
    case "fill_formula":
    case "conditional_format":
    case "comments":
    case "python_transform_range":
      if (params.range !== undefined) requireRange(params.range);
      if (params.output_start_cell !== undefined) requireRange(params.output_start_cell);
      break;
    case "format_cells":
      if (params.range !== undefined) requireRange(params.range.split(/[;,]/u)[0] ?? "");
      break;
    case "modify_structure":
      if (params.action === "add_sheet") break;
      if (!params.sheet?.trim()) throw new Error("Name the target sheet explicitly in sheet.");
      break;
    case "view_settings":
      if (!params.sheet?.trim()) throw new Error("Name the target sheet explicitly in sheet.");
      break;
    case "charts":
      if (!params.sheet?.trim() && !(params.action === "create" && params.source_range && parseRangeRef(params.source_range).sheet?.trim())) {
        throw new Error("Name the target sheet explicitly in sheet (or the source range for chart creation).");
      }
      break;
  }
}
