import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import { excelRun, parseRangeRef, parseCell, computeRangeAddress, qualifiedAddress } from "../excel/helpers.js";
import { createRecoveryScopeResolver } from "../workbook/recovery-scope.js";
import { CORE_TOOL_NAMES } from "../tools/names.js";
import type { WorkbookMutationCall } from "../tools/mutation/gate.js";
import { captureFormatCellsState } from "../workbook/recovery/format-state.js";
import { captureConditionalFormatState } from "../workbook/recovery/conditional-format-state.js";
import { captureCommentThreadState } from "../workbook/recovery/comment-state.js";
import { captureChartPresentState } from "../workbook/recovery/chart-state.js";
import { localAddressPart, splitRangeList } from "../workbook/recovery/address.js";
import type { MutationHost, MutationObservation, MutationPlan } from "./mutations.js";

const gridSchema = Type.Array(Type.Array(Type.Union([Type.String(), Type.Number(), Type.Boolean(), Type.Null()])));
const paramsSchema = Type.Object({
  action: Type.Optional(Type.String()), sheet: Type.Optional(Type.String()),
  range: Type.Optional(Type.String()), start_cell: Type.Optional(Type.String()),
  output_start_cell: Type.Optional(Type.String()), source_range: Type.Optional(Type.String()),
  name: Type.Optional(Type.String()), new_name: Type.Optional(Type.String()),
  values: Type.Optional(gridSchema), content: Type.Optional(Type.String()),
});
const targetSchema = Type.Object({
  kind: Type.Union([Type.Literal("cells"), Type.Literal("format"), Type.Literal("conditional"),
    Type.Literal("comment"), Type.Literal("chart"), Type.Literal("sheet"), Type.Literal("opaque")]),
  sheetId: Type.Union([Type.String(), Type.Null()]),
  ranges: Type.Array(Type.String()),
  chartId: Type.Union([Type.String(), Type.Null()]),
  action: Type.String(),
});
type Target = Static<typeof targetSchema>;
type Params = Static<typeof paramsSchema>;

const identitiesSchema = Type.Array(Type.Object({
  id: Type.String(), name: Type.String(), position: Type.Number(), visibility: Type.String(), active: Type.Boolean(),
  tables: Type.Array(Type.Object({ id: Type.String(), name: Type.String(), range: Type.String(),
    columns: Type.Array(Type.Object({ id: Type.Number(), name: Type.String() })),
  })),
}));
type SheetIdentity = Static<typeof identitiesSchema>[number];

function json<T>(value: T): JsonValue {
  const decoded: unknown = JSON.parse(JSON.stringify(value));
  if (!isJsonValue(decoded)) throw new Error("Workbook observation is not JSON serializable");
  return decoded;
}

function content(values: Static<typeof gridSchema>, formulas: Static<typeof gridSchema>): JsonValue {
  if (values.length !== formulas.length || values.some((row, i) => row.length !== formulas[i]?.length)) {
    throw new Error("Workbook returned mismatched cell grids");
  }
  return values.map((row, r) => row.map((value, c) => {
    const formula = formulas[r]?.[c];
    return typeof formula === "string" && formula.startsWith("=") ? formula : value ?? "";
  }));
}

/** Read IDs and table extents; no name-based fallback on reconciliation. */
async function structure(): Promise<SheetIdentity[]> {
  return excelRun(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/id,items/name,items/position,items/visibility");
    const active = sheets.getActiveWorksheet();
    active.load("id");
    await context.sync();
    for (const sheet of sheets.items) sheet.tables.load("items/id,items/name");
    await context.sync();
    const entries = sheets.items.map((sheet) => ({ sheet, tables: sheet.tables.items.map((table) => {
      const range = table.getRange();
      range.load("address");
      table.columns.load("items/id,items/name");
      return { table, range };
    }) }));
    await context.sync();
    return entries.map(({ sheet, tables }) => ({
      id: sheet.id, name: sheet.name, position: sheet.position, visibility: String(sheet.visibility), active: sheet.id === active.id,
      tables: tables.map(({ table, range }) => ({
        id: table.id, name: table.name, range: localAddressPart(range.address),
        columns: table.columns.items.map((column) => ({ id: column.id, name: column.name })),
      })).sort((a, b) => a.id.localeCompare(b.id)),
    })).sort((a, b) => a.id.localeCompare(b.id));
  });
}

async function targetFor(toolName: string, params: Params): Promise<Target> {
  const target: Target = { kind: "opaque", sheetId: null, ranges: [], chartId: null, action: params.action ?? "" };
  let ref = params.output_start_cell ?? params.range ?? params.start_cell ?? params.source_range;
  switch (toolName) {
    case "write_cells": target.kind = "cells"; break;
    case "fill_formula": target.kind = "cells"; break;
    case "python_transform_range": target.kind = "cells"; break;
    case "format_cells": target.kind = "format"; break;
    case "conditional_format": target.kind = "conditional"; break;
    case "comments": target.kind = "comment"; break;
    case "charts": target.kind = "chart"; break;
    case "modify_structure":
    case "view_settings": target.kind = "sheet"; break;
    default: return target;
  }
  if (toolName === "write_cells" && params.start_cell && params.values?.length) {
    const columns = params.values.reduce((max, row) => Math.max(max, row.length), 0);
    if (columns > 0) {
      const parsed = parseRangeRef(params.start_cell);
      const local = computeRangeAddress(parsed.address, params.values.length, columns);
      ref = parsed.sheet ? qualifiedAddress(parsed.sheet, local) : local;
    }
  }
  const refs = ref ? splitRangeList(ref).map(parseRangeRef) : [];
  const sheetName = params.sheet ?? refs.find((part) => part.sheet)?.sheet;
  if (refs.some((part) => part.sheet && sheetName && part.sheet !== sheetName)) {
    throw new Error("Mutation observation requires unambiguous sheet references");
  }
  target.ranges = refs.map((part) => part.address);
  await excelRun(async (context) => {
    const sheet = sheetName ? context.workbook.worksheets.getItemOrNullObject(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    sheet.load("id,isNullObject");
    await context.sync();
    if (sheet.isNullObject) return;
    target.sheetId = sheet.id;
    if (target.kind === "chart" && params.name && params.action !== "create") {
      const chart = sheet.charts.getItemOrNullObject(params.name);
      chart.load("id,isNullObject");
      await context.sync();
      if (!chart.isNullObject) target.chartId = chart.id;
    }
  });
  return target;
}

const allFormats = {
  numberFormat: true, fillColor: true, fontColor: true, bold: true, italic: true,
  underlineStyle: true, fontName: true, fontSize: true, horizontalAlignment: true,
  verticalAlignment: true, wrapText: true, columnWidth: true, rowHeight: true,
  mergedAreas: true, borderTop: true, borderBottom: true, borderLeft: true, borderRight: true,
  borderInsideHorizontal: true, borderInsideVertical: true,
};

async function cellAnchors(sheetId: string, addresses: string[]): Promise<JsonValue> {
  const bounds = addresses.map((address) => {
    const [first, last] = address.split(":");
    return { first: parseCell(first ?? ""), last: parseCell(last ?? first ?? "") };
  });
  return excelRun(async (context) => {
    const used = context.workbook.worksheets.getItem(sheetId).getUsedRangeOrNullObject(true);
    used.load("isNullObject");
    await context.sync();
    if (used.isNullObject) return [];
    used.load("rowIndex,columnIndex,values,formulas");
    await context.sync();
    const values: unknown = used.values;
    const formulas: unknown = used.formulas;
    if (!Value.Check(gridSchema, values) || !Value.Check(gridSchema, formulas)) throw new Error("Unsupported anchor data");
    const anchors: JsonValue[] = [];
    for (const [r, row] of values.entries()) {
      for (const [c, value] of row.entries()) {
        const rowIndex = used.rowIndex + r + 1;
        const columnIndex = used.columnIndex + c;
        if (bounds.some((bound) => rowIndex >= bound.first.row && rowIndex <= bound.last.row
          && columnIndex >= bound.first.col && columnIndex <= bound.last.col)) continue;
        const formula = formulas[r]?.[c];
        const cell = typeof formula === "string" && formula.startsWith("=") ? formula : value;
        if (cell !== null && cell !== "") anchors.push([rowIndex, columnIndex, cell]);
      }
    }
    return anchors;
  });
}

async function observeTarget(workbookId: string, target: Target, toolName: string): Promise<MutationObservation> {
  const identities = await structure();
  const sheet = identities.find((entry) => entry.id === target.sheetId);
  const base: MutationObservation = {
    workbookId, structure: json(identities), state: null, complete: false,
    limitation: "The host cannot prove all effects of this operation. Explicit review is required.",
  };
  if (target.kind === "opaque") {
    // Include named items for review, but never mistake them or cell contents for a complete JS effect set.
    base.state = await excelRun(async (context) => {
      const names = context.workbook.names;
      names.load("items/name,items/formula,items/type");
      await context.sync();
      return names.items.map((name) => {
        const formula: unknown = name.formula;
        if (!isJsonValue(formula)) throw new Error("Unsupported named-item formula");
        return { name: name.name, formula, type: String(name.type) };
      });
    });
    return base;
  }
  if (!sheet) {
    return { ...base, state: { missingSheetId: target.sheetId }, limitation: "Recorded sheet identity is absent; names are not substitutes." };
  }
  const addresses = target.ranges.map((range) => qualifiedAddress(sheet.name, range));
  if (target.kind === "cells") {
    const state = await excelRun(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheet.id);
      const ranges = target.ranges.map((address) => worksheet.getRange(address));
      for (const range of ranges) range.load("values,formulas,address,rowCount,columnCount");
      await context.sync();
      return ranges.map((range) => {
        const values: unknown = range.values;
        const formulas: unknown = range.formulas;
        if (!Value.Check(gridSchema, values) || !Value.Check(gridSchema, formulas)) {
          throw new Error("Unsupported workbook cell values");
        }
        return { range: localAddressPart(range.address), content: content(values, formulas) };
      });
    });
    // AutoFill also copies formatting; Python's output shape is determined only after running external code.
    return { ...base, state, structure: { sheets: json(identities), anchors: await cellAnchors(sheet.id, target.ranges) },
      complete: toolName === "write_cells" && state.length > 0,
      limitation: toolName === "write_cells" ? null : "AutoFill formatting or dynamic Python output extends beyond the observed cell contents." };
  }
  if (target.kind === "format") {
    const captures = [];
    for (const address of addresses) captures.push(await captureFormatCellsState(address, allFormats));
    return { ...base, state: json(captures), limitation: "Format snapshots may omit mixed properties and merge-induced data loss; inspect the workbook." };
  }
  if (target.kind === "conditional") {
    const captures = [];
    for (const address of addresses) captures.push(await captureConditionalFormatState(address));
    return { ...base, state: json(captures), limitation: "Conditional-format snapshots do not prove rule identity, priority or every host-specific property." };
  }
  if (target.kind === "comment") {
    const captures = [];
    for (const address of addresses) captures.push(await captureCommentThreadState(address));
    return { ...base, state: json(captures), limitation: "Comment text does not prove thread/reply identity, authorship or notifications." };
  }
  if (target.kind === "chart") {
    const charts = await excelRun(async (context) => {
      const collection = context.workbook.worksheets.getItem(sheet.id).charts;
      collection.load("items/id,items/name");
      await context.sync();
      return collection.items.map((chart) => ({ id: chart.id, name: chart.name }));
    });
    const states = [];
    for (const chart of charts) {
      if (target.chartId === null || chart.id === target.chartId) {
        states.push({ id: chart.id, state: await captureChartPresentState(chart.name, sheet.name) });
      }
    }
    return { ...base, state: json({ charts, states }), limitation: "Chart source data and some axis properties cannot be fully observed by the recovery reader." };
  }
  if (toolName === "view_settings") {
    const state = await excelRun(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheet.id);
      worksheet.load("showGridlines,showHeadings,tabColor,standardWidth");
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("id");
      const frozen = worksheet.freezePanes.getLocationOrNullObject();
      frozen.load("isNullObject,address");
      await context.sync();
      return { showGridlines: worksheet.showGridlines, showHeadings: worksheet.showHeadings,
        tabColor: worksheet.tabColor, standardWidth: worksheet.standardWidth,
        activeSheetId: active.id, frozen: frozen.isNullObject ? null : localAddressPart(frozen.address) };
    });
    return { ...base, state, complete: true, limitation: null };
  }
  if (target.action === "rename_sheet" || target.action === "hide_sheet" || target.action === "unhide_sheet") {
    return { ...base, complete: true, limitation: null };
  }
  // Structural insertion/deletion has no persistent row/column identity in Office.js.
  // Capture all used value/formula data without the undo log's size/retention limits.
  const state = await excelRun(async (context) => {
    const range = context.workbook.worksheets.getItem(sheet.id).getUsedRangeOrNullObject(true);
    range.load("isNullObject");
    await context.sync();
    if (range.isNullObject) return null;
    range.load("address,values,formulas");
    await context.sync();
    const values: unknown = range.values;
    const formulas: unknown = range.formulas;
    if (!Value.Check(gridSchema, values) || !Value.Check(gridSchema, formulas)) throw new Error("Unsupported structural cell state");
    return { range: localAddressPart(range.address), content: content(values, formulas) };
  });
  return { ...base, state, limitation: "Rows/columns lack stable identities; matching value data cannot prove insertion, deletion or copy did not occur." };
}

/** Production read-only Office adapter; WPS/generic execution retains explicit uncertainty. */
export function createWorkbookMutationHost(options: {
  kind: "office" | "wps" | "browser";
  getWorkbookId?: () => Promise<string | null>;
}): MutationHost {
  const scope = createRecoveryScopeResolver();
  const getId = options.getWorkbookId ?? (async () => (await scope.resolveForRead())?.workbookId ?? null);
  const observe = async (target: Target, toolName: string): Promise<MutationObservation> => {
    const workbookId = await getId();
    if (!workbookId) throw new Error("Stable workbook identity is unavailable");
    let observation: MutationObservation;
    try {
      observation = options.kind === "office" ? await observeTarget(workbookId, target, toolName) : {
        workbookId, structure: null, state: null, complete: false,
        limitation: "This host has no stable-identity mutation observer; inspect the workbook and explicitly accept uncertainty.",
      };
    } catch {
      observation = { workbookId, structure: null, state: null, complete: false,
        limitation: "Workbook observation failed or an API is unsupported. No outcome has been proved." };
    }
    if (await getId() !== workbookId) throw new Error("Workbook identity changed during observation");
    return observation;
  };
  return {
    prepare: async (call: WorkbookMutationCall): Promise<MutationPlan> => {
      const structured = CORE_TOOL_NAMES.some((name) => name === call.toolName);
      const params = Value.Check(paramsSchema, call.params) ? call.params : {};
      let target: Target = { kind: "opaque", sheetId: null, chartId: null, ranges: [], action: "" };
      if (options.kind === "office" && (structured || call.toolName === "python_transform_range")) {
        try { target = await targetFor(call.toolName, params); } catch { /* Observation remains explicitly incomplete. */ }
      }
      const before = await observe(target, call.toolName);
      let expected: MutationObservation | null = null;
      if (call.toolName === "write_cells" && before.complete && params.values?.length && target.ranges[0]) {
        const width = params.values.reduce((max, row) => Math.max(max, row.length), 0);
        const padded = params.values.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ""));
        expected = { ...before, state: [{ range: target.ranges[0], content: padded }] };
      }
      if (call.toolName === "modify_structure" && before.complete && Value.Check(identitiesSchema, before.structure)) {
        const identities = structuredClone(before.structure);
        const sheet = identities.find((entry) => entry.id === target.sheetId);
        if (sheet) {
          if (target.action === "rename_sheet" && params.new_name) sheet.name = params.new_name;
          if (target.action === "hide_sheet") sheet.visibility = "Hidden";
          if (target.action === "unhide_sheet") sheet.visibility = "Visible";
          expected = { ...before, structure: identities };
        }
      }
      return { target, before, expected,
        intent: structured ? call.params : { toolName: call.toolName, scope: "unbounded", argumentsOmitted: true } };
    },
    observe: async (record) => {
      if (!Value.Check(targetSchema, record.target)) throw new Error("Invalid persisted mutation target");
      return observe(record.target, record.toolName);
    },
  };
}
