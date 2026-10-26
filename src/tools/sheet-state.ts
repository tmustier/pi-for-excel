/**
 * Sheet state the read tools surface to the agent: protection, merges,
 * hidden rows/columns, filters, tables, data validation, spills and the
 * calculation mode. Each is invisible in `range.values`, and without them the
 * agent falls back to execute_office_js to find out why an edit failed or
 * why a filtered table "loses" rows.
 *
 * Collection is bounded by the block it is given (one read_range page); every
 * list is capped so the notes stay a few lines.
 */

import {
  blockAddress,
  intersectBlocks,
  parseCell,
  parseRangeRef,
  type CellBlock,
} from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";

const MAX_LISTED_AREAS = 10;
const MAX_VALIDATION_AREAS = 12;
const MAX_SPILL_CHECKS = 100;

export interface TableState {
  name: string;
  address: string;
  filtered: boolean;
}

export interface ValidationState {
  address: string;
  rule: string;
}

export interface SpillState {
  anchor: string;
  address: string;
}

export interface RangeState {
  protected: boolean;
  /** Locked state of the block's cells: true, false, or null when mixed. Matters only when protected. */
  locked: boolean | null;
  calculationMode: string;
  /** 1-based row numbers inside the block that are hidden (filtered or hidden). */
  hiddenRows: number[];
  /** 0-based column indexes inside the block that are hidden. */
  hiddenColumns: number[];
  merged: string[];
  mergedMore: number;
  tables: TableState[];
  /** Sheet AutoFilter range when one is applied (tables have their own). */
  autoFilter?: { address: string; filtered: boolean };
  validation: ValidationState[];
  validationMore: number;
  spills: SpillState[];
  /** Set when part of the state could not be read; the rest is still reported. */
  unavailable?: string;
}

function stripSheet(address: string): string {
  return parseRangeRef(address).address;
}

/** Parse "A1:C2" / "C5" (no sheet) into a block. */
export function parseBlock(address: string): CellBlock | null {
  const [rawStart, rawEnd] = stripSheet(address).replace(/\$/gu, "").split(":");
  if (!rawStart) return null;
  try {
    const start = parseCell(rawStart);
    const end = parseCell(rawEnd ?? rawStart);
    return {
      rowIndex: Math.min(start.row, end.row) - 1,
      columnIndex: Math.min(start.col, end.col),
      rowCount: Math.abs(end.row - start.row) + 1,
      columnCount: Math.abs(end.col - start.col) + 1,
    };
  } catch {
    return null;
  }
}

/** Split a RangeAreas address ("S!A1:C2,S!A6:C6") into blocks. */
function parseAreas(address: string): CellBlock[] {
  return address
    .split(",")
    .map((part) => parseBlock(part.trim()))
    .filter((block): block is CellBlock => block !== null);
}

/** "3-5, 9, 12-14" from sorted numbers, listing at most `maxRuns` runs. */
export function formatNumberRuns(
  numbers: readonly number[],
  label: (n: number) => string = String,
  maxRuns = 20,
): string {
  const parts: string[] = [];
  let index = 0;
  while (index < numbers.length) {
    const first = numbers[index];
    if (first === undefined) break;
    let last = first;
    while (numbers[index + 1] === last + 1) {
      index += 1;
      last = last + 1;
    }
    parts.push(first === last ? label(first) : `${label(first)}-${label(last)}`);
    index += 1;
  }
  if (parts.length <= maxRuns) return parts.join(", ");
  return `${parts.slice(0, maxRuns).join(", ")}, … (+${parts.length - maxRuns} more runs, ${numbers.length} in total)`;
}

type BasicRule = Excel.BasicDataValidation | Excel.DateTimeDataValidation;

function ruleOperand(value: string | number | Date | Excel.Range | undefined): string {
  if (typeof value === "string" || typeof value === "number") return String(value);
  return value instanceof Date ? value.toISOString().slice(0, 10) : "";
}

function describeBasicRule(kind: string, rule: BasicRule): string {
  const operator = String(rule.operator);
  const first = ruleOperand(rule.formula1);
  const second = ruleOperand(rule.formula2);
  const operatorText = operator.replace(/([a-z])([A-Z])/gu, "$1 $2").toLowerCase();
  const bounds = operator === "Between" || operator === "NotBetween" ? `${first} and ${second}` : first;
  return `${kind} ${operatorText} ${bounds}`.trim();
}

/** Compact text for an Office.js data validation rule. */
export function describeValidationRule(type: string, rule: Excel.DataValidationRule | null | undefined): string {
  if (rule?.list) return `list: ${ruleOperand(rule.list.source)}`;
  if (rule?.custom) return `custom: ${ruleOperand(rule.custom.formula)}`;
  const basic: Array<[string, BasicRule | null | undefined]> = [
    ["whole number", rule?.wholeNumber],
    ["decimal", rule?.decimal],
    ["date", rule?.date],
    ["time", rule?.time],
    ["text length", rule?.textLength],
  ];
  for (const [kind, shape] of basic) {
    if (shape) return describeBasicRule(kind, shape);
  }
  return type.replace(/([a-z])([A-Z])/gu, "$1 $2").toLowerCase();
}

/** Visible cells → hidden rows/columns of the block. */
function hiddenFromVisible(block: CellBlock, visibleAddress: string | null): { rows: number[]; columns: number[] } {
  const visible = visibleAddress === null ? [] : parseAreas(visibleAddress);
  const rowVisible = new Set<number>();
  const columnVisible = new Set<number>();
  for (const area of visible) {
    for (let r = area.rowIndex; r < area.rowIndex + area.rowCount; r += 1) rowVisible.add(r);
    for (let c = area.columnIndex; c < area.columnIndex + area.columnCount; c += 1) columnVisible.add(c);
  }
  const rows: number[] = [];
  for (let r = block.rowIndex; r < block.rowIndex + block.rowCount; r += 1) {
    if (!rowVisible.has(r)) rows.push(r + 1);
  }
  const columns: number[] = [];
  for (let c = block.columnIndex; c < block.columnIndex + block.columnCount; c += 1) {
    if (!columnVisible.has(c)) columns.push(c);
  }
  // A fully hidden block reports its rows, not every column as well.
  return rows.length === block.rowCount ? { rows, columns: [] } : { rows, columns };
}

/**
 * Read the state of one bounded block. `spillCandidates` are formula cells
 * (A1, no sheet) to check for a spill range; at most 100 are checked.
 * Failures degrade to `unavailable` instead of failing the read.
 */
export async function collectRangeState(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  block: CellBlock,
  spillCandidates: readonly string[],
): Promise<RangeState> {
  const state: RangeState = {
    protected: false,
    locked: null,
    calculationMode: "Automatic",
    hiddenRows: [],
    hiddenColumns: [],
    merged: [],
    mergedMore: 0,
    tables: [],
    validation: [],
    validationMore: 0,
    spills: [],
  };
  const range = sheet.getRangeByIndexes(block.rowIndex, block.columnIndex, block.rowCount, block.columnCount);

  const problems: string[] = [];
  const protection = sheet.protection;
  protection.load("protected");
  const application = context.workbook.application;
  application.load("calculationMode");
  const cellProtection = range.format.protection;
  cellProtection.load("locked");

  try {
    const merged = range.getMergedAreasOrNullObject();
    merged.load("address,areaCount");
    const visible = range.getSpecialCellsOrNullObject("Visible");
    visible.load("address");
    const validation = range.getSpecialCellsOrNullObject("DataValidations");
    validation.load("address,areaCount");
    const tables = sheet.tables;
    tables.load("items/name");
    const autoFilter = sheet.autoFilter;
    autoFilter.load("enabled,isDataFiltered");
    const autoFilterRange = autoFilter.getRangeOrNullObject();
    autoFilterRange.load("address");
    await context.sync();

    state.protected = protection.protected;
    state.locked = cellProtection.locked;
    state.calculationMode = String(application.calculationMode);
    if (!merged.isNullObject) {
      const areas = merged.address.split(",").map((part) => stripSheet(part.trim()));
      state.merged = areas.slice(0, MAX_LISTED_AREAS);
      state.mergedMore = Math.max(0, merged.areaCount - state.merged.length);
    }
    const hidden = hiddenFromVisible(block, visible.isNullObject ? null : visible.address);
    state.hiddenRows = hidden.rows;
    state.hiddenColumns = hidden.columns;
    if (autoFilter.enabled && !autoFilterRange.isNullObject) {
      state.autoFilter = { address: stripSheet(autoFilterRange.address), filtered: autoFilter.isDataFiltered };
    }

    const tableRanges = tables.items.map((table) => {
      const tableRange = table.getRange();
      tableRange.load("address");
      const tableFilter = table.autoFilter;
      tableFilter.load("isDataFiltered");
      return { table, tableRange, tableFilter };
    });

    const validationAreas: Array<{ address: string; range: Excel.Range }> = [];
    if (!validation.isNullObject) {
      const parts = parseAreas(validation.address);
      state.validationMore = Math.max(0, parts.length - MAX_VALIDATION_AREAS);
      for (const part of parts.slice(0, MAX_VALIDATION_AREAS)) {
        validationAreas.push({
          address: blockAddress(part),
          range: sheet.getRangeByIndexes(part.rowIndex, part.columnIndex, part.rowCount, part.columnCount),
        });
      }
      for (const area of validationAreas) area.range.dataValidation.load("type,rule");
    }
    if (tableRanges.length > 0 || validationAreas.length > 0) await context.sync();

    for (const { table, tableRange, tableFilter } of tableRanges) {
      const tableBlock = parseBlock(tableRange.address);
      if (!tableBlock || !intersectBlocks(tableBlock, block)) continue;
      state.tables.push({ name: table.name, address: stripSheet(tableRange.address), filtered: tableFilter.isDataFiltered });
    }

    await collectValidation(context, sheet, validationAreas, state);
  } catch (error) {
    problems.push(getErrorMessage(error));
  }

  try {
    await collectSpills(context, sheet, spillCandidates, state);
  } catch (error) {
    problems.push(getErrorMessage(error));
  }

  if (problems.length > 0) state.unavailable = problems.join("; ");
  return state;
}

/**
 * Adjacent validated cells come back as one area even with different rules
 * (type "Inconsistent"); split those by column so each rule is reported.
 */
async function collectValidation(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  areas: Array<{ address: string; range: Excel.Range }>,
  state: RangeState,
): Promise<void> {
  const resolved: Array<{ address: string; range: Excel.Range }> = [];
  const split: Array<{ address: string; range: Excel.Range }> = [];
  for (const area of areas) {
    if (area.range.dataValidation.type !== "Inconsistent") {
      resolved.push(area);
      continue;
    }
    const block = parseBlock(area.address);
    if (!block) continue;
    for (let c = 0; c < block.columnCount && split.length < MAX_VALIDATION_AREAS; c += 1) {
      const column = { ...block, columnIndex: block.columnIndex + c, columnCount: 1 };
      const range = sheet.getRangeByIndexes(column.rowIndex, column.columnIndex, column.rowCount, 1);
      range.dataValidation.load("type,rule");
      split.push({ address: blockAddress(column), range });
    }
  }
  if (split.length > 0) await context.sync();

  const all = [...resolved, ...split].sort((a, b) => {
    const left = parseBlock(a.address);
    const right = parseBlock(b.address);
    if (!left || !right) return 0;
    return left.rowIndex - right.rowIndex || left.columnIndex - right.columnIndex;
  });
  for (const area of all) {
    const type = String(area.range.dataValidation.type);
    if (type === "None") continue;
    if (state.validation.length >= MAX_VALIDATION_AREAS) {
      state.validationMore += 1;
      continue;
    }
    const rule = area.range.dataValidation.rule;
    state.validation.push({ address: area.address, rule: type === "Inconsistent" ? "mixed rules" : describeValidationRule(type, rule) });
  }
}

/** Formulas that can return arrays: any function call, range or spill reference. */
export function canSpill(formula: string): boolean {
  return /[(:#]/u.test(formula);
}

async function collectSpills(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  spillCandidates: readonly string[],
  state: RangeState,
): Promise<void> {
  const candidates = spillCandidates.slice(0, MAX_SPILL_CHECKS);
  if (candidates.length === 0) return;
  const checks = candidates.map((address) => {
    const spill = sheet.getRange(address).getSpillingToRangeOrNullObject();
    spill.load("address");
    return { anchor: address, spill };
  });
  await context.sync();
  for (const { anchor, spill } of checks) {
    if (!spill.isNullObject) state.spills.push({ anchor, address: stripSheet(spill.address) });
  }
}

/** Note lines for the agent; empty when nothing is notable. */
export function renderRangeState(state: RangeState, columnLabel: (index: number) => string): string[] {
  const lines: string[] = [];
  if (state.calculationMode !== "Automatic") {
    lines.push(`⚠️ Calculation mode is ${state.calculationMode}: formula results may be stale until the workbook recalculates.`);
  }
  if (state.protected) {
    const cells = state.locked === true
      ? "every cell here is locked, so writes here fail until the sheet is unprotected"
      : state.locked === false
        ? "the cells here are unlocked and can be edited"
        : "cells here are a mix of locked and unlocked; writes to locked ones fail";
    lines.push(`🔒 Sheet is protected: ${cells}.`);
  }
  if (state.hiddenRows.length > 0) {
    lines.push(`Hidden rows (filtered or hidden): ${formatNumberRuns(state.hiddenRows)}`);
  }
  if (state.hiddenColumns.length > 0) {
    lines.push(`Hidden columns: ${formatNumberRuns(state.hiddenColumns, columnLabel)}`);
  }
  for (const table of state.tables) {
    lines.push(`Table "${table.name}": ${table.address}${table.filtered ? " (filter applied)" : ""}`);
  }
  if (state.autoFilter) {
    lines.push(`AutoFilter: ${state.autoFilter.address}${state.autoFilter.filtered ? " (filter applied)" : ""}`);
  }
  if (state.merged.length > 0) {
    const more = state.mergedMore > 0 ? `, +${state.mergedMore} more` : "";
    lines.push(`Merged: ${state.merged.join(", ")}${more} (value is in the top-left cell)`);
  }
  if (state.spills.length > 0) {
    lines.push(`Spills: ${state.spills.map((spill) => `${spill.anchor} → ${spill.address}`).join(", ")} (only the anchor holds the formula)`);
  }
  if (state.validation.length > 0) {
    const more = state.validationMore > 0 ? `; +${state.validationMore} more` : "";
    lines.push(`Data validation: ${state.validation.map((rule) => `${rule.address} ${rule.rule}`).join("; ")}${more}`);
  }
  if (state.unavailable) {
    lines.push(`(Some sheet state could not be read: ${state.unavailable})`);
  }
  return lines;
}
