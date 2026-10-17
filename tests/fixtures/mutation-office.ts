import { parseCell, colToLetter, parseRangeRef } from "../../src/excel/helpers.ts";

type Grid = (string | number | boolean | null)[][];

/** Narrow Office boundary double. It models queued cell writes and row/column shifts, not Excel semantics. */
export function mutationOfficeFixture(initial: Grid = [["old", 17], ["anchor", 29]]) {
  let data = structuredClone(initial);
  const commands: (() => void)[] = [];
  let writes = 0;
  let onWrite = (): void => {};
  let onRangeLoad = (_address: string, _properties: string): void => {};
  let usedRangeAddress: string | null = null;
  let failAfterWrite = false;
  const load = (): void => {};
  const bounds = (address: string) => {
    const local = parseRangeRef(address).address;
    const [first = "", last = first] = local.split(":");
    if (/^\d+$/u.test(first)) return { row: Number(first) - 1, column: 0,
      rows: Number(last) - Number(first) + 1, columns: data.reduce((max, row) => Math.max(max, row.length), 1), band: "rows" };
    if (/^[A-Z]+$/iu.test(first)) {
      const start = parseCell(`${first}1`), end = parseCell(`${last}1`);
      return { row: 0, column: start.col, rows: data.length, columns: end.col - start.col + 1, band: "columns" };
    }
    const start = parseCell(first), end = parseCell(last);
    return { row: start.row - 1, column: start.col, rows: end.row - start.row + 1, columns: end.col - start.col + 1, band: "cells" };
  };
  const range = (address: string) => {
    const b = bounds(address);
    const raw = () => Array.from({ length: b.rows }, (_, r) =>
      Array.from({ length: b.columns }, (_, c) => data[b.row + r]?.[b.column + c] ?? ""));
    return {
      load: (properties: string): void => { onRangeLoad(address, properties); },
      isNullObject: false, rowIndex: b.row, columnIndex: b.column,
      rowCount: b.rows, columnCount: b.columns,
      get address() { return `${sheet.name}!${parseRangeRef(address).address}`; },
      get values(): Grid { return raw().map((row) => row.map((cell) => cell === "=1+1" ? 2 : cell)); },
      set values(value: Grid) {
        const copied = structuredClone(value);
        commands.push(() => {
          for (const [r, row] of copied.entries()) {
            const destination = data[b.row + r] ?? [];
            data[b.row + r] = destination;
            for (const [c, cell] of row.entries()) destination[b.column + c] = cell;
          }
        });
      },
      get formulas(): Grid { return raw().map((row) => row.map((cell) => typeof cell === "string" && cell.startsWith("=") ? cell : "")); },
      insert: () => commands.push(() => {
        if (b.band === "rows") data.splice(b.row, 0, ...Array.from({ length: b.rows }, () => Array.from({ length: b.columns }, () => "")));
        else for (const row of data) row.splice(b.column, 0, ...Array.from({ length: b.columns }, () => ""));
      }),
      delete: () => commands.push(() => {
        if (b.band === "rows") data.splice(b.row, b.rows);
        else for (const row of data) row.splice(b.column, b.columns);
      }),
      getUsedRangeOrNullObject: () => range(address),
    };
  };
  const table = { id: "table-1", name: "ExistingTable", columns: { load, items: [{ id: 1, name: "First" }] },
    getRange: () => ({ load, address: `${sheet.name}!D4:D8` }) };
  const sheet = {
    load, id: "sheet-1", name: "Original", position: 0, visibility: "Visible", isNullObject: false,
    tables: { load, items: [table] },
    getRange: range,
    getUsedRangeOrNullObject: () => range(usedRangeAddress ?? `A1:${colToLetter(data.reduce((max, row) => Math.max(max, row.length), 1) - 1)}${Math.max(data.length, 1)}`),
  };
  const worksheets = { load, items: [sheet], getActiveWorksheet: () => sheet,
    getItem: (id: string) => {
      if (id !== sheet.id && id !== sheet.name) throw new Error("Sheet identity absent");
      return sheet;
    }, getItemOrNullObject: (id: string) => {
      if (id === sheet.id || id === sheet.name) return sheet;
      throw new Error("Sheet identity absent");
    } };
  const names = {
    load, items: [] as { name: string; formula: string; type: string }[],
    add: (name: string, formula: string) => commands.push(() => { names.items.push({ name, formula, type: "Range" }); }),
  };
  const context = {
    workbook: { worksheets, names },
    sync: () => {
      if (!commands.length) return Promise.resolve();
      onWrite();
      for (const apply of commands.splice(0)) apply();
      writes += 1;
      if (failAfterWrite) return Promise.reject(new Error("Disconnected after workbook sync"));
      return Promise.resolve();
    },
  };
  return { sheet, table, names: () => structuredClone(names.items), read: () => structuredClone(data), writes: () => writes,
    edit: (grid: Grid) => { data = structuredClone(grid); },
    onWrite: (callback: () => void) => { onWrite = callback; },
    onRangeLoad: (callback: (address: string, properties: string) => void) => { onRangeLoad = callback; },
    setUsedRange: (address: string) => { usedRangeAddress = address; },
    failAfterWrite: () => { failAfterWrite = true; },
    excel: { run: <T>(action: (ctx: typeof context) => Promise<T>): Promise<T> => action(context) },
  };
}
