/** Address helpers shared by recovery modules. */

export function localAddressPart(address: string): string {
  const trimmed = address.trim();
  const separatorIndex = trimmed.lastIndexOf("!");
  if (separatorIndex < 0) {
    return trimmed;
  }

  return trimmed.slice(separatorIndex + 1);
}

export function quoteSheetName(sheetName: string): string {
  const escaped = sheetName.replace(/'/g, "''");
  const needsQuote = /[\s'!]/.test(sheetName);
  return needsQuote ? `'${escaped}'` : sheetName;
}

export function qualifyAddressWithSheet(sheetName: string, address: string): string {
  const local = localAddressPart(address);
  return `${quoteSheetName(sheetName)}!${local}`;
}

export function firstCellAddress(address: string): string {
  const local = localAddressPart(address);
  const firstArea = local.split(",")[0] ?? local;
  const first = firstArea.split(":")[0] ?? firstArea;
  return first.trim();
}

function parseCellCoordinates(address: string): { column: number; row: number } | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/iu.exec(address.trim());
  const columnLetters = match?.[1];
  const rowText = match?.[2];
  if (!columnLetters || !rowText) return null;

  let column = 0;
  for (const letter of columnLetters.toUpperCase()) {
    column = column * 26 + letter.charCodeAt(0) - 64;
  }

  const row = Number(rowText);
  if (!Number.isSafeInteger(row) || row <= 0) return null;
  return { column, row };
}

function parseCellRangeBounds(address: string): {
  start: { column: number; row: number };
  end: { column: number; row: number };
} | null {
  const parts = localAddressPart(address).split(":");
  if (parts.length > 2) return null;

  const start = parseCellCoordinates(parts[0] ?? "");
  const end = parseCellCoordinates(parts[1] ?? parts[0] ?? "");
  if (!start || !end || end.row < start.row || end.column < start.column) return null;
  return { start, end };
}

export function addressMatchesRangeShape(
  address: string,
  rowCount: number,
  columnCount: number,
): boolean {
  const bounds = parseCellRangeBounds(address);
  if (!bounds) return false;

  return bounds.end.row - bounds.start.row + 1 === rowCount &&
    bounds.end.column - bounds.start.column + 1 === columnCount;
}

export function addressIsWithinRowBand(address: string, position: number, count: number): boolean {
  const bounds = parseCellRangeBounds(address);
  const end = position + count - 1;
  if (!bounds || !Number.isSafeInteger(end)) return false;
  return bounds.start.row >= position && bounds.end.row <= end;
}

export function addressIsWithinColumnBand(address: string, position: number, count: number): boolean {
  const bounds = parseCellRangeBounds(address);
  const end = position + count - 1;
  if (!bounds || !Number.isSafeInteger(end)) return false;
  return bounds.start.column >= position && bounds.end.column <= end;
}

export function splitRangeList(range: string): string[] {
  return range
    .split(/[;,]/)
    .map((part) => part.trim())
    .filter((part) => part.length > 0);
}
