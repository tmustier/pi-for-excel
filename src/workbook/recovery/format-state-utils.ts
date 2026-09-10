import type { RecoveryFormatRangeState } from "./types.js";

function normalizeRecoveryAddress(address: string): string {
  return address.trim();
}

export function dedupeRecoveryAddresses(addresses: readonly string[]): string[] {
  const unique = new Set<string>();
  const ordered: string[] = [];

  for (const rawAddress of addresses) {
    const address = normalizeRecoveryAddress(rawAddress);
    if (address.length === 0 || unique.has(address)) {
      continue;
    }

    unique.add(address);
    ordered.push(address);
  }

  return ordered;
}

export function collectMergedAreaAddresses(state: RecoveryFormatRangeState): string[] {
  const addresses: string[] = [];

  for (const area of state.areas) {
    if (!Array.isArray(area.mergedAreas)) {
      continue;
    }

    for (const address of area.mergedAreas) {
      addresses.push(address);
    }
  }

  return dedupeRecoveryAddresses(addresses);
}

export function validateStringGrid(
  value: unknown,
  rowCount: number,
  columnCount: number,
): string[][] | null {
  if (!Array.isArray(value) || value.length !== rowCount) {
    return null;
  }

  const out: string[][] = [];

  for (const rowValue of value) {
    if (!Array.isArray(rowValue) || rowValue.length !== columnCount) {
      return null;
    }

    const outRow: string[] = [];
    for (const cellValue of rowValue) {
      if (typeof cellValue !== "string") {
        return null;
      }
      outRow.push(cellValue);
    }

    out.push(outRow);
  }

  return out;
}
