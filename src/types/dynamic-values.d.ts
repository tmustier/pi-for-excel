declare global {
  /**
   * Single sanctioned marker for values at external runtime boundaries
   * (JSON payloads, Office.js, browser events, and extension sandboxes) before
   * they are normalized into domain-specific types.
   */
  type DynamicValue = unknown;

  /** Boundary object shape for dynamic payloads before domain-specific parsing. */
  interface DynamicObject {
    [key: string]: DynamicValue;
  }
}

export {};
