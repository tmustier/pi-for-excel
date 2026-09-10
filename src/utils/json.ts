/** A value that JSON can represent. What `JSON.parse` returns, and nothing else. */
export type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue };

/** `JSON.parse` typed by what it actually produces. Throws on invalid JSON like `JSON.parse`. */
export function parseJson(text: string): JsonValue {
  return JSON.parse(text) as JsonValue;
}

/** Round-trip any serialisable value through JSON: drops `undefined` members, keeps everything else. */
export function toJsonValue(value: Parameters<typeof JSON.stringify>[0]): JsonValue {
  return parseJson(JSON.stringify(value));
}

/** Deterministic JSON: keys sorted at every depth. Equal values give equal strings. */
export function canonicalJson(value: JsonValue): string {
  if (Array.isArray(value)) {
    return `[${value.map(canonicalJson).join(",")}]`;
  }
  if (typeof value === "object" && value !== null) {
    const members = Object.entries(value).sort(([left], [right]) => (left < right ? -1 : left > right ? 1 : 0))
      .map(([key, member]) => `${JSON.stringify(key)}:${canonicalJson(member)}`);
    return `{${members.join(",")}}`;
  }
  return JSON.stringify(value);
}
