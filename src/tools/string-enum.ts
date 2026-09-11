/**
 * pi-ai's `StringEnum` with a `const` type parameter.
 *
 * pi-ai infers `T` as `readonly string[]` for an inline array, which turns the
 * schema's static type into `string`. The `const` modifier keeps the literal
 * tuple, so `Static<typeof schema>["action"]` stays the union of the listed
 * values without `as const` at every call site. Runtime output is identical.
 */

import { StringEnum as piStringEnum } from "@earendil-works/pi-ai";
import type { TUnsafe } from "typebox";

export function StringEnum<const T extends readonly string[]>(
  values: T,
  options?: { description?: string; default?: T[number] },
): TUnsafe<T[number]> {
  return piStringEnum(values, options);
}
