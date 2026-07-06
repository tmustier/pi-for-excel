/**
 * Error utilities.
 *
 * We often catch `unknown` (or anything) at runtime; this helper normalizes
 * it into a user-facing string without relying on `any`.
 */
export function getErrorMessage(error: DynamicValue): string {
  if (error instanceof Error) return error.message;

  if (typeof error === "string") return error;

  if (error && typeof error === "object" && "message" in error) {
    const maybeMessage = (error as { message?: DynamicValue }).message;
    if (typeof maybeMessage === "string") return maybeMessage;
  }

  try {
    const json = JSON.stringify(error);
    return typeof json === "string" ? json : Object.prototype.toString.call(error);
  } catch {
    return Object.prototype.toString.call(error);
  }
}
