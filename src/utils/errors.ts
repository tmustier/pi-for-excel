export function getErrorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;

  if (typeof error === "string") return error;

  if (error && typeof error === "object" && "message" in error) {
    const maybeMessage = (error as { message?: unknown }).message;
    if (typeof maybeMessage === "string") return maybeMessage;
  }

  try {
    const json = JSON.stringify(error);
    return typeof json === "string" ? json : Object.prototype.toString.call(error);
  } catch {
    return Object.prototype.toString.call(error);
  }
}
