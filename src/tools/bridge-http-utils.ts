function isToolsBridgeHttpUtilsPayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


export function tryParseBridgeJson(text: string): DynamicValue {
  const trimmed = text.trim();
  if (trimmed.length === 0) return null;

  try {
    return JSON.parse(trimmed);
  } catch {
    return null;
  }
}

export function extractBridgeErrorMessage(value: DynamicValue): string | null {
  if (isToolsBridgeHttpUtilsPayloadShape(value) && typeof value.error === "string") {
    return value.error;
  }

  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : null;
  }

  return null;
}

export function joinBridgeUrl(baseUrl: string, path: string): string {
  return `${baseUrl.replace(/\/+$/, "")}${path}`;
}

export function isAbortError(error: DynamicValue): boolean {
  if (error instanceof DOMException) {
    return error.name === "AbortError";
  }

  if (error instanceof Error) {
    return error.name === "AbortError";
  }

  return false;
}
