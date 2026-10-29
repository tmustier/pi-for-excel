export interface DocumentInstanceIdentity {
  read(): string | null;

  ensure(): Promise<string | null>;
}

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/u;

export function isDocumentInstanceToken(value: string | null | undefined): value is string {
  return typeof value === "string" && UUID_PATTERN.test(value);
}
