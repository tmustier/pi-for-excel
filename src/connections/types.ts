export type ConnectionStatus = "connected" | "missing" | "invalid" | "error";

export type ConnectionAuthKind = "api_key" | "bearer_token" | "oauth" | "custom";

export interface ConnectionSecretFieldDefinition {
  id: string;
  label: string;
  required: boolean;
  maskInUi?: boolean;
}

export interface ConnectionDefinition {
  id: string;
  title: string;
  /** Short capability sentence shown to the assistant in prompt context. */
  capability: string;
  authKind: ConnectionAuthKind;
  secretFields: readonly ConnectionSecretFieldDefinition[];
  /** Optional custom setup hint for structured tool failures. */
  setupHint?: string;
}

export interface ConnectionSecrets {
  [fieldId: string]: string;
}

export interface ConnectionState {
  connectionId: string;
  status: ConnectionStatus;
  lastValidatedAt?: string;
  lastError?: string;
}

export interface ConnectionSnapshot extends ConnectionState {
  title: string;
  capability: string;
  setupHint: string;
}

export interface ConnectionPromptEntry {
  id: string;
  title: string;
  capability: string;
  status: ConnectionStatus;
  setupHint: string;
  lastError?: string;
}

export type ConnectionToolErrorCode =
  | "missing_connection"
  | "invalid_connection"
  | "connection_auth_failed";

export interface ConnectionToolErrorDetails {
  kind: "connection_error";
  ok: false;
  errorCode: ConnectionToolErrorCode;
  connectionId: string;
  connectionTitle: string;
  status: ConnectionStatus;
  setupHint: string;
  reason?: string;
}

export interface ConnectionRuntimeAuthFailure {
  message: string;
}
