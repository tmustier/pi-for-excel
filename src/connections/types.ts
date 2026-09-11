import { Type, type Static } from "typebox";

import { StringEnum } from "../tools/string-enum.js";

export const connectionStatusSchema = StringEnum(["connected", "missing", "invalid", "error"]);

export type ConnectionStatus = Static<typeof connectionStatusSchema>;

export type ConnectionAuthKind = "api_key" | "bearer_token" | "oauth" | "custom";

export interface ConnectionSecretFieldDefinition {
  id: string;
  label: string;
  required: boolean;
  maskInUi?: boolean;
}

export interface ConnectionHttpAuthDefinition {
  /** Currently only header auth injection is supported. */
  placement: "header";
  /** Header name to inject (for example: Authorization, X-API-Key). */
  headerName: string;
  /** Template where placeholders reference secret field ids (for example: Bearer {apiKey}). */
  valueTemplate: string;
  /** Exact target hosts allowed for auth injection. */
  allowedHosts: readonly string[];
}

export interface ConnectionDefinition {
  id: string;
  title: string;
  /** Short capability sentence shown to the assistant in prompt context. */
  capability: string;
  authKind: ConnectionAuthKind;
  secretFields: readonly ConnectionSecretFieldDefinition[];
  /** Optional host-injected HTTP auth mapping for api.http.fetch({ connection }). */
  httpAuth?: ConnectionHttpAuthDefinition;
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

export const connectionToolErrorCodeSchema = StringEnum([
  "missing_connection",
  "invalid_connection",
  "connection_auth_failed",
]);

export type ConnectionToolErrorCode = Static<typeof connectionToolErrorCodeSchema>;

export const connectionToolErrorDetailsSchema = Type.Object({
  kind: Type.Literal("connection_error"),
  ok: Type.Literal(false),
  errorCode: connectionToolErrorCodeSchema,
  connectionId: Type.String(),
  connectionTitle: Type.String(),
  status: connectionStatusSchema,
  setupHint: Type.String(),
  reason: Type.Optional(Type.String()),
});

export type ConnectionToolErrorDetails = Static<typeof connectionToolErrorDetailsSchema>;

export interface ConnectionRuntimeAuthFailure {
  message: string;
}
