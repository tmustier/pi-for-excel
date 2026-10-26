export const TMUX_TOOL_NAME = "tmux";
export const FILES_TOOL_NAME = "files";
export const EXTENSIONS_MANAGER_TOOL_NAME = "extensions_manager";
export const EXECUTE_OFFICE_JS_TOOL_NAME = "execute_office_js";
export const EXECUTE_WPS_JS_TOOL_NAME = "execute_wps_js";

/** Python tools that can run without a native bridge (via Pyodide fallback). */
export const PYTHON_FALLBACK_TOOL_NAMES = new Set<string>([
  "python_run",
  "python_transform_range",
]);

/** Python tools that strictly require the native bridge. */
export const PYTHON_BRIDGE_ONLY_TOOL_NAMES = new Set<string>([
  "libreoffice_convert",
]);

export const TMUX_BRIDGE_URL_SETTING_KEY = "tmux.bridge.url";
export const PYTHON_BRIDGE_URL_SETTING_KEY = "python.bridge.url";
export const PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY = "python.bridge.approved.url";

export const DEFAULT_TMUX_BRIDGE_URL = "https://localhost:3341";
export const DEFAULT_PYTHON_BRIDGE_URL = "https://localhost:3340";

export type TmuxBridgeGateReason =
  | "missing_bridge_url"
  | "invalid_bridge_url"
  | "bridge_unreachable";

export interface TmuxBridgeGateResult {
  allowed: boolean;
  bridgeUrl?: string;
  reason?: TmuxBridgeGateReason;
}

export interface TmuxBridgeGateDependencies {
  getTmuxBridgeUrl?: () => Promise<string | undefined>;
  validateBridgeUrl?: (url: string) => string | null;
  probeTmuxBridge?: (bridgeUrl: string) => Promise<boolean>;
}

export type PythonBridgeGateReason =
  | "missing_bridge_url"
  | "invalid_bridge_url"
  | "bridge_unreachable";

export interface PythonBridgeGateResult {
  allowed: boolean;
  bridgeUrl?: string;
  reason?: PythonBridgeGateReason;
}

export interface PythonBridgeGateDependencies {
  getPythonBridgeUrl?: () => Promise<string | undefined>;
  validatePythonBridgeUrl?: (url: string) => string | null;
  probePythonBridge?: (bridgeUrl: string) => Promise<boolean>;
}

export interface PythonBridgeApprovalRequest {
  toolName: string;
  bridgeUrl: string;
  params: unknown;
  /** Aborts when the turn is stopped; the request should then be withdrawn. */
  signal?: AbortSignal;
}

export interface OfficeJsExecuteApprovalRequest {
  explanation: string;
  code: string;
  apiName?: "Office.js" | "WPS JSAPI";
  /**
   * Ambient-authority identifiers detected by the static risk lint
   * (`office-js-risk.ts`). When present, approval is requested even in Auto
   * mode and the approval message includes an explicit warning.
   */
  riskIdentifiers?: string[];
  /** Aborts when the turn is stopped; the request should then be withdrawn. */
  signal?: AbortSignal;
}

/** A Files workspace entry as seen by the files approval gate. */
export interface FilesGateFileInfo {
  size: number;
  readOnly: boolean;
  /** Set when the file lives in a connected local folder (changes hit the disk). */
  nativeFolderName?: string;
}

/** A consent request for an action that changes local state outside the workbook. */
export interface LocalChangeApprovalRequest {
  toolName: string;
  title: string;
  message: string;
  confirmLabel: string;
  destructive: boolean;
  /** The tool call's signal: when the turn stops, the pending request is withdrawn. */
  signal?: AbortSignal;
}

export interface ExperimentalToolGateDependencies extends
  TmuxBridgeGateDependencies,
  PythonBridgeGateDependencies {
  requestPythonBridgeApproval?: (request: PythonBridgeApprovalRequest) => Promise<boolean>;
  getApprovedPythonBridgeUrl?: () => Promise<string | undefined>;
  setApprovedPythonBridgeUrl?: (bridgeUrl: string) => Promise<void>;
  requestOfficeJsExecuteApproval?: (request: OfficeJsExecuteApprovalRequest) => Promise<boolean>;
  /** When provided, Auto mode skips the Office.js approval prompt. */
  getExecutionMode?: () => Promise<"yolo" | "safe">;
  /**
   * Consent for `files` writes/deletes and `extensions_manager` changes.
   * Without it those actions fail closed whenever approval is required.
   */
  requestLocalChangeApproval?: (request: LocalChangeApprovalRequest) => Promise<boolean>;
  /** Look up a Files workspace path; `null` when it does not exist. */
  inspectWorkspaceFile?: (path: string) => Promise<FilesGateFileInfo | null>;
  /**
   * Paths the assistant created in this session. Auto mode lets Pi overwrite
   * or delete its own new files without asking; anything else asks.
   */
  assistantCreatedFiles?: Set<string>;
  /** Installed extensions, for naming the target of an approval. */
  listExtensions?: () => ReadonlyArray<{ id: string; name: string; grantedCapabilities: readonly string[] }>;
}
