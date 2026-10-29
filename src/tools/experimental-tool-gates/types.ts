export const TMUX_TOOL_NAME = "tmux";
export const FILES_TOOL_NAME = "files";
export const EXTENSIONS_MANAGER_TOOL_NAME = "extensions_manager";
export const EXECUTE_OFFICE_JS_TOOL_NAME = "execute_office_js";
export const EXECUTE_WPS_JS_TOOL_NAME = "execute_wps_js";

export const PYTHON_FALLBACK_TOOL_NAMES = new Set<string>([
  "python_run",
  "python_transform_range",
]);

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
  signal?: AbortSignal;
}

export interface OfficeJsExecuteApprovalRequest {
  toolCallId: string;
  explanation: string;
  code: string;
  apiName?: "Office.js" | "WPS JSAPI";
  riskIdentifiers?: string[];
  destructive?: boolean;
  signal?: AbortSignal;
}

export interface FilesGateFileInfo {
  size: number;
  readOnly: boolean;
  nativeFolderName?: string;
}

export interface LocalChangeApprovalRequest {
  toolName: string;
  title: string;
  message: string;
  confirmLabel: string;
  destructive: boolean;
  signal?: AbortSignal;
}

export interface ExperimentalToolGateDependencies extends
  TmuxBridgeGateDependencies,
  PythonBridgeGateDependencies {
  requestPythonBridgeApproval?: (request: PythonBridgeApprovalRequest) => Promise<boolean>;
  getApprovedPythonBridgeUrl?: () => Promise<string | undefined>;
  setApprovedPythonBridgeUrl?: (bridgeUrl: string) => Promise<void>;
  requestOfficeJsExecuteApproval?: (request: OfficeJsExecuteApprovalRequest) => Promise<boolean>;
  getExecutionMode?: () => Promise<"yolo" | "safe">;
  requestLocalChangeApproval?: (request: LocalChangeApprovalRequest) => Promise<boolean>;
  inspectWorkspaceFile?: (path: string) => Promise<FilesGateFileInfo | null>;
  assistantCreatedFiles?: Set<string>;
  listExtensions?: () => ReadonlyArray<{ id: string; name: string; grantedCapabilities: readonly string[] }>;
}
