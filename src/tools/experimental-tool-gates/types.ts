export const TMUX_TOOL_NAME = "tmux";
export const EXECUTE_OFFICE_JS_TOOL_NAME = "execute_office_js";

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

export type TmuxBridgeGateReason =
  | "tmux_experiment_disabled"
  | "missing_bridge_url"
  | "invalid_bridge_url"
  | "bridge_unreachable";

export interface TmuxBridgeGateResult {
  allowed: boolean;
  bridgeUrl?: string;
  reason?: TmuxBridgeGateReason;
}

export interface TmuxBridgeGateDependencies {
  isTmuxExperimentEnabled?: () => boolean;
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
}

export interface OfficeJsExecuteApprovalRequest {
  explanation: string;
  code: string;
}

export interface ExperimentalToolGateDependencies extends
  TmuxBridgeGateDependencies,
  PythonBridgeGateDependencies {
  requestPythonBridgeApproval?: (request: PythonBridgeApprovalRequest) => Promise<boolean>;
  getApprovedPythonBridgeUrl?: () => Promise<string | undefined>;
  setApprovedPythonBridgeUrl?: (bridgeUrl: string) => Promise<void>;
  requestOfficeJsExecuteApproval?: (request: OfficeJsExecuteApprovalRequest) => Promise<boolean>;
}
