/**
 * Public entrypoint for tool execution gates.
 *
 * This module intentionally re-exports split implementations to keep
 * imports stable while keeping each implementation file focused/smaller.
 */

export {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
  EXECUTE_WPS_JS_TOOL_NAME,
  PYTHON_BRIDGE_APPROVED_URL_SETTING_KEY,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
  type ExperimentalToolGateDependencies,
  type OfficeJsExecuteApprovalRequest,
  type PythonBridgeApprovalRequest,
  type PythonBridgeGateDependencies,
  type PythonBridgeGateReason,
  type PythonBridgeGateResult,
  type TmuxBridgeGateDependencies,
  type TmuxBridgeGateReason,
  type TmuxBridgeGateResult,
} from "./experimental-tool-gates/types.js";

export {
  buildPythonBridgeGateErrorMessage,
  buildTmuxBridgeGateErrorMessage,
  evaluatePythonBridgeGate,
  evaluateTmuxBridgeGate,
} from "./experimental-tool-gates/evaluation.js";

export {
  applyExperimentalToolGates,
  buildOfficeJsExecuteApprovalMessage,
  buildPythonBridgeApprovalMessage,
} from "./experimental-tool-gates/wrappers.js";

export {
  assessOfficeJsCodeRisk,
  type OfficeJsCodeRiskAssessment,
} from "./experimental-tool-gates/office-js-risk.js";
