/**
 * Public entrypoint for tool execution gates.
 *
 * This module intentionally re-exports split implementations to keep
 * imports stable while keeping each implementation file focused/smaller.
 */

export {
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

export { applyExperimentalToolGates } from "./experimental-tool-gates/wrappers.js";
