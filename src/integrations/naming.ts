/**
 * Terminology for built-in external capability bundles.
 *
 * Keep user-facing strings centralized so renaming "integrations"
 * later only requires edits in one place.
 */

export const INTEGRATIONS_COMMAND_NAME = "integrations";

export const INTEGRATION_LABEL = "Integration";
export const INTEGRATION_LABEL_LOWER = "integration";

export const INTEGRATIONS_LABEL = "Integrations";
export const INTEGRATIONS_LABEL_LOWER = "integrations";

export const INTEGRATIONS_MANAGER_LABEL = "Tools & MCP";
export const INTEGRATIONS_MANAGER_LABEL_LOWER = "tools & MCP";

export const ACTIVE_INTEGRATIONS_PROMPT_HEADING = "Active Integrations";
export const ACTIVE_INTEGRATIONS_TOOLTIP_PREFIX = "Active integrations";

export function integrationsCommandHint(): string {
  return `/${INTEGRATIONS_COMMAND_NAME}`;
}

export function formatIntegrationCountLabel(count: number): string {
  const noun = count === 1 ? INTEGRATION_LABEL_LOWER : INTEGRATIONS_LABEL_LOWER;
  return `${count} ${noun}`;
}
