import {
  INTEGRATIONS_LABEL,
} from "../../integrations/naming.js";
import {
  WEB_SEARCH_PROVIDERS,
  WEB_SEARCH_PROVIDER_INFO,
} from "../../tools/web-search-config.js";
import {
  createOverlayButton,
  createOverlayInput,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";

export interface IntegrationsDialogElements {
  body: HTMLDivElement;
  externalToggle: HTMLInputElement;
  activeSummary: HTMLDivElement;
  integrationsList: HTMLDivElement;
  webSearchStatus: HTMLDivElement;
  webSearchProviderSelect: HTMLSelectElement;
  webSearchProviderSignupLink: HTMLAnchorElement;
  webSearchApiKeyInput: HTMLInputElement;
  webSearchSaveButton: HTMLButtonElement;
  webSearchValidateButton: HTMLButtonElement;
  webSearchClearButton: HTMLButtonElement;
  webSearchHint: HTMLParagraphElement;
  webSearchValidationStatus: HTMLParagraphElement;
  mcpList: HTMLDivElement;
  mcpNameInput: HTMLInputElement;
  mcpUrlInput: HTMLInputElement;
  mcpTokenInput: HTMLInputElement;
  mcpEnabledInput: HTMLInputElement;
  mcpAddButton: HTMLButtonElement;
}

export function createIntegrationsDialogElements(): IntegrationsDialogElements {
  const body = document.createElement("div");
  body.className = "pi-overlay-body";

  const externalSection = document.createElement("section");
  externalSection.className = "pi-overlay-section";
  externalSection.appendChild(createOverlaySectionTitle("External tools gate"));

  const externalCard = document.createElement("div");
  externalCard.className = "pi-overlay-surface";

  const externalToggleLabel = document.createElement("label");
  externalToggleLabel.className = "pi-integrations-toggle-label";

  const externalToggle = document.createElement("input");
  externalToggle.type = "checkbox";

  const externalToggleText = document.createElement("span");
  externalToggleText.textContent = "Allow external tools (web search / MCP)";

  externalToggleLabel.append(externalToggle, externalToggleText);

  const activeSummary = document.createElement("div");
  activeSummary.className = "pi-integrations-active-summary";

  externalCard.append(externalToggleLabel, activeSummary);
  externalSection.appendChild(externalCard);

  const integrationsSection = document.createElement("section");
  integrationsSection.className = "pi-overlay-section";
  integrationsSection.appendChild(createOverlaySectionTitle(`${INTEGRATIONS_LABEL} bundles`));

  const integrationsList = document.createElement("div");
  integrationsList.className = "pi-overlay-list";
  integrationsSection.appendChild(integrationsList);

  const webSearchSection = document.createElement("section");
  webSearchSection.className = "pi-overlay-section";
  webSearchSection.appendChild(createOverlaySectionTitle("Web search config"));

  const webSearchCard = document.createElement("div");
  webSearchCard.className = "pi-overlay-surface";

  const webSearchStatus = document.createElement("div");
  webSearchStatus.className = "pi-integrations-web-search-status";

  const webSearchProviderRow = document.createElement("div");
  webSearchProviderRow.className = "pi-integrations-web-search-provider-row";

  const webSearchProviderSelect = document.createElement("select");
  webSearchProviderSelect.className = "pi-overlay-input";

  for (const providerId of WEB_SEARCH_PROVIDERS) {
    const option = document.createElement("option");
    option.value = providerId;
    option.textContent = WEB_SEARCH_PROVIDER_INFO[providerId].title;
    webSearchProviderSelect.appendChild(option);
  }

  const webSearchProviderSignupLink = document.createElement("a");
  webSearchProviderSignupLink.className = "pi-overlay-link";
  webSearchProviderSignupLink.target = "_blank";
  webSearchProviderSignupLink.rel = "noopener noreferrer";

  webSearchProviderRow.append(webSearchProviderSelect, webSearchProviderSignupLink);

  const webSearchInputRow = document.createElement("div");
  webSearchInputRow.className = "pi-integrations-web-search-row";

  const webSearchApiKeyInput = createOverlayInput({ placeholder: "API key", type: "password" });
  const webSearchSaveButton = createOverlayButton({ text: "Save key" });
  const webSearchValidateButton = createOverlayButton({ text: "Validate" });
  const webSearchClearButton = createOverlayButton({ text: "Clear" });

  webSearchInputRow.append(
    webSearchApiKeyInput,
    webSearchSaveButton,
    webSearchValidateButton,
    webSearchClearButton,
  );

  const webSearchHint = document.createElement("p");
  webSearchHint.className = "pi-overlay-hint";

  const webSearchValidationStatus = document.createElement("p");
  webSearchValidationStatus.className = "pi-overlay-hint";

  webSearchCard.append(
    webSearchStatus,
    webSearchProviderRow,
    webSearchInputRow,
    webSearchHint,
    webSearchValidationStatus,
  );
  webSearchSection.appendChild(webSearchCard);

  const mcpSection = document.createElement("section");
  mcpSection.className = "pi-overlay-section";
  mcpSection.appendChild(createOverlaySectionTitle("MCP servers"));

  const mcpList = document.createElement("div");
  mcpList.className = "pi-overlay-list";

  const mcpAddCard = document.createElement("div");
  mcpAddCard.className = "pi-overlay-surface";

  const mcpAddTitle = document.createElement("div");
  mcpAddTitle.textContent = "Add server";
  mcpAddTitle.className = "pi-integrations-mcp-add-title";

  const mcpAddRow = document.createElement("div");
  mcpAddRow.className = "pi-integrations-mcp-add-row";

  const mcpNameInput = createOverlayInput({ placeholder: "Name" });
  const mcpUrlInput = createOverlayInput({ placeholder: "https://example.com/mcp" });
  const mcpTokenInput = createOverlayInput({ placeholder: "Bearer token (optional)", type: "password" });

  const mcpEnabledLabel = document.createElement("label");
  mcpEnabledLabel.className = "pi-integrations-toggle-label";
  const mcpEnabledInput = document.createElement("input");
  mcpEnabledInput.type = "checkbox";
  mcpEnabledInput.checked = true;
  mcpEnabledLabel.append(mcpEnabledInput, document.createTextNode("Enabled"));

  const mcpAddButton = createOverlayButton({ text: "Add" });

  mcpAddRow.append(mcpNameInput, mcpUrlInput, mcpTokenInput, mcpEnabledLabel, mcpAddButton);

  const mcpHint = document.createElement("p");
  mcpHint.className = "pi-overlay-hint";
  mcpHint.textContent = "Server URL, optional bearer token, and one-click connection test.";

  mcpAddCard.append(mcpAddTitle, mcpAddRow, mcpHint);

  mcpSection.append(mcpList, mcpAddCard);

  body.append(externalSection, integrationsSection, webSearchSection, mcpSection);

  return {
    body,
    externalToggle,
    activeSummary,
    integrationsList,
    webSearchStatus,
    webSearchProviderSelect,
    webSearchProviderSignupLink,
    webSearchApiKeyInput,
    webSearchSaveButton,
    webSearchValidateButton,
    webSearchClearButton,
    webSearchHint,
    webSearchValidationStatus,
    mcpList,
    mcpNameInput,
    mcpUrlInput,
    mcpTokenInput,
    mcpEnabledInput,
    mcpAddButton,
  };
}
