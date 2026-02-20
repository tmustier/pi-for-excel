/**
 * Settings section for custom OpenAI-compatible gateways.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  deleteOpenAiGatewayConfig,
  listOpenAiGatewayConfigs,
  saveOpenAiGatewayConfig,
  type OpenAiGatewayConfig,
} from "../../auth/custom-gateways.js";
import {
  createButton,
  createConfigInput,
  createConfigRow,
} from "../../ui/extensions-hub-components.js";
import { createOverlaySectionTitle } from "../../ui/overlay-dialog.js";
import { showToast } from "../../ui/toast.js";

interface BuildCustomGatewaySectionOptions {
  onProvidersChanged: () => void;
}

function createHint(text: string): HTMLParagraphElement {
  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = text;
  return hint;
}

function createGatewayCard(args: {
  gateway: OpenAiGatewayConfig;
  onEdit: (gateway: OpenAiGatewayConfig) => void;
  onDelete: (gateway: OpenAiGatewayConfig) => void;
}): HTMLElement {
  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-settings-gateway-item";

  const topRow = document.createElement("div");
  topRow.className = "pi-settings-gateway-item__top";

  const titleGroup = document.createElement("div");
  titleGroup.className = "pi-settings-gateway-item__title-group";

  const title = document.createElement("p");
  title.className = "pi-settings-gateway-item__title";
  title.textContent = args.gateway.displayName;

  const provider = document.createElement("p");
  provider.className = "pi-settings-gateway-item__provider";
  provider.textContent = args.gateway.providerName;

  titleGroup.append(title, provider);

  const actions = document.createElement("div");
  actions.className = "pi-settings-gateway-item__actions";

  const editButton = createButton("Edit", {
    compact: true,
    onClick: () => {
      args.onEdit(args.gateway);
    },
  });

  const deleteButton = createButton("Delete", {
    compact: true,
    danger: true,
    onClick: () => {
      args.onDelete(args.gateway);
    },
  });

  actions.append(editButton, deleteButton);
  topRow.append(titleGroup, actions);

  const endpoint = document.createElement("p");
  endpoint.className = "pi-settings-gateway-item__meta";
  endpoint.textContent = `Endpoint: ${args.gateway.endpointUrl}`;

  const model = document.createElement("p");
  model.className = "pi-settings-gateway-item__meta";
  model.textContent = `Model: ${args.gateway.modelId}`;

  const keyState = document.createElement("p");
  keyState.className = "pi-settings-gateway-item__meta";
  keyState.textContent = args.gateway.apiKey.length > 0 ? "API key: configured" : "API key: none";

  card.append(topRow, endpoint, model, keyState);
  return card;
}

export async function buildCustomGatewaySection(
  options: BuildCustomGatewaySectionOptions,
): Promise<HTMLElement> {
  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-settings-section";
  section.dataset.settingsAnchor = "custom-gateways";

  const title = createOverlaySectionTitle("Custom OpenAI-compatible gateways");
  const hint = createHint(
    "Use this for company LLM gateways or local OpenAI-compatible servers.",
  );

  const content = document.createElement("div");
  content.className = "pi-settings-section__content";

  const formCard = document.createElement("div");
  formCard.className = "pi-overlay-surface pi-settings-gateway-form";

  const nameInput = createConfigInput({
    placeholder: "Gateway name (optional)",
  });

  const endpointInput = createConfigInput({
    placeholder: "https://your-gateway.example.com/v1",
  });
  endpointInput.spellcheck = false;

  const modelInput = createConfigInput({
    placeholder: "model-id",
  });

  const apiKeyInput = createConfigInput({
    placeholder: "API key (optional for local servers)",
    type: "password",
  });

  const errorText = document.createElement("p");
  errorText.className = "pi-overlay-hint pi-overlay-text-warning";
  errorText.hidden = true;

  const formActions = document.createElement("div");
  formActions.className = "pi-overlay-actions";

  const cancelButton = createButton("Cancel", {
    compact: true,
  });
  cancelButton.hidden = true;

  const saveButton = createButton("Save gateway", {
    compact: true,
    primary: true,
  });

  formActions.append(cancelButton, saveButton);

  formCard.append(
    createConfigRow("Name", nameInput),
    createConfigRow("Endpoint", endpointInput),
    createConfigRow("Model", modelInput),
    createConfigRow("API key", apiKeyInput),
    errorText,
    formActions,
  );

  const listTitle = document.createElement("p");
  listTitle.className = "pi-settings-gateway-list__title";
  listTitle.textContent = "Configured gateways";

  const listHost = document.createElement("div");
  listHost.className = "pi-settings-gateway-list";

  let editingGatewayId: string | null = null;
  let gateways: OpenAiGatewayConfig[] = [];

  const setError = (message: string | null): void => {
    if (!message) {
      errorText.hidden = true;
      errorText.textContent = "";
      return;
    }

    errorText.hidden = false;
    errorText.textContent = message;
  };

  const resetForm = (): void => {
    editingGatewayId = null;
    nameInput.value = "";
    endpointInput.value = "";
    modelInput.value = "";
    apiKeyInput.value = "";
    cancelButton.hidden = true;
    saveButton.textContent = "Save gateway";
    setError(null);
  };

  const startEditing = (gateway: OpenAiGatewayConfig): void => {
    editingGatewayId = gateway.id;
    nameInput.value = gateway.displayName;
    endpointInput.value = gateway.endpointUrl;
    modelInput.value = gateway.modelId;
    apiKeyInput.value = gateway.apiKey;
    cancelButton.hidden = false;
    saveButton.textContent = "Update gateway";
    setError(null);
    nameInput.focus();
  };

  const reloadGateways = async (): Promise<void> => {
    gateways = await listOpenAiGatewayConfigs(getAppStorage().customProviders);
  };

  const renderList = (): void => {
    listHost.replaceChildren();

    if (gateways.length === 0) {
      listHost.appendChild(createHint("No custom gateways configured yet."));
      return;
    }

    for (const gateway of gateways) {
      listHost.appendChild(createGatewayCard({
        gateway,
        onEdit: startEditing,
        onDelete: (targetGateway) => {
          const confirmed = confirm(`Delete gateway \"${targetGateway.displayName}\"?`);
          if (!confirmed) {
            return;
          }

          void (async () => {
            try {
              await deleteOpenAiGatewayConfig(getAppStorage().customProviders, targetGateway.id);
              await reloadGateways();
              renderList();
              options.onProvidersChanged();
              showToast(`Deleted gateway ${targetGateway.displayName}.`);

              if (editingGatewayId === targetGateway.id) {
                resetForm();
              }
            } catch (error: unknown) {
              const message = error instanceof Error ? error.message : String(error);
              showToast(`Failed to delete gateway: ${message}`);
            }
          })();
        },
      }));
    }
  };

  cancelButton.addEventListener("click", () => {
    resetForm();
  });

  saveButton.addEventListener("click", () => {
    void (async () => {
      setError(null);
      saveButton.disabled = true;
      cancelButton.disabled = true;

      try {
        const saved = await saveOpenAiGatewayConfig(getAppStorage().customProviders, {
          id: editingGatewayId ?? undefined,
          displayName: nameInput.value,
          endpointUrl: endpointInput.value,
          modelId: modelInput.value,
          apiKey: apiKeyInput.value,
        });

        await reloadGateways();
        renderList();
        options.onProvidersChanged();
        showToast(
          editingGatewayId
            ? `Updated gateway ${saved.displayName}.`
            : `Saved gateway ${saved.displayName}.`,
        );

        resetForm();
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : String(error);
        setError(message);
      } finally {
        saveButton.disabled = false;
        cancelButton.disabled = false;
      }
    })();
  });

  await reloadGateways();
  renderList();

  content.append(formCard, listTitle, listHost);
  section.append(title, hint, content);
  return section;
}
