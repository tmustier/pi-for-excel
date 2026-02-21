/**
 * Extensions hub — Extension connections section.
 *
 * Renders manager-backed connection cards in the Connections tab,
 * between the web search section and MCP servers.
 */

import type { ConnectionManager } from "../../connections/manager.js";
import type { ConnectionDefinition, ConnectionSnapshot, ConnectionStatus } from "../../connections/types.js";
import {
  createCallout,
  createItemCard,
  createConfigInput,
  createActionsRow,
  createSectionHeader,
  createEmptyInline,
  createButton,
} from "../../ui/extensions-hub-components.js";
import { lucide, AlertTriangle, Plug } from "../../ui/lucide-icons.js";
import { showToast } from "../../ui/toast.js";
import { formatRelativeDate } from "./overlay-relative-date.js";

// ── Badge mapping ───────────────────────────────────

interface BadgeSpec {
  text: string;
  tone: "ok" | "warn" | "muted";
}

function connectionBadge(status: ConnectionStatus): BadgeSpec {
  switch (status) {
    case "connected": return { text: "Connected", tone: "ok" };
    case "missing": return { text: "Not configured", tone: "muted" };
    case "invalid": return { text: "Invalid", tone: "warn" };
    case "error": return { text: "Error", tone: "warn" };
  }
}

// ── Connection card ─────────────────────────────────

function renderConnectionCard(args: {
  definition: ConnectionDefinition;
  snapshot: ConnectionSnapshot;
  presence: Record<string, boolean>;
  connectionManager: ConnectionManager;
}): HTMLElement {
  const { definition, snapshot, presence, connectionManager } = args;
  const badge = connectionBadge(snapshot.status);
  const needsAction = snapshot.status === "missing" || snapshot.status === "error";

  const card = createItemCard({
    icon: lucide(Plug),
    iconColor: "purple",
    name: definition.title,
    description: definition.capability,
    expandable: true,
    expanded: needsAction,
    badges: [badge],
  });

  // Error callout
  if (snapshot.lastError) {
    card.body.appendChild(
      createCallout("warn", lucide(AlertTriangle), snapshot.lastError, { compact: true }),
    );
  }

  // Secret field inputs
  const inputs = new Map<string, HTMLInputElement>();

  for (const field of definition.secretFields) {
    const fieldPresent = presence[field.id] ?? false;
    const shouldMask = field.maskInUi !== false;

    const input = createConfigInput({
      placeholder: fieldPresent ? "Saved — enter to replace" : field.label,
      type: shouldMask ? "password" : "text",
    });
    inputs.set(field.id, input);

    const row = document.createElement("div");
    row.className = "pi-item-card__config-row";

    const label = document.createElement("span");
    label.className = "pi-item-card__config-label";

    if (fieldPresent) {
      label.textContent = `${field.label} ✓`;
      label.title = "Saved";
    } else {
      label.textContent = field.label;
    }

    row.append(label, input);
    card.body.appendChild(row);
  }

  // Last validated timestamp
  if (snapshot.status === "connected" && snapshot.lastValidatedAt) {
    const meta = document.createElement("div");
    meta.className = "pi-item-card__meta";
    meta.textContent = `Last validated: ${formatRelativeDate(snapshot.lastValidatedAt)}`;
    card.body.appendChild(meta);
  }

  // Actions
  const hasSavedSecrets = Object.values(presence).some(Boolean);

  const saveBtn = createButton("Save", {
    primary: true,
    compact: true,
    onClick: () => {
      const patch: Record<string, string> = {};
      for (const [fieldId, input] of inputs) {
        const value = input.value.trim();
        if (value.length > 0) {
          patch[fieldId] = value;
        }
      }

      if (Object.keys(patch).length === 0) {
        showToast("Enter at least one field to save.");
        return;
      }

      void (async () => {
        try {
          await connectionManager.updateSecretsFromHost(definition.id, patch);
          // Clear inputs after successful save
          for (const input of inputs.values()) {
            input.value = "";
          }
          showToast(`Saved ${definition.title} credentials`);
        } catch (err: unknown) {
          showToast(`Save failed: ${err instanceof Error ? err.message : String(err)}`);
        }
      })();
    },
  });

  const clearBtn = createButton("Clear", {
    compact: true,
    onClick: () => {
      void (async () => {
        try {
          await connectionManager.clearSecretsFromHost(definition.id);
          showToast(`Cleared ${definition.title} credentials`);
        } catch (err: unknown) {
          showToast(`Clear failed: ${err instanceof Error ? err.message : String(err)}`);
        }
      })();
    },
  });

  if (hasSavedSecrets) {
    card.body.appendChild(createActionsRow(saveBtn, clearBtn));
  } else {
    card.body.appendChild(createActionsRow(saveBtn));
  }

  return card.root;
}

// ── Section renderer ────────────────────────────────

interface ExtensionListSource {
  list: () => readonly { id: string }[];
}

export async function renderExtensionConnectionsSection(args: {
  container: HTMLElement;
  connectionManager: ConnectionManager;
  extensionManager: ExtensionListSource;
}): Promise<void> {
  const { container, connectionManager, extensionManager } = args;

  // Visibility: hide when no extensions are installed
  const extensions = extensionManager.list();
  if (extensions.length === 0) return;

  const definitions = connectionManager.listDefinitions();

  container.appendChild(createSectionHeader({ label: "Extension connections" }));

  if (definitions.length === 0) {
    container.appendChild(
      createEmptyInline(lucide(Plug), "Installed extensions haven't registered any connections."),
    );
    return;
  }

  const list = document.createElement("div");
  list.className = "pi-hub-stack";

  for (const definition of definitions) {
    const [snapshot, presence] = await Promise.all([
      connectionManager.getSnapshot(definition.id),
      connectionManager.getSecretFieldPresence(definition.id),
    ]);

    if (!snapshot) continue;

    const cardEl = renderConnectionCard({
      definition,
      snapshot,
      presence,
      connectionManager,
    });
    list.appendChild(cardEl);
  }

  container.appendChild(list);
}
