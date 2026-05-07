/**
 * Register third-party web components used by the Excel taskpane UI.
 *
 * IMPORTANT:
 * - Do NOT `import "@earendil-works/pi-web-ui"`.
 *   The package root is a wide barrel export that pulls in optional UI
 *   features (ChatPanel, artifacts, attachments, etc.). Many of those modules
 *   register custom elements at import time and bring heavy dependencies,
 *   defeating tree-shaking and bloating the taskpane bundle.
 *
 * - Instead, deep-import only the specific components we render.
 */

// Message list + streaming container (used by <pi-sidebar>)
import { MessageList } from "@earendil-works/pi-web-ui/dist/components/MessageList.js";
import { StreamingMessageContainer } from "@earendil-works/pi-web-ui/dist/components/StreamingMessageContainer.js";

// Registers <user-message>, <assistant-message>, <tool-message>, etc.
import "@earendil-works/pi-web-ui/dist/components/Messages.js";

// Registers <attachment-tile> (rendered for user-with-attachments messages)
import "@earendil-works/pi-web-ui/dist/components/AttachmentTile.js";

// Force module evaluation (customElements.define side effects).
void MessageList;
void StreamingMessageContainer;

// mini-lit elements used in pi-web-ui message components + our renderers.
import "@mariozechner/mini-lit/dist/MarkdownBlock.js";
import "@mariozechner/mini-lit/dist/CodeBlock.js";
