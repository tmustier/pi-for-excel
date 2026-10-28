/**
 * Register the first-party message web components used by the taskpane UI.
 *
 * All chat rendering elements (<message-list>, <turn-group>, <user-message>,
 * <assistant-message>, <thinking-block>, <markdown-block>, <code-block>,
 * <attachment-tile>) are first-party — see
 * docs/ui-ownership.md. Importing these modules defines the custom elements.
 */

import "./messages/message-list.js";
import "./messages/turn-group.js";
import "./messages/messages.js";
import "./messages/attachment-tile.js";
import "./messages/markdown-block.js";
import "./messages/code-block.js";
import "./messages/thinking-block.js";
