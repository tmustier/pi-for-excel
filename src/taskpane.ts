/**
 * Pi for Excel — taskpane entrypoint.
 *
 * Keep this file as thin as possible.
 *
 * MUST import `./boot.js` first:
 * - loads the first-party theme CSS
 * - installs compat patches (marked safety, theme mode, …)
 */

// MUST be first
import "./boot.js";

// Register first-party web components we rely on.
import "./ui/register-components.js";

// Custom tool + message renderers (Excel tools return markdown)
import "./ui/tool-renderers.js";
import "./ui/message-renderers.js";

import { bootstrapTaskpane } from "./taskpane/bootstrap.js";

bootstrapTaskpane();
