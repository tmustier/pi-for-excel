/**
 * Boot — runs before any pi-web-ui components mount.
 *
 * 1. Imports Tailwind CSS (pi-web-ui/app.css)
 * 2. Patches Lit's ReactiveElement to fix tsgo class field shadowing
 *
 * MUST be imported as the first module in taskpane.ts.
 */

import "@mariozechner/pi-web-ui/app.css";
import "./ui/theme.css";

import { ReactiveElement } from "lit";

/**
 * Fix: Lit class field shadowing (tsgo bug)
 *
 * pi-web-ui is compiled with tsgo which emits native class field declarations
 * despite useDefineForClassFields:false. Native class fields use [[Define]]
 * semantics, creating own properties that shadow Lit's @state() / @property()
 * prototype accessors. Lit's dev-mode check in performUpdate() throws.
 *
 * Fix: monkey-patch ReactiveElement.prototype.performUpdate to auto-delete
 * shadowed properties before the first update. ~15 lines, handles ALL Lit
 * components.
 *
 * See: https://lit.dev/msg/class-field-shadowing
 */
const _origPerformUpdate = (ReactiveElement.prototype as any).performUpdate;
(ReactiveElement.prototype as any).performUpdate = function (this: ReactiveElement) {
  if (!this.hasUpdated) {
    const proto = Object.getPrototypeOf(this);
    for (const key of Object.getOwnPropertyNames(this)) {
      if (key.startsWith("__") || key === "renderRoot" || key === "isUpdatePending" || key === "hasUpdated") continue;
      const protoDesc = Object.getOwnPropertyDescriptor(proto, key);
      if (protoDesc && (protoDesc.get || protoDesc.set)) {
        const ownDesc = Object.getOwnPropertyDescriptor(this, key);
        if (ownDesc && !ownDesc.get && !ownDesc.set) {
          const value = (this as any)[key];
          delete (this as any)[key];
          if (protoDesc.set) {
            (this as any)[key] = value;
          }
        }
      }
    }
  }
  return _origPerformUpdate.call(this);
};
