/**
 * Lit class field shadowing patch (tsgo).
 *
 * pi-web-ui is compiled with tsgo which emits native class field declarations
 * despite useDefineForClassFields:false. Native class fields use [[Define]]
 * semantics, creating own properties that shadow Lit's @state() / @property()
 * prototype accessors. Lit's dev-mode check in performUpdate() throws.
 *
 * Fix: monkey-patch ReactiveElement.prototype.performUpdate to auto-delete
 * shadowed properties before the first update. This is ~15 lines and applies
 * to all Lit components.
 *
 * See: https://lit.dev/msg/class-field-shadowing
 */

import { ReactiveElement } from "lit";

type PerformUpdateFn = (this: ReactiveElement) => unknown;

type ReactiveElementProto = {
  performUpdate: PerformUpdateFn;
};

let _installed = false;

export function installLitClassFieldShadowingPatch(): void {
  if (_installed) return;
  _installed = true;

  const reactiveProto = ReactiveElement.prototype as unknown as ReactiveElementProto;
  const orig = reactiveProto.performUpdate;

  reactiveProto.performUpdate = function (this: ReactiveElement) {
    if (!this.hasUpdated) {
      const proto = Object.getPrototypeOf(this);
      const self = this as unknown as Record<string, unknown>;

      for (const key of Object.getOwnPropertyNames(this)) {
        if (
          key.startsWith("__") ||
          key === "renderRoot" ||
          key === "isUpdatePending" ||
          key === "hasUpdated"
        ) {
          continue;
        }

        const protoDesc = Object.getOwnPropertyDescriptor(proto, key);
        if (protoDesc && (protoDesc.get || protoDesc.set)) {
          const ownDesc = Object.getOwnPropertyDescriptor(this, key);
          // Own, data-only property shadows a proto accessor.
          if (ownDesc && !ownDesc.get && !ownDesc.set) {
            const value = self[key];
            delete self[key];
            if (protoDesc.set) {
              self[key] = value;
            }
          }
        }
      }
    }

    return orig.call(this);
  };
}
