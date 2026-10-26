import { eslintCompatPlugin } from "@oxlint/plugins";

import { exportedFunctionReturnTypesRule } from "./rules/exported-function-return-types.ts";
import { noAuthLocalStorageRule } from "./rules/no-auth-local-storage.ts";
import { noTypedJsonBoundaryCastRule } from "./rules/no-typed-json-boundary-cast.ts";

/** Repository-specific Oxlint rules. Scoped to their directories in `.oxlintrc.json`. */
const piForExcelPlugin = eslintCompatPlugin({
  meta: { name: "pi-for-excel" },
  rules: {
    "exported-function-return-types": exportedFunctionReturnTypesRule,
    "no-auth-local-storage": noAuthLocalStorageRule,
    "no-typed-json-boundary-cast": noTypedJsonBoundaryCastRule,
  },
});

export default piForExcelPlugin;
