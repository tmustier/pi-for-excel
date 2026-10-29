import { eslintCompatPlugin } from "@oxlint/plugins";

import { exportedFunctionReturnTypesRule } from "./rules/exported-function-return-types.ts";
import { noAuthLocalStorageRule } from "./rules/no-auth-local-storage.ts";
import { noModuleScopeTranslationRule } from "./rules/no-module-scope-translation.ts";
import { noNeverAssertionInTestsRule } from "./rules/no-never-assertion-in-tests.ts";
import { noTypedJsonBoundaryCastRule } from "./rules/no-typed-json-boundary-cast.ts";

const piForExcelPlugin = eslintCompatPlugin({
  meta: { name: "pi-for-excel" },
  rules: {
    "exported-function-return-types": exportedFunctionReturnTypesRule,
    "no-auth-local-storage": noAuthLocalStorageRule,
    "no-module-scope-translation": noModuleScopeTranslationRule,
    "no-never-assertion-in-tests": noNeverAssertionInTestsRule,
    "no-typed-json-boundary-cast": noTypedJsonBoundaryCastRule,
  },
});

export default piForExcelPlugin;
