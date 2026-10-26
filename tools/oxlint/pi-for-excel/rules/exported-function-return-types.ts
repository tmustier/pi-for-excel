import { defineRule } from "@oxlint/plugins";
import type { ESTree } from "@oxlint/plugins";

type FunctionValue = ESTree.ArrowFunctionExpression | ESTree.Function;

function isUntypedFunctionValue(node: ESTree.Node | null): node is FunctionValue {
  return (node?.type === "ArrowFunctionExpression" || node?.type === "FunctionExpression") && node.returnType == null;
}

function moduleExportName(name: ESTree.ModuleExportName): string {
  return name.type === "Literal" ? String(name.value) : name.name;
}

/** Names exported through `export { a, b as c }` (local side) or `export default name`. */
function namesExportedLater(program: ESTree.Program): Set<string> {
  const names = new Set<string>();
  for (const statement of program.body) {
    if (statement.type === "ExportDefaultDeclaration" && statement.declaration.type === "Identifier") {
      names.add(statement.declaration.name);
    }
    if (statement.type === "ExportNamedDeclaration" && statement.source === null && statement.declaration === null) {
      for (const specifier of statement.specifiers) names.add(moduleExportName(specifier.local));
    }
  }
  return names;
}

/** Exported top-level functions declare return types so public contracts read without re-inference. */
export const exportedFunctionReturnTypesRule = defineRule({
  meta: {
    type: "suggestion",
    docs: { description: "Require explicit return types on exported top-level functions." },
    messages: {
      missingReturnType: "Declare a return type for exported top-level function {{name}} so future agents can read public contracts without re-inference.",
    },
  },
  create(context) {
    return {
      Program(program) {
        const exportedLater = namesExportedLater(program);
        const report = (node: ESTree.Node, name: string) => context.report({ node, messageId: "missingReturnType", data: { name } });

        for (const statement of program.body) {
          if (statement.type === "ExportDefaultDeclaration") {
            const declaration = statement.declaration;
            if (isUntypedFunctionValue(declaration)) report(declaration, "<default>");
            if (declaration.type === "FunctionDeclaration" && declaration.body !== null && declaration.returnType == null) {
              report(statement, declaration.id?.name ?? "<default>");
            }
            continue;
          }

          const directlyExported = statement.type === "ExportNamedDeclaration";
          const declaration = directlyExported ? statement.declaration : statement;
          if (declaration === null) continue;

          if (declaration.type === "FunctionDeclaration" && declaration.body !== null && declaration.returnType == null) {
            const name = declaration.id?.name ?? "<default>";
            if (directlyExported || exportedLater.has(name)) report(statement, name);
            continue;
          }

          if (declaration.type !== "VariableDeclaration") continue;
          for (const declarator of declaration.declarations) {
            if (declarator.id.type !== "Identifier" || declarator.id.typeAnnotation != null) continue;
            if (!directlyExported && !exportedLater.has(declarator.id.name)) continue;
            if (isUntypedFunctionValue(declarator.init)) report(declarator, declarator.id.name);
          }
        }
      },
    };
  },
});
