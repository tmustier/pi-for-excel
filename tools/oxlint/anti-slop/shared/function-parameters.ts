import type { ESTree, SourceCode } from "@oxlint/plugins";

export type FunctionParameter = ESTree.ParamPattern;

export function containsUnknownType(type: ESTree.TSType): boolean {
	if (type.type === "TSUnknownKeyword") return true;
	if (type.type === "TSParenthesizedType") return containsUnknownType(type.typeAnnotation);
	return type.type === "TSUnionType" && type.types.some(containsUnknownType);
}

export function functionParameterTypeAnnotation(
	parameter: FunctionParameter,
): ESTree.TSTypeAnnotation | null | undefined {
	if (parameter.type === "TSParameterProperty") {
		return functionParameterTypeAnnotation(parameter.parameter);
	}
	if (parameter.type === "RestElement") {
		return parameter.typeAnnotation ?? functionParameterTypeAnnotation(parameter.argument);
	}
	if (parameter.type === "AssignmentPattern") {
		return parameter.typeAnnotation ?? functionParameterTypeAnnotation(parameter.left);
	}
	return parameter.typeAnnotation;
}

export function functionParameterBindingName(
	parameter: FunctionParameter,
	sourceCode: SourceCode,
): string {
	if (parameter.type === "TSParameterProperty") {
		return functionParameterBindingName(parameter.parameter, sourceCode);
	}
	if (parameter.type === "AssignmentPattern") {
		return functionParameterBindingName(parameter.left, sourceCode);
	}
	if (parameter.type === "RestElement") {
		return functionParameterBindingName(parameter.argument, sourceCode);
	}
	if (parameter.type === "Identifier") return parameter.name;

	const sourceText = sourceCode.getText(parameter);
	const annotationStart = parameter.typeAnnotation?.start;
	return annotationStart === undefined
		? sourceText
		: sourceText.slice(0, annotationStart - parameter.start).trimEnd();
}
