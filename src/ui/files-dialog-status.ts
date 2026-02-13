import type { FilesDialogFilterValue } from "./files-dialog-filtering.js";

export function buildFilesDialogStatusMessage(args: {
  filesExperimentEnabled: boolean;
  totalCount: number;
  filteredCount: number;
  selectedFilter: FilesDialogFilterValue;
  activeFilterLabel: string;
  builtinDocsCount: number;
  workspaceFilesCount: number;
}): string {
  if (!args.filesExperimentEnabled) {
    return (
      `Built-in docs stay available (${args.builtinDocsCount}). `
      + "Enable write access for assistant file management "
      + `(${args.workspaceFilesCount} file${args.workspaceFilesCount === 1 ? "" : "s"}).`
    );
  }

  if (args.selectedFilter === "all") {
    return `${args.totalCount} file${args.totalCount === 1 ? "" : "s"} available to the assistant.`;
  }

  return (
    `${args.filteredCount} of ${args.totalCount} file${args.totalCount === 1 ? "" : "s"} shown`
    + ` · ${args.activeFilterLabel}.`
  );
}
