import type { FilesDialogFilterValue } from "./files-dialog-filtering.js";

export function buildFilesDialogStatusMessage(args: {
  totalCount: number;
  filteredCount: number;
  selectedFilter: FilesDialogFilterValue;
  activeFilterLabel: string;
  builtinDocsCount: number;
  workspaceFilesCount: number;
}): string {
  if (args.selectedFilter === "all") {
    return `${args.totalCount} file${args.totalCount === 1 ? "" : "s"} available to the assistant.`;
  }

  return (
    `${args.filteredCount} of ${args.totalCount} file${args.totalCount === 1 ? "" : "s"} shown`
    + ` · ${args.activeFilterLabel}.`
  );
}
