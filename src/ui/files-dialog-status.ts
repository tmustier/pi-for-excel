import { t } from "../language/index.js";
import { formatBytes } from "../files/mime.js";

function formatFileCount(totalCount: number): string {
  return t("files-dialog.fileCount", { count: String(totalCount) });
}

function formatBackendLabel(args: {
  backendLabel: string;
  nativeDirectoryName?: string | null;
}): string {
  const directoryName = args.nativeDirectoryName?.trim();
  if (!directoryName) {
    return args.backendLabel;
  }

  return `${args.backendLabel}: ${directoryName}`;
}

export function buildFilesDialogStatusMessage(args: {
  totalCount: number;
  totalSizeBytes: number;
  backendLabel: string;
  nativeDirectoryName?: string | null;
}): string {
  return [
    formatFileCount(args.totalCount),
    formatBytes(args.totalSizeBytes),
    formatBackendLabel({
      backendLabel: args.backendLabel,
      ...(args.nativeDirectoryName !== undefined ? { nativeDirectoryName: args.nativeDirectoryName } : {}),
    }),
  ].join(" · ");
}
