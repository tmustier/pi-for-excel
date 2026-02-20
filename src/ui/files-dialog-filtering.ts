import type { WorkspaceBackendStatus, WorkspaceFileEntry } from "../files/types.js";

export interface FilesDialogBadge {
  tone: "ok" | "muted" | "info";
  label: string;
  title?: string;
}

export interface FilesDialogFolderGroup {
  name: string;
  files: WorkspaceFileEntry[];
}

export interface FilesDialogSection {
  key: string;
  label: string;
  files: WorkspaceFileEntry[];
  folders: FilesDialogFolderGroup[];
}

const YOUR_FILES_SECTION_KEY = "your-files";
const NOTES_SECTION_KEY = "notes";
const SKILLS_SECTION_KEY = "skills";
const BUILTIN_DOCS_SECTION_KEY = "built-in-docs";

export function normalizeFilesDialogFilterText(value: string): string {
  return value.trim().toLowerCase();
}

export function isFilesDialogBuiltInDoc(file: WorkspaceFileEntry): boolean {
  return file.sourceKind === "builtin-doc" || file.locationKind === "builtin-doc";
}

export function isFilesDialogConnectedFolderFile(file: WorkspaceFileEntry): boolean {
  return file.locationKind === "native-directory";
}

export function isAgentWrittenNotesFilePath(path: string): boolean {
  const lowerPath = path.toLowerCase();
  return lowerPath.startsWith("notes/") && lowerPath.endsWith(".md");
}

export function resolveFilesDialogBadge(file: WorkspaceFileEntry): FilesDialogBadge | null {
  if (isFilesDialogBuiltInDoc(file)) {
    return { tone: "muted", label: "Read only" };
  }

  if (file.workbookTag) {
    const workbookLabel = file.workbookTag.workbookLabel.trim();
    if (workbookLabel.length > 0) {
      return {
        tone: "muted",
        label: "Workbook",
        title: `Tagged to ${workbookLabel}`,
      };
    }

    return { tone: "muted", label: "Workbook" };
  }

  if (isAgentWrittenNotesFilePath(file.path)) {
    return { tone: "muted", label: "Agent" };
  }

  if (isFilesDialogConnectedFolderFile(file)) {
    return { tone: "info", label: "Folder" };
  }

  return null;
}

export function resolveFilesDialogSourceLabel(file: WorkspaceFileEntry): string {
  if (isFilesDialogBuiltInDoc(file)) {
    return "Pi documentation";
  }

  if (isAgentWrittenNotesFilePath(file.path)) {
    return "Written by agent";
  }

  if (isFilesDialogConnectedFolderFile(file)) {
    return "Local file";
  }

  return "Uploaded";
}

export function fileMatchesFilesDialogFilter(args: {
  file: WorkspaceFileEntry;
  filterText: string;
}): boolean {
  const query = normalizeFilesDialogFilterText(args.filterText);
  if (query.length === 0) {
    return true;
  }

  return args.file.path.toLowerCase().includes(query);
}

export function filterFilesDialogEntries(args: {
  files: readonly WorkspaceFileEntry[];
  filterText: string;
}): WorkspaceFileEntry[] {
  const query = normalizeFilesDialogFilterText(args.filterText);
  if (query.length === 0) {
    return [...args.files];
  }

  return args.files.filter((file) => file.path.toLowerCase().includes(query));
}

function sortByModifiedAtDescending(files: readonly WorkspaceFileEntry[]): WorkspaceFileEntry[] {
  return [...files].sort((left, right) => {
    if (left.modifiedAt !== right.modifiedAt) {
      return right.modifiedAt - left.modifiedAt;
    }

    return left.path.localeCompare(right.path);
  });
}

function connectedFolderSectionLabel(backendStatus: WorkspaceBackendStatus | null): string {
  const folderName = backendStatus?.nativeDirectoryName?.trim();
  if (!folderName) {
    return "FROM CONNECTED FOLDER";
  }

  return `FROM ${folderName.toUpperCase()}`;
}

function connectedFolderSectionKey(backendStatus: WorkspaceBackendStatus | null): string {
  const folderName = backendStatus?.nativeDirectoryName?.trim().toLowerCase();
  if (!folderName) {
    return "from-connected-folder";
  }

  return `from-${folderName}`;
}

function isNotesFile(file: WorkspaceFileEntry): boolean {
  return file.path.toLowerCase().startsWith("notes/");
}

function isSkillsFile(file: WorkspaceFileEntry): boolean {
  return file.path.toLowerCase().startsWith("skills/");
}

/**
 * Split files into root-level items and first-level subdirectory groups.
 *
 * `stripPrefix` is removed from each path before detecting the first `/`.
 * For example, stripping `"notes/"` from `"notes/index.md"` yields `"index.md"` → root.
 */
function groupByFirstDirectory(
  files: readonly WorkspaceFileEntry[],
  stripPrefix: string,
): { rootFiles: WorkspaceFileEntry[]; folders: FilesDialogFolderGroup[] } {
  const rootFiles: WorkspaceFileEntry[] = [];
  const folderMap = new Map<string, WorkspaceFileEntry[]>();

  for (const file of files) {
    const relative = stripPrefix && file.path.startsWith(stripPrefix)
      ? file.path.slice(stripPrefix.length)
      : file.path;

    const slash = relative.indexOf("/");
    if (slash < 0) {
      rootFiles.push(file);
    } else {
      const folderName = relative.slice(0, slash);
      let bucket = folderMap.get(folderName);
      if (!bucket) {
        bucket = [];
        folderMap.set(folderName, bucket);
      }
      bucket.push(file);
    }
  }

  const folders = [...folderMap.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([name, bucket]) => ({ name, files: sortByModifiedAtDescending(bucket) }));

  return { rootFiles, folders };
}

/**
 * Group skill files by skill folder name.
 *
 * Handles both `skills/<name>/…` and `skills/external/<name>/…`.
 */
function groupSkillFiles(files: readonly WorkspaceFileEntry[]): {
  rootFiles: WorkspaceFileEntry[];
  folders: FilesDialogFolderGroup[];
} {
  const rootFiles: WorkspaceFileEntry[] = [];
  const folderMap = new Map<string, WorkspaceFileEntry[]>();

  for (const file of files) {
    const lower = file.path.toLowerCase();
    let rest: string;

    if (lower.startsWith("skills/external/")) {
      rest = file.path.slice("skills/external/".length);
    } else if (lower.startsWith("skills/")) {
      rest = file.path.slice("skills/".length);
    } else {
      rootFiles.push(file);
      continue;
    }

    const slash = rest.indexOf("/");
    if (slash < 0) {
      rootFiles.push(file);
    } else {
      const name = rest.slice(0, slash);
      let bucket = folderMap.get(name);
      if (!bucket) {
        bucket = [];
        folderMap.set(name, bucket);
      }
      bucket.push(file);
    }
  }

  const folders = [...folderMap.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([name, bucket]) => ({ name, files: sortByModifiedAtDescending(bucket) }));

  return { rootFiles, folders };
}

/** Total file count across root files and all folder groups. */
export function sectionTotalCount(section: FilesDialogSection): number {
  return section.files.length + section.folders.reduce((sum, f) => sum + f.files.length, 0);
}

export function buildFilesDialogSections(args: {
  files: readonly WorkspaceFileEntry[];
  filterText: string;
  backendStatus: WorkspaceBackendStatus | null;
}): FilesDialogSection[] {
  const filtered = filterFilesDialogEntries({
    files: args.files,
    filterText: args.filterText,
  });

  // ── Classify ──
  const userFiles: WorkspaceFileEntry[] = [];
  const notesFiles: WorkspaceFileEntry[] = [];
  const skillsFiles: WorkspaceFileEntry[] = [];
  const connectedFiles: WorkspaceFileEntry[] = [];
  const builtinFiles: WorkspaceFileEntry[] = [];

  for (const file of filtered) {
    if (isFilesDialogBuiltInDoc(file)) {
      builtinFiles.push(file);
    } else if (isFilesDialogConnectedFolderFile(file)) {
      connectedFiles.push(file);
    } else if (isNotesFile(file)) {
      notesFiles.push(file);
    } else if (isSkillsFile(file)) {
      skillsFiles.push(file);
    } else {
      userFiles.push(file);
    }
  }

  const sections: FilesDialogSection[] = [];

  // ── YOUR FILES ──
  if (userFiles.length > 0) {
    const { rootFiles, folders } = groupByFirstDirectory(userFiles, "");
    sections.push({
      key: YOUR_FILES_SECTION_KEY,
      label: "YOUR FILES",
      files: sortByModifiedAtDescending(rootFiles),
      folders,
    });
  }

  // ── PI'S NOTES ──
  if (notesFiles.length > 0) {
    const { rootFiles, folders } = groupByFirstDirectory(notesFiles, "notes/");
    sections.push({
      key: NOTES_SECTION_KEY,
      label: "PI'S NOTES",
      files: sortByModifiedAtDescending(rootFiles),
      folders,
    });
  }

  // ── SKILLS ──
  if (skillsFiles.length > 0) {
    const { rootFiles, folders } = groupSkillFiles(skillsFiles);
    sections.push({
      key: SKILLS_SECTION_KEY,
      label: "SKILLS",
      files: sortByModifiedAtDescending(rootFiles),
      folders,
    });
  }

  // ── Connected folder ──
  if (connectedFiles.length > 0) {
    const { rootFiles, folders } = groupByFirstDirectory(connectedFiles, "");
    sections.push({
      key: connectedFolderSectionKey(args.backendStatus),
      label: connectedFolderSectionLabel(args.backendStatus),
      files: sortByModifiedAtDescending(rootFiles),
      folders,
    });
  }

  // ── BUILT-IN DOCS ──
  if (builtinFiles.length > 0) {
    sections.push({
      key: BUILTIN_DOCS_SECTION_KEY,
      label: "BUILT-IN DOCS",
      files: [...builtinFiles].sort((a, b) => a.path.localeCompare(b.path)),
      folders: [],
    });
  }

  return sections;
}

export interface FilesDialogConnectFolderButtonState {
  hidden: boolean;
  disabled: boolean;
  label: string;
  title: string;
}

export function resolveFilesDialogConnectFolderButtonState(
  backendStatus: WorkspaceBackendStatus | null,
): FilesDialogConnectFolderButtonState {
  if (!backendStatus || !backendStatus.nativeSupported) {
    return {
      hidden: true,
      disabled: true,
      label: "Connect folder",
      title: "",
    };
  }

  if (backendStatus.nativeConnected) {
    return {
      hidden: false,
      disabled: true,
      label: "Connected ✓",
      title: "Folder already connected",
    };
  }

  return {
    hidden: false,
    disabled: false,
    label: "Connect folder",
    title: "Connect local folder",
  };
}
