import type { WorkspaceFileEntry } from "../files/types.js";

interface MutableFolderNode {
  folderName: string;
  folderPath: string;
  files: WorkspaceFileEntry[];
  children: Map<string, MutableFolderNode>;
  totalFileCount: number;
}

export interface FilesDialogFolderNode {
  folderName: string;
  folderPath: string;
  files: readonly WorkspaceFileEntry[];
  children: readonly FilesDialogFolderNode[];
  totalFileCount: number;
}

export interface FilesDialogTree {
  rootFiles: readonly WorkspaceFileEntry[];
  folders: readonly FilesDialogFolderNode[];
}

function createMutableFolderNode(folderName: string, folderPath: string): MutableFolderNode {
  return {
    folderName,
    folderPath,
    files: [],
    children: new Map<string, MutableFolderNode>(),
    totalFileCount: 0,
  };
}

function sortFiles(files: readonly WorkspaceFileEntry[]): WorkspaceFileEntry[] {
  return [...files].sort((left, right) => left.path.localeCompare(right.path));
}

function finalizeFolders(map: ReadonlyMap<string, MutableFolderNode>): FilesDialogFolderNode[] {
  const folders = Array.from(map.values())
    .sort((left, right) => left.folderName.localeCompare(right.folderName));

  return folders.map((folder) => ({
    folderName: folder.folderName,
    folderPath: folder.folderPath,
    files: sortFiles(folder.files),
    children: finalizeFolders(folder.children),
    totalFileCount: folder.totalFileCount,
  }));
}

export function buildFilesDialogTree(files: readonly WorkspaceFileEntry[]): FilesDialogTree {
  const rootFiles: WorkspaceFileEntry[] = [];
  const rootFolders = new Map<string, MutableFolderNode>();

  const sortedFiles = sortFiles(files);

  for (const file of sortedFiles) {
    const segments = file.path.split("/");

    if (segments.length <= 1) {
      rootFiles.push(file);
      continue;
    }

    const folderSegments = segments.slice(0, -1);

    let parentMap = rootFolders;
    let parentPath = "";
    let currentFolder: MutableFolderNode | null = null;

    for (const segment of folderSegments) {
      const folderPath = parentPath.length > 0
        ? `${parentPath}/${segment}`
        : segment;

      let folder = parentMap.get(segment);
      if (!folder) {
        folder = createMutableFolderNode(segment, folderPath);
        parentMap.set(segment, folder);
      }

      folder.totalFileCount += 1;
      currentFolder = folder;
      parentMap = folder.children;
      parentPath = folderPath;
    }

    if (currentFolder) {
      currentFolder.files.push(file);
    }
  }

  return {
    rootFiles,
    folders: finalizeFolders(rootFolders),
  };
}
