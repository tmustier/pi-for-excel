/** Typed boundary adapter for the Office document file capability. */

export interface OfficeAsyncErrorLike {
  message?: unknown;
}

export interface OfficeAsyncResultLike<T> {
  status?: unknown;
  value?: T;
  error?: OfficeAsyncErrorLike;
}

export interface OfficeSliceLike {
  data: unknown;
}

export interface OfficeFileLike {
  size: number;
  sliceCount: number;
  getSliceAsync: (
    sliceIndex: number,
    callback?: (result: OfficeAsyncResultLike<OfficeSliceLike>) => void,
  ) => void;
  closeAsync: (callback?: (result: OfficeAsyncResultLike<void>) => void) => void;
}

interface OfficeDocumentFileBoundary {
  getFileAsync?: OfficeDocumentFileAdapter["getFileAsync"];
}

interface OfficeContextFileBoundary {
  document?: unknown;
}

interface OfficeFileBoundary {
  context?: unknown;
}

export interface OfficeDocumentFileAdapter {
  getFileAsync: (
    fileType: string,
    options: { sliceSize?: number },
    callback?: (result: OfficeAsyncResultLike<OfficeFileLike>) => void,
  ) => void;
}

function parseOfficeDocumentFileAdapter(rawOffice: unknown): OfficeDocumentFileAdapter | null {
  if (typeof rawOffice !== "object" || rawOffice === null || Array.isArray(rawOffice)) return null;
  const office = rawOffice as OfficeFileBoundary;

  if (typeof office.context !== "object" || office.context === null || Array.isArray(office.context)) {
    return null;
  }
  const context = office.context as OfficeContextFileBoundary;

  if (typeof context.document !== "object" || context.document === null || Array.isArray(context.document)) {
    return null;
  }
  const document = context.document as OfficeDocumentFileBoundary;
  if (typeof document.getFileAsync !== "function") return null;

  // The function check above establishes the required file-capability contract.
  const fileDocument = document as OfficeDocumentFileAdapter;
  return {
    getFileAsync(fileType, options, callback): void {
      // Office document methods require their owning document as the receiver.
      fileDocument.getFileAsync(fileType, options, callback);
    },
  };
}

export function getOfficeDocumentFileAdapter(): OfficeDocumentFileAdapter | null {
  const officeRoot: unknown = typeof Office === "undefined" ? undefined : Office;
  return parseOfficeDocumentFileAdapter(officeRoot);
}
