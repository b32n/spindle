/**
 * Why a file could not be imported. A stable code (hosts may switch on it) and a message fit to show a user as it stands.
 *
 * - `encrypted`: a password-protected package (an OLE container holding an encrypted OOXML package)
 * - `legacy-binary`: a pre-2007 binary file (`.xls`, `.ppt`, `.doc`) or `.xlsb`
 * - `not-a-package`: not an OOXML package at all (empty, another format, another office suite's file)
 * - `size-limit`: over one of the import limits (package, entries, part, total, compression ratio)
 * - `corrupt`: damaged, truncated, or holding an unsafe part name
 */
export type ImportErrorCode = 'encrypted' | 'legacy-binary' | 'not-a-package' | 'size-limit' | 'corrupt';

export class ImportError extends Error {
  readonly code: ImportErrorCode;

  constructor(code: ImportErrorCode, message: string) {
    super(message);
    this.name = 'ImportError';
    this.code = code;
  }
}

/** The error `importer` callers get for anything that went wrong: an `ImportError` is passed through, anything else becomes `corrupt`. */
export function toImportError(e: unknown): ImportError {
  if (e instanceof ImportError) return e;
  const detail = e instanceof Error && e.message ? e.message : String(e);
  return new ImportError('corrupt', `The file could not be read (${detail}).`);
}
