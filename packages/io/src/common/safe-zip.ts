import JSZip from 'jszip';
import { isEncryptedOle, isOleFile } from '../xlsx/package-kind';
import { ImportError } from './import-error';

// ---------------------------------------------------------------------------
// Opening an OOXML package safely. `openSafeZip` either returns the zip or throws an `ImportError` (and nothing else):
//
//   1. container sniffing: empty, OLE compound file (encrypted package or legacy binary), OpenDocument, iWork, not a zip
//   2. the zip's own metadata: entry count (read from the end-of-central-directory record before JSZip builds anything),
//      unsafe part names, declared per-part and total sizes, declared compression ratio
//   3. a streaming guard: every entry is inflated into the void with a running byte count, and stopped the moment it
//      passes a limit, so a forged header (declaring a small size for a bomb) cannot make anything inflate for real
//
// The limits are options on the importers (host-tunable); the defaults are the SD16 proposal of the fidelity plan.
// ---------------------------------------------------------------------------

const MB = 1024 * 1024;

/**
 * The limits of every importer. Each field says which importers honour it: the first six are the safe zip's and apply to all of
 * them (DOCX, XLSX, PPTX); the rest are the XLSX reader's own budgets (`importXlsx`) and are ignored by the others.
 */
export interface ImportLimits {
  /** Size of the file itself. All importers. */
  maxPackageBytes: number;
  /** Number of zip entries (directories included). All importers. */
  maxEntries: number;
  /** Uncompressed bytes over all entries. All importers. */
  maxTotalBytes: number;
  /** Uncompressed bytes of one entry. All importers. */
  maxPartBytes: number;
  /** Uncompressed bytes of one XML entry (`.xml`, `.rels`, `.vml`). All importers. */
  maxXmlPartBytes: number;
  /** Largest compression ratio (uncompressed : compressed) of a part or of the whole package. All importers. */
  maxRatio: number;
  /**
   * The ratio is only judged for a part (or package) that inflates to more than this. A small part can only cost this
   * much memory however well it compresses, and real files do contain tiny, extremely regular parts: without a floor they
   * would be refused for nothing. All importers.
   */
  ratioFloorBytes: number;
  /**
   * XLSX only. Cost, in characters, of the formula text one worksheet may gain by expanding shared formulas (each cell of a group gets its
   * own copy of the master's text; a formula costs its length plus a fixed overhead for the string itself). Over it the remaining
   * cells of that sheet keep their cached value and lose the formula, with one warning.
   */
  maxSheetFormulaChars: number;
  /** XLSX only. The same cost over the whole package: all its worksheets together draw on it, so many sheets cannot add up past it. */
  maxFormulaChars: number;
  /**
   * XLSX only. Cells kept over all the worksheets of a package. Past it the rest of the file is not read (one warning). It also bounds
   * the shared strings read (a cell names one, so no more are of use).
   */
  maxTotalCells: number;
  /**
   * XLSX only. Worksheets read from one package. Past it the other sheets are not read (one warning). The workbook's `<sheet>` entries are
   * listed up to four times this number (chart sheets, sheets without a part and repeats of a part count as entries, not as sheets read).
   */
  maxSheets: number;
  /**
   * XLSX only. Merged ranges and hyperlinks kept over all the worksheets of a package (one budget for both, each of which is also capped per
   * sheet: 1,048,576 merges, Excel's 65,530 hyperlinks). Past it the rest are left out with one counted warning per sheet. A range
   * object is about 100 bytes, so the default is about 100 MB at most.
   */
  maxRanges: number;
  /**
   * XLSX only. Rows with their own settings (height, hidden, style, outline) kept over all the worksheets of a package: a row element
   * with none of those costs nothing, and a cell-less row that has some is an entry of about 150 bytes. Past it the settings of
   * the other rows are left out, with one counted warning per sheet (the cells are not affected). The default is two worksheets of
   * rows, 300 MB at most.
   */
  maxRows: number;
  /**
   * XLSX only. Elements (start tags) kept from one workbook-level part: the styles, the theme and the workbook itself, as element trees;
   * the entries of a relationships or content-types part; the rich-text runs of the shared strings. The default is five times the
   * most a sound file has (Excel's 64,000 cell formats and as many differential formats make about 200,000 elements, and a theme or a
   * workbook is a few hundred). A tree part over it is not read at all (one warning: its readers give their defaults), a list part
   * keeps its first entries, runs past it are joined into their string's text without fonts. The memory a part can cost is
   * about 420 bytes per element (measured, Node 22), so the default bounds each part at about 420 MB, however long the part is.
   */
  maxPartElements: number;
  /**
   * XLSX only. Column settings kept over all the worksheets of a package: each `<col>` range read costs one, and so does each column
   * of a sheet that a range gives a width or hides (up to the sheet's last cell), as those are entries of the result too. A full-width
   * `<col>` (16,384 columns) on a sheet with cells to column T costs 21, but one on a sheet with a cell in `XFD` costs 16,385. Past it
   * the rest are left out, with one warning (the sheets that lose some are in the import report). An entry is about 100 bytes, so the default is about 100 MB at most.
   */
  maxColumns: number;
  /**
   * XLSX only. Characters of raw extension lists (`extLst`, kept for the sheet's `details`) over all the worksheets of a package (each
   * is also capped at 8 Mi characters). Past it the extension lists that do not fit are not kept, with one warning.
   * Characters are 2 bytes at most, so the default is 128 MB at most.
   */
  maxExtensionChars: number;
  /**
   * XLSX only. Relationship entries read over all the parts of a package (the workbook's and each sheet's `.rels`; each part is also
   * capped at `maxPartElements`). Only the relationships the reader uses are kept (those of sheets, hyperlinks, comments, drawings and
   * the workbook-level parts), but every entry is read, so this bounds the time too. Past it the rest of that part and the
   * relationships of the other parts are not read, with one warning.
   */
  maxRelationships: number;
  /**
   * XLSX only. Cell (`<c>`), row (`<row>`) and column (`<col>`) elements scanned over all the worksheets of a package, whether they are kept or not (a
   * cell with no value is not kept, and a part of empty `<c/>` costs time and nothing else; the cells kept are also bounded by
   * `maxTotalCells`). Scanning costs about a microsecond an element, so the default is about 15 seconds at most. Past it the rest of that
   * sheet and the sheets after it are not read, with one warning.
   */
  maxScannedElements: number;
}

/**
 * Defaults: package 100 MB, 20,000 entries, 1 GB in total, 500 MB per part (256 MB for XML), ratio 200:1 above 16 MB,
 * expanded shared-formula text of 32 Mi characters per worksheet and 64 Mi (128 MB of strings) per package, 5,000,000 cells and
 * 1,000 worksheets and 1,048,576 merged ranges and hyperlinks per package, 2,097,152 rows with settings and 1,000,000 elements
 * per workbook-level part, 1,048,576 column settings, 64 Mi characters of raw extension lists, 1,048,576 relationship entries and
 * 16,777,216 scanned cell and row elements per package. The 100,000 x 10 performance file (1,000,000 cells) is
 * five times under the cell limit. Memory, as measured for the reader alone (`importXlsx`, Node 24, a 4 GB heap): about 0.17 GB of
 * process memory (RSS) per million cells of plain numbers, and 5,000,000 cells took 0.86 GB (numbers), 1.04 GB (shared strings), 1.22 GB
 * (formulas) and 1.0 GB (with 1,048,576 merges over them: the cells exist before the merged ones are dropped) in the shapes measured
 * for this reader, and up to 2.0 GB in the review's variants (1.67 GB of formulas, 1.82 GB of shared strings, 1.96 GB of merges,
 * measured before the strings it keeps were copied): plan for 2 GB. The 1.1 GB that the `--runtime` run of
 * `npm run sheets:xlsx-fidelity` reports for the 100,000 x 10 file is that whole harness (loss detection and mapping included), not the
 * reader. A host that also builds the editor's own model from the result (a second copy of every cell), and any browser host (a tab is
 * commonly limited to 2 to 4 GB), should lower `maxTotalCells`: 1,000,000 keeps the 100,000 x 10 file whole at a fifth of the cost. A sheet of two million shared-formula cells of 15
 * characters is under the formula limits.
 */
export const DEFAULT_IMPORT_LIMITS: ImportLimits = {
  maxPackageBytes: 100 * MB,
  maxEntries: 20_000,
  maxTotalBytes: 1024 * MB,
  maxPartBytes: 500 * MB,
  maxXmlPartBytes: 256 * MB,
  maxRatio: 200,
  ratioFloorBytes: 16 * MB,
  maxSheetFormulaChars: 32 * MB,
  maxFormulaChars: 64 * MB,
  maxTotalCells: 5_000_000,
  maxSheets: 1_000,
  maxRanges: 1_048_576,
  maxRows: 2_097_152,
  maxPartElements: 1_000_000,
  maxColumns: 1_048_576,
  maxExtensionChars: 64 * 1024 * 1024,
  maxRelationships: 1_048_576,
  maxScannedElements: 16_777_216,
};

export interface SafeZipOptions {
  limits?: Partial<ImportLimits>;
  /** What the caller wants, for messages: "an Excel workbook (.xlsx)". */
  target?: string;
}

export const resolveLimits = (limits?: Partial<ImportLimits>): ImportLimits => ({ ...DEFAULT_IMPORT_LIMITS, ...limits });

const mb = (n: number): string => (n >= MB ? `${Math.round(n / MB)} MB` : `${n} bytes`);

// ---- container sniffing -----------------------------------------------------------------------------------------

const ZIP_MAGIC = [0x50, 0x4b];
const startsWith = (bytes: Uint8Array, magic: number[]) => bytes.length >= magic.length && magic.every((b, i) => bytes[i] === b);

/** `PK\x03\x04` (an entry) or `PK\x05\x06` (an empty archive). */
const isZip = (bytes: Uint8Array): boolean => startsWith(bytes, ZIP_MAGIC) && (bytes[2] === 3 || bytes[2] === 5) && (bytes[3] === 4 || bytes[3] === 6);

const ODF_TYPE = /^application\/vnd\.oasis\.opendocument\./;
/** Longest `mimetype` entry looked at (an ODF marker is ~60 bytes); a longer one is not a marker. */
const MIMETYPE_MAX_BYTES = 255;
const IWORK_PART = /^(?:Index\/[^/]*\.iwa|Index\.zip|index\.apxl|index\.xml\.gz)$/;

// ---- the zip end-of-central-directory record ---------------------------------------------------------------------

/** How many entries the zip declares, read from its last bytes (zip64-aware), or `undefined` when there is no such record. */
export function declaredEntryCount(bytes: Uint8Array): number | undefined {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.length);
  const oldest = Math.max(0, bytes.length - 22 - 0xffff);
  for (let at = bytes.length - 22; at >= oldest; at--) {
    if (view.getUint32(at, true) !== 0x06054b50) continue;
    const total = view.getUint16(at + 10, true);
    if (total !== 0xffff || at < 20 || view.getUint32(at - 20, true) !== 0x07064b50) return total;
    // zip64: the real count is in the zip64 end record the locator points to
    const offset = Number(view.getBigUint64(at - 20 + 8, true));
    if (offset + 56 > bytes.length || view.getUint32(offset, true) !== 0x06064b50) return total;
    return Number(view.getBigUint64(offset + 32, true));
  }
  return undefined;
}

// ---- part names ---------------------------------------------------------------------------------------------------

/** A name that could leave the extraction root if anyone wrote it to disk: `..`, absolute, backslash, drive letter, NUL. */
export function isUnsafePartName(name: string): boolean {
  return name.includes('\\') || name.includes('\0') || name.startsWith('/') || /^[A-Za-z]:/.test(name) || name.split('/').includes('..');
}

const isXmlPart = (name: string): boolean => /\.(?:xml|rels|vml)$/i.test(name);

// ---- streaming guard ----------------------------------------------------------------------------------------------

/** The part of JSZip's (untyped) streaming helper used here. */
interface Stream {
  on(event: 'data', fn: (chunk: Uint8Array) => void): Stream;
  on(event: 'error' | 'end', fn: () => void): Stream;
  pause(): Stream;
  resume(): Stream;
}

/** The sizes JSZip read from the central directory (set on every entry of a loaded zip, directories included). */
interface EntryData {
  uncompressedSize: number;
  compressedSize: number;
}
const entryData = (file: JSZip.JSZipObject): EntryData => (file as unknown as { _data: EntryData })._data;

/** The limit that was passed: what it limits and the limit as text. */
interface Hit {
  what: string;
  limit: string;
}

/** Inflate an entry, handing each chunk to `stop(total, chunk)` and stopping as soon as it says so (the bytes are otherwise thrown away). Resolves to the bytes counted and why it stopped. */
function drain(file: JSZip.JSZipObject, stop: (total: number, chunk: Uint8Array) => Hit | undefined): Promise<{ total: number; stopped?: Hit; failed?: boolean }> {
  return new Promise((resolve) => {
    let total = 0;
    let done = false;
    const finish = (extra: { stopped?: Hit; failed?: boolean }) => {
      if (done) return;
      done = true;
      resolve({ total, ...extra });
    };
    const stream = (file as unknown as { internalStream(type: 'uint8array'): Stream }).internalStream('uint8array');
    stream
      .on('data', (chunk: Uint8Array) => {
        total += chunk.length;
        const stopped = stop(total, chunk);
        if (stopped) {
          stream.pause();
          finish({ stopped });
        }
      })
      .on('error', () => finish({ failed: true }))
      .on('end', () => finish({}))
      .resume();
  });
}

// ---- the entry point -----------------------------------------------------------------------------------------------

const tooBig = (what: string, limit: string) =>
  new ImportError('size-limit', `The file is too large to open safely (${what} is over the limit of ${limit}).`);

/**
 * Open `bytes` as a zip package, or throw an `ImportError`. Only returns a zip that passed every check above; reading from it
 * afterwards is bounded by the same limits (each entry was seen to inflate to what it declares, within the limits).
 */
export async function openSafeZip(bytes: Uint8Array, options: SafeZipOptions = {}): Promise<JSZip> {
  const limits = resolveLimits(options.limits);
  const target = options.target ?? 'an Open XML file (.xlsx, .pptx or .docx)';
  if (bytes.length === 0) throw new ImportError('not-a-package', 'The file is empty.');
  if (bytes.length > limits.maxPackageBytes) throw tooBig(`the file (${mb(bytes.length)})`, mb(limits.maxPackageBytes));
  if (isOleFile(bytes)) {
    if (isEncryptedOle(bytes))
      throw new ImportError('encrypted', 'This file is password-protected. Remove the password in the app that made it, save it again and open the new copy.');
    throw new ImportError('legacy-binary', `This is a legacy binary Office file (.xls, .ppt or .doc), which cannot be opened. Save it as ${target} first.`);
  }
  if (!isZip(bytes)) throw new ImportError('not-a-package', `This is not ${target}.`);

  const declared = declaredEntryCount(bytes);
  if (declared !== undefined && declared > limits.maxEntries) throw tooBig(`the number of entries (${declared})`, String(limits.maxEntries));

  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch (e) {
    if (/encrypted/i.test((e as Error).message)) throw new ImportError('encrypted', 'This file is password-protected. Remove the password, save it again and open the new copy.');
    throw new ImportError('corrupt', 'The file is damaged (it is not a complete zip archive).');
  }

  const names = Object.keys(zip.files);
  if (names.length > limits.maxEntries) throw tooBig(`the number of entries (${names.length})`, String(limits.maxEntries));
  const unsafe = names.find(isUnsafePartName);
  if (unsafe !== undefined) throw new ImportError('corrupt', 'The file is damaged or unsafe (it contains a part with an invalid path).');

  await rejectOtherFormats(zip, names, target);
  await guardInflate(zip, names, bytes.length, limits);
  return zip;
}

/** OpenDocument, iWork and binary-workbook packages are zips too: say what they are instead of "damaged". */
async function rejectOtherFormats(zip: JSZip, names: string[], target: string): Promise<void> {
  const mimetype = zip.file('mimetype');
  if (mimetype) {
    // read through the same bounded stream as the guard, whatever the header declares: an ODF marker is a few dozen bytes
    const chunks: number[] = [];
    const { stopped, failed } = await drain(mimetype, (total, chunk) => {
      if (total > MIMETYPE_MAX_BYTES) return { what: 'the mimetype entry', limit: `${MIMETYPE_MAX_BYTES} bytes` };
      chunks.push(...chunk);
      return undefined;
    });
    if (failed) throw new ImportError('corrupt', 'The file is damaged (a part cannot be read).');
    const type = stopped ? '' : new TextDecoder().decode(Uint8Array.from(chunks)).trim();
    if (ODF_TYPE.test(type)) throw new ImportError('not-a-package', `This is an OpenDocument file (.ods, .odp or .odt). Save it as ${target} first.`);
  }
  if (names.some((n) => IWORK_PART.test(n))) throw new ImportError('not-a-package', `This is an Apple Numbers, Keynote or Pages file. Export it as ${target} first.`);
  if (names.includes('xl/workbook.bin')) throw new ImportError('legacy-binary', `This is a binary Excel workbook (.xlsb), which cannot be opened. Save it as ${target} first.`);
  if (!names.includes('[Content_Types].xml')) throw new ImportError('not-a-package', `This is not ${target} (it has no [Content_Types].xml).`);
}

/**
 * Check the declared sizes, then inflate every entry for real against the same limits. JSZip itself fails a part whose inflated
 * size differs from the declared one, so once every part has been drained its declared sizes (and so the aggregate checks above)
 * are the true ones.
 */
async function guardInflate(zip: JSZip, names: string[], packageBytes: number, limits: ImportLimits): Promise<void> {
  const partLimit = (name: string) => (isXmlPart(name) ? limits.maxXmlPartBytes : limits.maxPartBytes);
  const ratioExceeded = (uncompressed: number, compressed: number) =>
    uncompressed > limits.ratioFloorBytes && uncompressed > limits.maxRatio * Math.max(compressed, 1);

  let declaredTotal = 0;
  for (const name of names) {
    const { uncompressedSize, compressedSize } = entryData(zip.files[name]);
    declaredTotal += uncompressedSize;
    if (uncompressedSize > partLimit(name)) throw tooBig(`a part (${mb(uncompressedSize)})`, mb(partLimit(name)));
    if (ratioExceeded(uncompressedSize, compressedSize)) throw tooBig('the compression ratio of a part', `${limits.maxRatio}:1`);
  }
  if (declaredTotal > limits.maxTotalBytes) throw tooBig(`the content (${mb(declaredTotal)})`, mb(limits.maxTotalBytes));
  if (ratioExceeded(declaredTotal, packageBytes)) throw tooBig('the compression ratio', `${limits.maxRatio}:1`);

  let inflated = 0;
  for (const name of names) {
    const file = zip.files[name];
    if (file.dir) continue;
    const compressed = entryData(file).compressedSize;
    const { total, stopped, failed } = await drain(file, (n) => {
      if (n > partLimit(name)) return { what: 'a part', limit: mb(partLimit(name)) };
      if (inflated + n > limits.maxTotalBytes) return { what: 'the content', limit: mb(limits.maxTotalBytes) };
      if (ratioExceeded(n, compressed)) return { what: 'the compression ratio of a part', limit: `${limits.maxRatio}:1` };
      return undefined;
    });
    if (stopped) throw tooBig(stopped.what, stopped.limit);
    if (failed) throw new ImportError('corrupt', 'The file is damaged (a part cannot be read).');
    inflated += total;
  }
}
