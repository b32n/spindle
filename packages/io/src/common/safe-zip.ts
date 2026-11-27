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

export interface ImportLimits {
  /** Size of the file itself. */
  maxPackageBytes: number;
  /** Number of zip entries (directories included). */
  maxEntries: number;
  /** Uncompressed bytes over all entries. */
  maxTotalBytes: number;
  /** Uncompressed bytes of one entry. */
  maxPartBytes: number;
  /** Uncompressed bytes of one XML entry (`.xml`, `.rels`, `.vml`). */
  maxXmlPartBytes: number;
  /** Largest compression ratio (uncompressed : compressed) of a part or of the whole package. */
  maxRatio: number;
  /**
   * The ratio is only judged for a part (or package) that inflates to more than this. A small part can only cost this
   * much memory however well it compresses, and real files do contain tiny, extremely regular parts: without a floor they
   * would be refused for nothing.
   */
  ratioFloorBytes: number;
}

/** Defaults: package 100 MB, 20,000 entries, 1 GB in total, 500 MB per part (256 MB for XML), ratio 200:1 above 16 MB. */
export const DEFAULT_IMPORT_LIMITS: ImportLimits = {
  maxPackageBytes: 100 * MB,
  maxEntries: 20_000,
  maxTotalBytes: 1024 * MB,
  maxPartBytes: 500 * MB,
  maxXmlPartBytes: 256 * MB,
  maxRatio: 200,
  ratioFloorBytes: 16 * MB,
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
