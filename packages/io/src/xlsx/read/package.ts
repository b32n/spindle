import type JSZip from 'jszip';
import type { XlsxRelationship } from '@b32nio/spindle-shared/export/xlsx';
import { openSafeZip, type ImportLimits } from '../../common/safe-zip';
import { parseXml, type XmlElement } from '../xml';
import { contentTypeOf, EMPTY_CONTENT_TYPES, parseContentTypes, type ContentTypes } from './content-types';
import { parseRelationships, relsPart } from './relationships';

export interface OpenXlsxOptions {
  /** Overrides of the import limits (`DEFAULT_IMPORT_LIMITS`). */
  limits?: Partial<ImportLimits>;
}

/** Text of a part: UTF-8 (BOM removed) or, for the rare producer that writes it, UTF-16 with a byte order mark. */
export function decodePart(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  const start = bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  return new TextDecoder('utf-8').decode(bytes.subarray(start));
}

/**
 * An opened XLSX package: the safe zip (`openSafeZip`, so every limit of X4 has already been enforced), its content types,
 * and lazy, DOM-per-part access. Reading never throws: a part that is absent or not well-formed XML comes back `undefined`
 * and is recorded in `warnings`.
 *
 * Parts are read when asked for and nothing is kept (apart from the small relationship lists), so a reader can take the
 * workbook-level parts through `xml()` and leave a large `sheetData` to a streaming reader over `text()`/`bytes()` (S2b, S20).
 */
export class XlsxPackage {
  /** Non-fatal problems found while reading, in the order met. */
  readonly warnings: string[] = [];
  contentTypes: ContentTypes = EMPTY_CONTENT_TYPES;
  private readonly byLowerName = new Map<string, string>();
  private readonly relCache = new Map<string, XlsxRelationship[]>();

  private constructor(private readonly zip: JSZip) {
    for (const name of Object.keys(zip.files)) {
      if (zip.files[name].dir) continue;
      // a duplicated name differing only by case keeps the first, as the exact spelling is looked up first anyway
      if (!this.byLowerName.has(name.toLowerCase())) this.byLowerName.set(name.toLowerCase(), name);
    }
  }

  /** Open `bytes`. The only errors are `ImportError`s from `openSafeZip` (not a package, encrypted, too large, corrupt). */
  static async open(bytes: Uint8Array, options: OpenXlsxOptions = {}): Promise<XlsxPackage> {
    const zip = await openSafeZip(bytes, { limits: options.limits, target: 'an Excel workbook (.xlsx)' });
    const pkg = new XlsxPackage(zip);
    const types = await pkg.xml('[Content_Types].xml');
    if (types) pkg.contentTypes = parseContentTypes(types);
    return pkg;
  }

  /** Names of the file entries (directories excluded), in archive order. */
  get partNames(): string[] {
    return [...this.byLowerName.values()];
  }

  /** The exact entry name for `name` (a part name with or without a leading slash, any case), or `undefined`. */
  resolve(name: string): string | undefined {
    const bare = name.replace(/^\/+/, '');
    return this.zip.file(bare) ? bare : this.byLowerName.get(bare.toLowerCase());
  }

  has(name: string): boolean {
    return this.resolve(name) !== undefined;
  }

  contentType(name: string): string | undefined {
    return contentTypeOf(this.contentTypes, name);
  }

  /** The inflated bytes of a part, or `undefined` when it is absent or cannot be inflated. */
  async bytes(name: string): Promise<Uint8Array | undefined> {
    const exact = this.resolve(name);
    if (exact === undefined) return undefined;
    try {
      return await (this.zip.file(exact) as JSZip.JSZipObject).async('uint8array');
    } catch {
      this.warnings.push(`The part ${exact} could not be read.`);
      return undefined;
    }
  }

  async text(name: string): Promise<string | undefined> {
    const bytes = await this.bytes(name);
    return bytes === undefined ? undefined : decodePart(bytes);
  }

  /** A part parsed into an element tree (not cached: the caller decides how long to hold it). */
  async xml(name: string): Promise<XmlElement | undefined> {
    const text = await this.text(name);
    if (text === undefined) return undefined;
    const root = parseXml(text);
    if (root === null) {
      this.warnings.push(`The part ${name} is not well-formed XML and was skipped.`);
      return undefined;
    }
    return root;
  }

  /** The relationships of a part (`''` for the package itself): empty when it has none or its `.rels` is unreadable. */
  async rels(part: string): Promise<XlsxRelationship[]> {
    const cached = this.relCache.get(part);
    if (cached) return cached;
    const owner = part.replace(/^\/+/, '');
    const root = await this.xml(owner === '' ? '_rels/.rels' : relsPart(owner));
    const rels = parseRelationships(root, owner);
    this.relCache.set(part, rels);
    return rels;
  }
}
