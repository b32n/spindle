import JSZip from 'jszip';
import { kids, parseXml, type XmlElement } from '../xlsx/xml';

// ---------------------------------------------------------------------------
// A safe, read-only view of an OOXML zip package for the loss detectors: every read is bounded by the actual
// inflated size (not by what the zip headers claim), parts are parsed once, and nothing here throws. The XLSX
// detector has its own copy of the stream reader (xlsx/loss.ts); this one is for the formats that came after it.
// ---------------------------------------------------------------------------

/**
 * Safety limits. A part over `partBytes`, or one that would take the total over `totalBytes` (uncompressed), is not read.
 * A package with more than `entries` entries is not scanned at all.
 */
export const PACKAGE_LIMITS = { partBytes: 64 * 1024 * 1024, totalBytes: 256 * 1024 * 1024, entries: 50_000 };

export interface Rel {
  id: string;
  /** Last segment of the relationship type URI: `slide`, `image`, `hyperlink`, `comments`... */
  type: string;
  /** Resolved part name, or the raw URL for an external target. */
  target: string;
  external: boolean;
  /** The `Target` exactly as written. */
  raw: string;
}

/** Resolve a relationship target against the part that owns it. A target that climbs above the root stays at the root. */
export function resolveTarget(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

/** The part that holds a part's relationships: `ppt/slides/slide1.xml` -> `ppt/slides/_rels/slide1.xml.rels`. */
export const relsPart = (part: string): string => {
  const slash = part.lastIndexOf('/');
  return `${part.slice(0, slash + 1)}_rels/${part.slice(slash + 1)}.rels`;
};

/** The part of JSZip's (untyped) streaming helper used here. */
interface Stream {
  on(event: 'data', fn: (chunk: Uint8Array) => void): Stream;
  on(event: 'error' | 'end', fn: () => void): Stream;
  pause(): Stream;
  resume(): Stream;
}

/** Inflate a zip entry, giving up (`'limit'`) as soon as more than `limit` bytes have come out. `undefined` when it cannot be read. */
function inflate(file: JSZip.JSZipObject, limit: number): Promise<Uint8Array | 'limit' | undefined> {
  return new Promise((resolve) => {
    const chunks: Uint8Array[] = [];
    let total = 0;
    let done = false;
    const finish = (value: Uint8Array | 'limit' | undefined) => {
      if (done) return;
      done = true;
      resolve(value);
    };
    const stream = (file as unknown as { internalStream(type: 'uint8array'): Stream }).internalStream('uint8array');
    stream
      .on('data', (chunk: Uint8Array) => {
        total += chunk.length;
        if (total > limit) {
          stream.pause();
          finish('limit');
        } else chunks.push(chunk);
      })
      .on('error', () => finish(undefined))
      .on('end', () => {
        const out = new Uint8Array(total);
        let at = 0;
        for (const c of chunks) {
          out.set(c, at);
          at += c.length;
        }
        finish(out);
      })
      .resume();
  });
}

export class OoxmlPackage {
  /** Uncompressed bytes read so far, against `PACKAGE_LIMITS.totalBytes`. */
  read = 0;
  /** Parts that were not read because of the limits. */
  readonly skipped: string[] = [];
  private readonly parsed = new Map<string, XmlElement | undefined>();
  private readonly relCache = new Map<string, Rel[]>();

  constructor(private readonly zip: JSZip) {}

  /** The names of the file entries (directories excluded), in archive order. */
  get names(): string[] {
    return Object.keys(this.zip.files).filter((n) => !n.endsWith('/'));
  }

  has(name: string): boolean {
    return this.zip.file(name) !== null;
  }

  /**
   * The text of a part, or `undefined` when it is absent, over the limits or unreadable. Never throws. The declared size
   * is a cheap first check; the inflated bytes are counted as they come out and the stream is stopped at the limit, so a
   * forged header cannot make a bomb inflate.
   */
  async text(name: string): Promise<string | undefined> {
    const file = this.zip.file(name);
    if (!file) return undefined;
    const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
    const limit = Math.min(PACKAGE_LIMITS.partBytes, PACKAGE_LIMITS.totalBytes - this.read);
    if (typeof size === 'number' && size > limit) {
      this.skipped.push(name);
      return undefined;
    }
    const bytes = await inflate(file, limit);
    if (bytes === 'limit') this.skipped.push(name);
    if (typeof bytes === 'string' || bytes === undefined) return undefined;
    this.read += bytes.length;
    return new TextDecoder().decode(bytes);
  }

  /** A parsed part; each part is read, counted against the budget and parsed once however many parts refer to it. */
  async xml(name: string): Promise<XmlElement | undefined> {
    if (this.parsed.has(name)) return this.parsed.get(name);
    const text = await this.text(name);
    const root = text === undefined ? undefined : (parseXml(text) ?? undefined);
    this.parsed.set(name, root);
    return root;
  }

  /** The relationships of a part (empty when it has none or its `.rels` cannot be read). */
  async rels(part: string): Promise<Rel[]> {
    const cached = this.relCache.get(part);
    if (cached) return cached;
    const root = await this.xml(relsPart(part));
    const rels = kids(root, 'Relationship').map((r) => {
      const external = r.attrs.TargetMode === 'External';
      const target = r.attrs.Target ?? '';
      const type = r.attrs.Type ?? '';
      return {
        id: r.attrs.Id ?? '',
        type: type.slice(type.lastIndexOf('/') + 1),
        target: external ? target : resolveTarget(part, target),
        external,
        raw: target,
      };
    });
    this.relCache.set(part, rels);
    return rels;
  }
}
