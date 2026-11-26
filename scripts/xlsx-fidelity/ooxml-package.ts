import JSZip from 'jszip';

/** Generic OOXML package reader: never throws, so malformed containers can be probed too. */
export interface OoxmlParts {
  /** `zip` = readable package, `ole` = OLE compound file (encrypted or legacy binary), `unreadable` = anything else. */
  kind: 'zip' | 'ole' | 'unreadable';
  raw: Uint8Array;
  names: string[];
  /** Text of every .xml / .rels part. Parts that inflate past the size cap (by header or by actual size) or cannot be inflated are left out and listed in `oversized`. */
  xml: Record<string, string>;
  /** Sizes from the zip central directory, read without inflating anything. */
  sizes: Record<string, { compressed: number; uncompressed: number }>;
  oversized: string[];
}

export const DEFAULT_MAX_ENTRY_BYTES = 64 * 1024 * 1024;
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/** Entry sizes from the end-of-central-directory record and central directory; {} when they cannot be found. */
export function zipEntrySizes(bytes: Uint8Array): OoxmlParts['sizes'] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sizes: OoxmlParts['sizes'] = {};
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 0xffff); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return sizes;
  let at = view.getUint32(eocd + 16, true);
  for (let n = view.getUint16(eocd + 10, true); n > 0; n--) {
    if (at + 46 > bytes.length || view.getUint32(at, true) !== 0x02014b50) break;
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    const name = new TextDecoder().decode(bytes.subarray(at + 46, at + 46 + nameLen));
    sizes[name] = {
      compressed: view.getUint32(at + 20, true),
      uncompressed: view.getUint32(at + 24, true),
    };
    at += 46 + nameLen + extraLen + commentLen;
  }
  return sizes;
}

/** Inflate one entry as text, stopping as soon as the output passes `cap` bytes; null when too large or unreadable. */
function readCapped(entry: JSZip.JSZipObject, cap: number): Promise<string | null> {
  return new Promise((resolve) => {
    const chunks: Buffer[] = [];
    let total = 0;
    const stream = entry.nodeStream('nodebuffer');
    stream
      .on('data', (c: Buffer) => {
        total += c.length;
        if (total > cap) {
          stream.pause();
          resolve(null);
        } else chunks.push(c);
      })
      .on('error', () => resolve(null))
      .on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
  });
}

export async function readParts(
  bytes: Uint8Array,
  maxEntryBytes = DEFAULT_MAX_ENTRY_BYTES
): Promise<OoxmlParts> {
  const empty: OoxmlParts = {
    kind: 'unreadable',
    raw: bytes,
    names: [],
    xml: {},
    sizes: {},
    oversized: [],
  };
  if (OLE_MAGIC.every((b, i) => bytes[i] === b)) return { ...empty, kind: 'ole' };
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return empty;
  }
  const sizes = zipEntrySizes(bytes);
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const xml: Record<string, string> = {};
  const oversized: string[] = [];
  for (const n of names) {
    if (!/\.(xml|rels)$/i.test(n)) continue;
    const text =
      (sizes[n]?.uncompressed ?? 0) > maxEntryBytes
        ? null
        : await readCapped(zip.files[n], maxEntryBytes);
    if (text === null) oversized.push(n);
    else xml[n] = text;
  }
  return { kind: 'zip', raw: bytes, names, xml, sizes, oversized };
}

/** Tag-balance check; true when every element opens and closes properly (comments, CDATA, PIs and a DOCTYPE skipped). */
export function isWellFormed(xml: string): boolean {
  const stack: string[] = [];
  const stripped = xml
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '')
    .replace(/<\?[\s\S]*?\?>/g, '')
    .replace(/<!DOCTYPE(?:[^[>]|\[[\s\S]*?\])*>/g, '');
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:[^>"']|"[^"]*"|'[^']*')*?)(\/?)>/g;
  let last = 0;
  for (let m = re.exec(stripped); m; m = re.exec(stripped)) {
    if (stripped.slice(last, m.index).includes('<')) return false;
    last = m.index + m[0].length;
    if (m[4]) continue;
    if (m[1]) {
      if (stack.pop() !== m[2]) return false;
    } else stack.push(m[2]);
  }
  return stack.length === 0 && !stripped.slice(last).includes('<');
}

/** Resolve a relationship target against the part's folder; null when it climbs out of the package root. */
function resolveTarget(base: string, target: string): string | null {
  const segs = target.startsWith('/') ? [] : base.split('/').filter(Boolean);
  for (const seg of target.split('/')) {
    if (seg === '..') {
      if (segs.pop() === undefined) return null;
    } else if (seg !== '.' && seg !== '') segs.push(seg);
  }
  return segs.join('/');
}

function safeDecode(s: string): string {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/**
 * Part names that a `.rels` part's internal relationships point at, resolved against the part it describes.
 * `target` is null for external targets and for relationships without one; a target that escapes the package
 * root resolves to the unusable name `../<target>` so that it counts as a missing part.
 */
export function relationshipTargets(
  relsName: string,
  xml: string
): Array<{ id: string; target: string | null }> {
  const base = relsName.replace(/_rels\/[^/]*\.rels$/, '');
  const out: Array<{ id: string; target: string | null }> = [];
  for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]*)"/.exec(m[0])?.[1] ?? '';
    const target = /\bTarget="([^"]*)"/.exec(m[0])?.[1];
    if (!target || /TargetMode="External"/.test(m[0])) {
      out.push({ id, target: null });
      continue;
    }
    out.push({ id, target: safeDecode(resolveTarget(base, target) ?? `../${target}`) });
  }
  return out;
}
