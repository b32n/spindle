import { decodeEntities } from '../xml';
import { XML_LIMITS } from '../../common/safe-xml';

// ---------------------------------------------------------------------------
// A forward-only XML event scanner for the one part that can be large, a worksheet's `sheetData`. It reads a string and calls
// back for each start tag, end tag and run of character data, never building a tree: what a caller holds is what it keeps from
// the events. The text itself is held whole (the package hands a part over as one string, so memory is at least the part's size);
// the scanner reads it once, front to back, and could be fed chunks of the same shape by a worker later (S20).
//
// - Names are local (the prefix is dropped), so `<x:c>` reads like `<c>`; attributes likewise (`r:id` is `id`; an unprefixed
//   attribute of the same name wins). `xmlns` attributes are not reported.
// - Attributes come as a null-prototype object: the names are the file's text, and `__proto__` or `constructor` must be an
//   attribute like any other and never reach `Object.prototype`.
// - Safety: entities are never expanded or fetched (the five predefined ones and numeric references are decoded in one pass by
//   `decodeEntities`, anything else stays as written), a DOCTYPE is skipped, nesting is capped (`XML_LIMITS.maxDepth`).
// - Never throws. Damaged markup (a stray or mismatched end tag, an unclosed element at the end, a tag cut short) stops the scan
//   and `scanXml` answers `false`; what was reported before that point stands.
// ---------------------------------------------------------------------------

export type SaxAttrs = Readonly<Record<string, string>>;

export interface SaxHandler {
  /** A start tag (`selfClosing` for `<a/>`, which is followed by its own `close`). `start`/`end` are the tag's offsets in the text. */
  open(name: string, attrs: SaxAttrs, selfClosing: boolean, start: number, end: number): void;
  /** An end tag, or the implied one of a self-closing element. `end` is the offset just after it. */
  close(name: string, end: number): void;
  /** Character data (entities decoded, CDATA included), possibly in several pieces per element. */
  text(text: string): void;
  /** Asked before each markup item: `true` ends the scan there (a caller that has all it will take). */
  halted?(): boolean;
}

const ATTR = /\s*([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/y;
const NAME = /[^\s/>]+/y;

/** The local part of a possibly prefixed name. */
const local = (name: string): string => name.slice(name.indexOf(':') + 1);

/** Scan `xml`, reporting to `handler`. Returns `true` when the whole text was well-formed (balanced, within the depth cap), or the handler ended the scan before the damage, if any, was reached. */
export function scanXml(xml: string, handler: SaxHandler, maxDepth: number = XML_LIMITS.maxDepth): boolean {
  const stack: string[] = [];
  const n = xml.length;
  let i = 0;
  while (i < n) {
    if (handler.halted?.()) return true;
    const lt = xml.indexOf('<', i);
    if (lt < 0) {
      // trailing text outside any element is whitespace in a sound part; elements still open are damage
      return stack.length === 0;
    }
    if (lt > i && stack.length > 0) handler.text(decodeEntities(xml.slice(i, lt)));
    i = lt + 1;
    const c = xml.charCodeAt(i);
    if (c === 63 /* ? */) {
      const end = xml.indexOf('?>', i);
      if (end < 0) return false;
      i = end + 2;
    } else if (c === 33 /* ! */) {
      if (xml.startsWith('!--', i)) {
        const end = xml.indexOf('-->', i + 3);
        if (end < 0) return false;
        i = end + 3;
      } else if (xml.startsWith('![CDATA[', i)) {
        const end = xml.indexOf(']]>', i + 8);
        if (end < 0) return false;
        if (stack.length > 0) handler.text(xml.slice(i + 8, end));
        i = end + 3;
      } else {
        // DOCTYPE: skipped with its internal subset (where entities would be declared); nothing in it is ever used
        const open = xml.indexOf('[', i);
        const close = xml.indexOf('>', i);
        if (close < 0) return false;
        if (open >= 0 && open < close) {
          const subset = xml.indexOf(']', open);
          const end = subset < 0 ? -1 : xml.indexOf('>', subset);
          if (end < 0) return false;
          i = end + 1;
        } else i = close + 1;
      }
    } else if (c === 47 /* / */) {
      const end = xml.indexOf('>', i);
      if (end < 0) return false;
      const name = local(xml.slice(i + 1, end).trim());
      if (stack.pop() !== name) return false;
      handler.close(name, end + 1);
      i = end + 1;
    } else {
      NAME.lastIndex = i;
      const m = NAME.exec(xml);
      if (m === null) return false;
      const name = local(m[0]);
      i += m[0].length;
      const attrs: Record<string, string> = Object.create(null);
      for (;;) {
        ATTR.lastIndex = i;
        const a = ATTR.exec(xml);
        if (a === null) break;
        i += a[0].length;
        const key = a[1];
        if (key === 'xmlns' || key.startsWith('xmlns:')) continue;
        const short = local(key);
        // an unprefixed attribute beats a prefixed one of the same local name
        if (short === key || !(short in attrs)) attrs[short] = decodeEntities(a[2] ?? a[3]);
      }
      while (i < n && /\s/.test(xml[i])) i++;
      const selfClosing = xml.charCodeAt(i) === 47;
      if (selfClosing) i++;
      if (xml.charCodeAt(i) !== 62 /* > */) return false;
      const start = lt;
      i++;
      handler.open(name, attrs, selfClosing, start, i);
      if (selfClosing) handler.close(name, i);
      else {
        stack.push(name);
        if (stack.length > maxDepth) return false;
      }
    }
  }
  return stack.length === 0;
}
