import { parseXmlTree } from '../common/safe-xml';

// ---------------------------------------------------------------------------
// A small, safe XML tree for the XLSX loss detector. Same parser and approach as docx/xml-tree.ts
// (preserveOrder, no value coercion), turned into a plain element tree with namespaces resolved, so a
// producer that writes `<x:worksheet>` or an x14 extension under any prefix is read like the usual markup.
//
// Safety: the parser and its limits (no entity expansion, DOCTYPE removed, depth cap) are `common/safe-xml.ts`. The
// five predefined entities and numeric references are decoded here, one pass, with no recursion. Parsing
// never throws: damaged XML yields `null`.
// ---------------------------------------------------------------------------

export interface XmlElement {
  /** Local name, no prefix. */
  name: string;
  /** Short code for the namespaces the detector cares about, the raw URI for others, '' when none. */
  ns: string;
  /** Attributes by local name (a prefixed attribute loses its prefix unless an unprefixed one of that name exists). */
  attrs: Record<string, string>;
  /** Concatenated character data directly inside the element. */
  text: string;
  children: XmlElement[];
}

const NAMESPACES: Record<string, string> = {
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main': 'main',
  // network-allow: XML namespace identifier of the ISO strict variant, never requested
  'http://purl.oclc.org/ooxml/spreadsheetml/main': 'main',
  'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main': 'x14',
  'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main': 'x15',
  'http://schemas.microsoft.com/office/excel/2006/main': 'xm',
  'http://schemas.openxmlformats.org/markup-compatibility/2006': 'mc',
  'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing': 'xdr',
  // network-allow: XML namespace identifier of the ISO strict variant, never requested
  'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing': 'xdr',
  'http://schemas.openxmlformats.org/package/2006/relationships': 'rel',
};

const ENTITY = /&(?:#(\d+)|#x([0-9a-fA-F]+)|(lt|gt|amp|quot|apos));/g;
const NAMED: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

/** Decode the predefined and numeric character references; anything else (a DOCTYPE entity) stays as written. */
export function decodeEntities(s: string): string {
  if (!s.includes('&')) return s;
  return s.replace(ENTITY, (whole, dec: string | undefined, hex: string | undefined, named: string | undefined) => {
    if (named) return NAMED[named];
    const code = dec !== undefined ? Number(dec) : parseInt(hex as string, 16);
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : whole;
  });
}

type RawNode = Record<string, unknown> & { ':@'?: Record<string, string> };

function convert(raw: RawNode[], scope: Map<string, string>): XmlElement[] {
  const out: XmlElement[] = [];
  for (const node of raw) {
    const tag = Object.keys(node).find((k) => k !== ':@');
    if (!tag || tag === '#text' || tag.startsWith('?') || tag.startsWith('!')) continue;
    const rawAttrs = node[':@'] ?? {};
    let inner = scope;
    for (const key of Object.keys(rawAttrs)) {
      if (key === '@_xmlns' || key.startsWith('@_xmlns:')) {
        if (inner === scope) inner = new Map(scope);
        inner.set(key === '@_xmlns' ? '' : key.slice(8), rawAttrs[key]);
      }
    }
    const colon = tag.indexOf(':');
    const uri = inner.get(colon < 0 ? '' : tag.slice(0, colon));
    const attrs: Record<string, string> = {};
    for (const key of Object.keys(rawAttrs)) {
      const name = key.slice(2);
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      const local = name.slice(name.indexOf(':') + 1);
      if (local === name || !(local in attrs)) attrs[local] = decodeEntities(rawAttrs[key]);
    }
    const kids = node[tag] as RawNode[];
    let text = '';
    for (const kid of kids) if ('#text' in kid) text += String(kid['#text']);
    out.push({
      name: tag.slice(colon + 1),
      ns: uri === undefined ? '' : (NAMESPACES[uri] ?? uri),
      attrs,
      text: decodeEntities(text),
      children: convert(kids, inner),
    });
  }
  return out;
}

/** Parse a part into its root element, or `null` when it is empty, not XML or too deeply nested. Never throws. */
export function parseXml(xml: string): XmlElement | null {
  const tree = parseXmlTree(xml);
  if (tree === null) return null;
  return convert(tree as RawNode[], new Map())[0] ?? null;
}

/** Main-namespace markup, tolerating a part that forgot to declare any namespace. */
export const isMain = (el: XmlElement): boolean => el.ns === 'main' || el.ns === '';

/** Direct children with a local name (any namespace). */
export function kids(el: XmlElement | undefined, name: string): XmlElement[] {
  return el ? el.children.filter((c) => c.name === name) : [];
}

export function kid(el: XmlElement | undefined, name: string): XmlElement | undefined {
  return el?.children.find((c) => c.name === name);
}

/**
 * Pre-order walk of the descendants (the element itself excluded). `mc:Fallback` branches are skipped: they repeat
 * what the matching `mc:Choice` holds, and counting both would double every object.
 */
export function walk(el: XmlElement | undefined, visit: (e: XmlElement) => void): void {
  if (!el) return;
  const stack: XmlElement[] = [...el.children].reverse();
  while (stack.length > 0) {
    const e = stack.pop() as XmlElement;
    if (e.name === 'Fallback' && e.ns === 'mc') continue;
    visit(e);
    for (let i = e.children.length - 1; i >= 0; i--) stack.push(e.children[i]);
  }
}

/** Every descendant with a local name (and optionally a namespace code), in document order. */
export function descendants(el: XmlElement | undefined, name: string, ns?: string): XmlElement[] {
  const found: XmlElement[] = [];
  walk(el, (e) => {
    if (e.name === name && (ns === undefined || e.ns === ns)) found.push(e);
  });
  return found;
}

/** An OOXML boolean attribute: `1` or `true`. */
export const isTrue = (v: string | undefined): boolean => v === '1' || v === 'true';

/** An OOXML boolean attribute that defaults to true: anything but `0` or `false`. */
export const isOn = (v: string | undefined): boolean => v !== '0' && v !== 'false';
