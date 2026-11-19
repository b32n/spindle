import type { DocParagraphStyle, DocRunBase } from '@b32nio/spindle-shared/export/docx';

import { attrsOf, childrenOf, find, findAll, parseXml, type XmlNode } from './xml-tree';

// ---------------------------------------------------------------------------
// Styles (word/styles.xml) -> effective formatting.
//
// A run's formatting is the result of layers, lowest first: docDefaults, the paragraph style chain (basedOn),
// the character style chain, and the run's direct properties. A paragraph's: docDefaults, its style chain,
// its direct properties. Properties are kept as partial records (a property that is absent means "inherit"),
// merged layer by layer, and only the final effective values leave this module as `DocRunBase` /
// `DocParagraphStyle`, so the editor model (which has no named styles yet) sees what Word would render.
//
// Toggle properties (bold, italic, small caps, strike): inside one basedOn chain a derived style's value
// replaces its base's; between the paragraph style chain and the character style chain the "on" states are
// XOR-ed (a character style that turns bold on inside a bold paragraph style makes the text not bold);
// direct formatting sets the value outright. A character style linked to the paragraph's own style
// (the "Heading 1 Char" companion of "Heading 1") does not toggle.
// ---------------------------------------------------------------------------

const HIGHLIGHT_TO_HEX: Record<string, string> = {
  black: '#000000', blue: '#0000ff', cyan: '#00ffff', darkBlue: '#00008b', darkCyan: '#008b8b',
  darkGray: '#a9a9a9', darkGreen: '#006400', darkMagenta: '#8b008b', darkRed: '#8b0000',
  darkYellow: '#808000', green: '#00ff00', lightGray: '#d3d3d3', magenta: '#ff00ff',
  red: '#ff0000', white: '#ffffff', yellow: '#ffff00',
};

/** Run properties as a style or a run states them. `null` is an explicit "automatic"/"none" that overrides a lower layer. */
export interface RunPartial {
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  smallCaps?: boolean;
  underline?: boolean;
  vertAlign?: 'superscript' | 'subscript' | 'baseline';
  color?: string | null;
  fontSize?: number;
  fontFamily?: string;
  backgroundColor?: string | null;
}

export interface ParagraphPartial {
  alignment?: DocParagraphStyle['alignment'];
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  leftIndentPt?: number;
  firstLineIndentPt?: number;
  /** numPr: `numId` 0 switches numbering off. */
  numId?: string;
  ilvl?: number;
  /** 0-8 for outline levels 1-9, 9 for body text. */
  outlineLvl?: number;
}

const TOGGLES = ['bold', 'italic', 'strike', 'smallCaps'] as const;

/** An on/off element: absent -> undefined, `w:val` 0/false/off -> false, anything else true. */
function onOff(node: XmlNode | undefined): boolean | undefined {
  if (!node) return undefined;
  const v = attrsOf(node)['@_w:val'];
  return !(v === '0' || v === 'false' || v === 'off');
}

export function parseRPr(rPr: XmlNode | undefined): RunPartial {
  if (!rPr) return {};
  const props = childrenOf(rPr);
  const out: RunPartial = {};
  const set = <K extends keyof RunPartial>(k: K, v: RunPartial[K] | undefined) => {
    if (v !== undefined) out[k] = v;
  };
  set('bold', onOff(find(props, 'w:b')));
  set('italic', onOff(find(props, 'w:i')));
  set('strike', onOff(find(props, 'w:strike')));
  set('smallCaps', onOff(find(props, 'w:smallCaps')));
  const uEl = find(props, 'w:u');
  if (uEl) set('underline', attrsOf(uEl)['@_w:val'] !== 'none');
  const vert = find(props, 'w:vertAlign');
  const vertVal = vert ? attrsOf(vert)['@_w:val'] : undefined;
  if (vertVal === 'superscript' || vertVal === 'subscript' || vertVal === 'baseline') out.vertAlign = vertVal;
  const colorEl = find(props, 'w:color');
  if (colorEl) {
    const val = attrsOf(colorEl)['@_w:val'];
    set('color', val && val !== 'auto' ? `#${val.toLowerCase()}` : null);
  }
  const szEl = find(props, 'w:sz');
  if (szEl) {
    const half = Number(attrsOf(szEl)['@_w:val']);
    if (half > 0) out.fontSize = half / 2;
  }
  const fontsEl = find(props, 'w:rFonts');
  if (fontsEl) set('fontFamily', attrsOf(fontsEl)['@_w:ascii'] || undefined);
  const shdEl = find(props, 'w:shd');
  const highlightEl = find(props, 'w:highlight');
  if (highlightEl) {
    const v = attrsOf(highlightEl)['@_w:val'];
    set('backgroundColor', v && v !== 'none' ? (HIGHLIGHT_TO_HEX[v] ?? null) : null);
  } else if (shdEl) {
    const fill = attrsOf(shdEl)['@_w:fill'];
    if (fill !== undefined) set('backgroundColor', fill && fill !== 'auto' && fill !== 'none' ? `#${fill.toLowerCase()}` : null);
  }
  return out;
}

const ALIGN_MAP: Record<string, DocParagraphStyle['alignment']> = {
  left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify',
};

export function parsePPr(pPr: XmlNode | undefined): ParagraphPartial {
  if (!pPr) return {};
  const kids = childrenOf(pPr);
  const out: ParagraphPartial = {};
  const jc = find(kids, 'w:jc');
  const alignment = jc ? ALIGN_MAP[attrsOf(jc)['@_w:val']] : undefined;
  if (alignment) out.alignment = alignment;
  const spacing = find(kids, 'w:spacing');
  if (spacing) {
    const a = attrsOf(spacing);
    if (a['@_w:before']) out.spaceBeforePt = Number(a['@_w:before']) / 20;
    if (a['@_w:after']) out.spaceAfterPt = Number(a['@_w:after']) / 20;
  }
  const ind = find(kids, 'w:ind');
  if (ind) {
    const a = attrsOf(ind);
    const left = a['@_w:left'] ?? a['@_w:start'];
    if (left) out.leftIndentPt = Number(left) / 20;
    if (a['@_w:firstLine']) out.firstLineIndentPt = Number(a['@_w:firstLine']) / 20;
  }
  const numPr = find(kids, 'w:numPr');
  if (numPr) {
    const numKids = childrenOf(numPr);
    const numIdEl = find(numKids, 'w:numId');
    const ilvlEl = find(numKids, 'w:ilvl');
    if (numIdEl) out.numId = attrsOf(numIdEl)['@_w:val'];
    if (ilvlEl) out.ilvl = Number(attrsOf(ilvlEl)['@_w:val']);
  }
  const outline = find(kids, 'w:outlineLvl');
  if (outline) {
    const n = Number(attrsOf(outline)['@_w:val']);
    if (Number.isFinite(n)) out.outlineLvl = n;
  }
  return out;
}

/** Later layers win key by key. */
function overlay<T extends object>(base: T, top: T): T {
  const out = { ...base };
  for (const [k, v] of Object.entries(top)) if (v !== undefined) (out as Record<string, unknown>)[k] = v;
  return out;
}

interface StyleDef {
  id: string;
  type: string;
  name: string;
  basedOn?: string;
  link?: string;
  isDefault: boolean;
  run: RunPartial;
  para: ParagraphPartial;
}

const MAX_CHAIN = 50;

export class StyleSheet {
  private readonly styles = new Map<string, StyleDef>();
  private docDefaultRun: RunPartial = {};
  private docDefaultPara: ParagraphPartial = {};
  private defaultParagraphStyle: string | undefined;
  private readonly paragraphCache = new Map<string, { run: RunPartial; para: ParagraphPartial; headingLevel?: number }>();
  private readonly characterCache = new Map<string, RunPartial>();

  constructor(xml?: string) {
    if (!xml) return;
    let root: XmlNode[];
    try {
      root = parseXml(xml);
    } catch {
      return; // a damaged styles part leaves everything unstyled, as before
    }
    const stylesEl = find(root, 'w:styles');
    if (!stylesEl) return;
    const kids = childrenOf(stylesEl);
    const defaults = find(kids, 'w:docDefaults');
    if (defaults) {
      const rDef = find(childrenOf(defaults), 'w:rPrDefault');
      const pDef = find(childrenOf(defaults), 'w:pPrDefault');
      this.docDefaultRun = parseRPr(rDef ? find(childrenOf(rDef), 'w:rPr') : undefined);
      this.docDefaultPara = parsePPr(pDef ? find(childrenOf(pDef), 'w:pPr') : undefined);
    }
    for (const el of findAll(kids, 'w:style')) {
      const a = attrsOf(el);
      const id = a['@_w:styleId'];
      if (!id) continue;
      const sk = childrenOf(el);
      const attr = (tag: string) => {
        const n = find(sk, tag);
        return n ? attrsOf(n)['@_w:val'] : undefined;
      };
      const def: StyleDef = {
        id,
        type: a['@_w:type'] ?? 'paragraph',
        name: attr('w:name') ?? '',
        basedOn: attr('w:basedOn'),
        link: attr('w:link'),
        isDefault: a['@_w:default'] === '1' || a['@_w:default'] === 'true',
        run: parseRPr(find(sk, 'w:rPr')),
        para: parsePPr(find(sk, 'w:pPr')),
      };
      this.styles.set(id, def);
      if (def.type === 'paragraph' && def.isDefault && !this.defaultParagraphStyle) this.defaultParagraphStyle = id;
    }
  }

  /** styleId -> style name (what the importer has always exposed for heading detection). */
  name(styleId: string): string | undefined {
    return this.styles.get(styleId)?.name || undefined;
  }

  /** The style and its ancestors, root first; stops at a missing base or a loop. */
  private chain(styleId: string | undefined): StyleDef[] {
    const out: StyleDef[] = [];
    const seen = new Set<string>();
    let id = styleId;
    while (id && out.length < MAX_CHAIN && !seen.has(id)) {
      const def = this.styles.get(id);
      if (!def) break;
      seen.add(id);
      out.unshift(def);
      id = def.basedOn;
    }
    return out;
  }

  private paragraphLayers(styleId: string | undefined): { run: RunPartial; para: ParagraphPartial; headingLevel?: number } {
    const key = styleId ?? '';
    const cached = this.paragraphCache.get(key);
    if (cached) return cached;
    const effectiveId = styleId && this.styles.has(styleId) ? styleId : this.defaultParagraphStyle;
    const chain = this.chain(effectiveId);
    let run = this.docDefaultRun;
    let para = this.docDefaultPara;
    for (const def of chain) {
      run = overlay(run, def.run);
      para = overlay(para, def.para);
    }
    const result = { run, para, headingLevel: this.headingLevelOf(chain, para) };
    this.paragraphCache.set(key, result);
    return result;
  }

  /** A "Heading N" style, else an outline level in the style chain, else a "Heading N" ancestor. */
  private headingLevelOf(chain: StyleDef[], para: ParagraphPartial): number | undefined {
    const own = chain[chain.length - 1];
    const byName = (def: StyleDef | undefined) => def?.name.match(/^Heading\s*(\d)$/i)?.[1];
    const ownMatch = byName(own);
    if (ownMatch) return Math.min(6, Math.max(1, Number(ownMatch)));
    if (para.outlineLvl !== undefined) return para.outlineLvl <= 8 ? Math.min(6, para.outlineLvl + 1) : undefined;
    for (let i = chain.length - 2; i >= 0; i--) {
      const m = byName(chain[i]);
      if (m) return Math.min(6, Math.max(1, Number(m)));
    }
    return undefined;
  }

  private characterLayer(styleId: string | undefined): RunPartial {
    if (!styleId) return {};
    const cached = this.characterCache.get(styleId);
    if (cached) return cached;
    let run: RunPartial = {};
    for (const def of this.chain(styleId)) run = overlay(run, def.run);
    this.characterCache.set(styleId, run);
    return run;
  }

  /** Effective paragraph properties: docDefaults, the style chain, then the paragraph's direct properties. */
  paragraph(pStyleId: string | undefined, direct: ParagraphPartial): ParagraphPartial & { headingLevel?: number } {
    const layers = this.paragraphLayers(pStyleId);
    const para = overlay(layers.para, direct);
    // The level of a heading is a property of the style unless the paragraph states its own outline level.
    const headingLevel = direct.outlineLvl !== undefined ? (direct.outlineLvl <= 8 ? Math.min(6, direct.outlineLvl + 1) : undefined) : layers.headingLevel;
    return { ...para, headingLevel };
  }

  /** Effective run formatting: layers from docDefaults up to the run's own `w:rPr` (which may name a character style). */
  run(rPr: XmlNode | undefined, pStyleId: string | undefined): DocRunBase {
    const direct = parseRPr(rPr);
    const rStyleEl = rPr ? find(childrenOf(rPr), 'w:rStyle') : undefined;
    const rStyleId = rStyleEl ? attrsOf(rStyleEl)['@_w:val'] : undefined;
    const paragraph = this.paragraphLayers(pStyleId).run;
    const character = this.characterLayer(rStyleId);

    const linked = !!rStyleId && !!pStyleId && (this.styles.get(rStyleId)?.link === pStyleId || this.styles.get(pStyleId)?.link === rStyleId);
    const merged: RunPartial = overlay(overlay(paragraph, character), direct);
    for (const t of TOGGLES) {
      if (direct[t] !== undefined) continue;
      const p = paragraph[t] ?? false;
      const c = linked ? false : (character[t] ?? false);
      merged[t] = p !== c;
    }
    return toRunBase(merged);
  }
}

function toRunBase(p: RunPartial): DocRunBase {
  return {
    ...(p.bold && { bold: true }),
    ...(p.italic && { italic: true }),
    ...(p.strike && { strikethrough: true }),
    ...(p.underline && { underline: true }),
    ...(p.smallCaps && { smallCaps: true }),
    ...(p.vertAlign === 'superscript' && { superscript: true }),
    ...(p.vertAlign === 'subscript' && { subscript: true }),
    ...(p.color && { color: p.color }),
    ...(p.fontSize && { fontSize: p.fontSize }),
    ...(p.fontFamily && { fontFamily: p.fontFamily }),
    ...(p.backgroundColor && { backgroundColor: p.backgroundColor }),
  };
}
