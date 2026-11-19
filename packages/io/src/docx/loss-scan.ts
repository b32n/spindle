import type { ImportLossItem } from '@b32nio/spindle-shared';
import { drawingKind } from './drawing-kind';
import { attrsOf, childrenOf, find, findAll, parseXml, tagOf, walk, type XmlNode } from './xml-tree';

/**
 * XML-based detectors for constructs the importer does not bring across. They parse the part and walk the
 * element tree, so they count real elements (not look-alike text) and know their ancestors, which a regex
 * cannot. Each returns zero or one loss item; the caller appends them to the report.
 */

export interface DocxScanInput {
  /** word/document.xml */
  document: string;
  /** word/styles.xml ('' when absent) */
  styles: string;
  /** word/numbering.xml ('' when absent) */
  numbering: string;
  /** word/settings.xml ('' when absent) */
  settings: string;
  /** word/header*.xml and word/footer*.xml */
  headerFooters: string[];
  /** Number of embedded font files under word/fonts/ */
  fontFiles: number;
}

type Item = Omit<ImportLossItem, 'where'> & { where?: string[] };

const MAX_WHERE = 10;

const isOff = (v: string | undefined) => v === '0' || v === 'false' || v === 'off';
const isOn = (node: XmlNode) => !isOff(attrsOf(node)['@_w:val']);
const attr = (node: XmlNode, name: string) => attrsOf(node)[`@_${name}`];

/** "PAGEREF ×111" entries, largest first; the tail is folded into one entry. */
function breakdown(counts: Map<string, number>, max = MAX_WHERE): string[] {
  const entries = [...counts.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const shown = entries.slice(0, max).map(([k, n]) => `${k} ×${n}`);
  const rest = entries.slice(max).reduce((sum, [, n]) => sum + n, 0);
  if (rest > 0) shown.push(`other ×${rest}`);
  return shown;
}

function bump(map: Map<string, number>, key: string, by = 1) {
  map.set(key, (map.get(key) ?? 0) + by);
}

function item(id: string, label: string, count: number, effect: string, where?: Map<string, number>, maxWhere = MAX_WHERE): Item[] {
  if (count <= 0) return [];
  return [{ id, label, count, effect, ...(where && where.size > 0 && { where: breakdown(where, maxWhere) }) }];
}

const inRevisionProps = (path: readonly string[]) => path.some((t) => t === 'w:pPrChange' || t === 'w:rPrChange' || t === 'w:sectPrChange' || t === 'w:tblPrChange' || t === 'w:tcPrChange' || t === 'w:trPrChange');

// ---------------------------------------------------------------------------
// Styles and numbering lookups
// ---------------------------------------------------------------------------

interface NumberingLookup {
  /** numId -> abstractNumId */
  abstractOf: Map<string, string>;
  /** `${abstractId}:${ilvl}` -> { fmt, start } */
  levels: Map<string, { fmt?: string; start?: string }>;
  /** `${numId}:${ilvl}` -> the override's own start value / level definition */
  overrides: Map<string, { fmt?: string; start?: string }>;
}

function numberingLookup(xml: string): NumberingLookup {
  const out: NumberingLookup = { abstractOf: new Map(), levels: new Map(), overrides: new Map() };
  const numberingEl = xml ? find(parseXml(xml), 'w:numbering') : undefined;
  if (!numberingEl) return out;
  for (const abs of findAll(childrenOf(numberingEl), 'w:abstractNum')) {
    const absId = attr(abs, 'w:abstractNumId');
    if (!absId) continue;
    for (const lvl of findAll(childrenOf(abs), 'w:lvl')) {
      const kids = childrenOf(lvl);
      const fmtEl = find(kids, 'w:numFmt');
      const startEl = find(kids, 'w:start');
      out.levels.set(`${absId}:${attr(lvl, 'w:ilvl') ?? '0'}`, {
        fmt: fmtEl ? attr(fmtEl, 'w:val') : undefined,
        start: startEl ? attr(startEl, 'w:val') : undefined,
      });
    }
  }
  for (const num of findAll(childrenOf(numberingEl), 'w:num')) {
    const numId = attr(num, 'w:numId');
    const absEl = find(childrenOf(num), 'w:abstractNumId');
    const absId = absEl ? attr(absEl, 'w:val') : undefined;
    if (!numId) continue;
    if (absId) out.abstractOf.set(numId, absId);
    for (const ov of findAll(childrenOf(num), 'w:lvlOverride')) {
      const kids = childrenOf(ov);
      const startEl = find(kids, 'w:startOverride');
      const lvl = find(kids, 'w:lvl');
      const lvlKids = lvl ? childrenOf(lvl) : [];
      const fmtEl = find(lvlKids, 'w:numFmt');
      const lvlStart = find(lvlKids, 'w:start');
      const start = startEl ? attr(startEl, 'w:val') : lvlStart ? attr(lvlStart, 'w:val') : undefined;
      const fmt = fmtEl ? attr(fmtEl, 'w:val') : undefined;
      if (start !== undefined || fmt !== undefined) out.overrides.set(`${numId}:${attr(ov, 'w:ilvl') ?? '0'}`, { fmt, start });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Detectors
// ---------------------------------------------------------------------------

/** Sections beyond the first: kept in the model and exported, but the editor shows and edits only the first. */
function extraSections(root: XmlNode[]): Item[] {
  let sectPrs = 0;
  // Only body-level and paragraph-level section properties end a section (as in the reader).
  walk(root, (_n, tag, path) => {
    if (tag !== 'w:sectPr') return;
    const p = path.join('/');
    if (p === 'w:document/w:body' || p === 'w:document/w:body/w:p/w:pPr') sectPrs++;
  });
  return item(
    'docx.sections',
    'extra section(s)',
    sectPrs - 1,
    'kept in the document and exported, but the editor shows and edits only the first section'
  );
}

const cellHasContent = (p: XmlNode) => {
  let has = false;
  walk([p], (n, tag) => {
    if (tag === 'w:drawing' || tag === 'w:pict') has = true;
    if (tag === 'w:t' && childrenOf(n).some((c) => String((c as Record<string, unknown>)['#text'] ?? '') !== '')) has = true;
  });
  return has;
};

/** Visits each table's own properties and rows, without descending into nested tables (those are visited as tables themselves). */
function ownTableNodes(tbl: XmlNode, visit: (node: XmlNode, tag: string, path: readonly string[]) => void) {
  walk(childrenOf(tbl), (n, tag, path) => visit(n, tag, path));
}

function tableDetectors(root: XmlNode[]): Item[] {
  const tables: { node: XmlNode; nested: boolean }[] = [];
  walk(root, (node, tag, path) => {
    if (tag === 'w:tbl') tables.push({ node, nested: path.includes('w:tc') });
  });

  const formatting = new Map<string, number>();
  let formattedTables = 0;
  let nestedTables = 0;
  let multiParagraphCells = 0;

  for (const { node: tbl, nested } of tables) {
    if (nested) nestedTables++;
    const kinds = new Set<string>();
    // Anything below a nested w:tbl belongs to that table, not this one.
    ownTableNodes(tbl, (n, tag, path) => {
      if (path.includes('w:tbl')) return;
      const parent = path[path.length - 1];
      if (tag === 'w:tblBorders' || tag === 'w:tcBorders') kinds.add('borders');
      else if (tag === 'w:shd' && (parent === 'w:tblPr' || parent === 'w:tcPr')) {
        if (attr(n, 'w:fill') && attr(n, 'w:fill') !== 'auto') kinds.add('shading');
      } else if ((tag === 'w:tcW' || tag === 'w:tblW') && attr(n, 'w:type') !== 'auto' && Number(attr(n, 'w:w') ?? 0) > 0) kinds.add('widths');
      else if (tag === 'w:tblHeader' && isOn(n)) kinds.add('repeating header rows');
      else if (tag === 'w:vAlign' && parent === 'w:tcPr' && attr(n, 'w:val') !== 'top') kinds.add('vertical alignment');
      else if (tag === 'w:tblStyle' && attr(n, 'w:val') !== 'TableNormal') kinds.add('table style');
      else if (tag === 'w:tc') {
        const paragraphs = findAll(childrenOf(n), 'w:p').filter(cellHasContent);
        if (paragraphs.length > 1) multiParagraphCells++;
      }
    });
    if (kinds.size > 0) {
      formattedTables++;
      for (const k of kinds) bump(formatting, k);
    }
  }

  return [
    ...item('docx.tableFormatting', 'table(s) with borders, shading, widths, header rows or styles', formattedTables, 'the formatting is dropped', formatting),
    ...item(
      'docx.tableStructure',
      'nested table(s) / multi-paragraph cell(s)',
      nestedTables + multiParagraphCells,
      'nested tables are dropped and only the first paragraph of a cell is kept',
      new Map([...(nestedTables ? [['nested tables', nestedTables] as const] : []), ...(multiParagraphCells ? [['multi-paragraph cells', multiParagraphCells] as const] : [])])
    ),
  ];
}

function listFormatting(root: XmlNode[], numbering: NumberingLookup): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  walk(root, (node, tag, path) => {
    if (tag !== 'w:p' || inRevisionProps(path)) return;
    const pPr = find(childrenOf(node), 'w:pPr');
    const numPr = pPr ? find(childrenOf(pPr), 'w:numPr') : undefined;
    if (!numPr) return;
    const numIdEl = find(childrenOf(numPr), 'w:numId');
    const ilvlEl = find(childrenOf(numPr), 'w:ilvl');
    const numId = numIdEl ? attr(numIdEl, 'w:val') : undefined;
    if (!numId || numId === '0') return;
    const ilvl = ilvlEl ? (attr(ilvlEl, 'w:val') ?? '0') : '0';
    const abs = numbering.abstractOf.get(numId);
    const base = abs ? numbering.levels.get(`${abs}:${ilvl}`) : undefined;
    const override = numbering.overrides.get(`${numId}:${ilvl}`);
    const fmt = override?.fmt ?? base?.fmt;
    if (!fmt || fmt === 'decimal' || fmt === 'bullet') return;
    total++;
    bump(kinds, fmt);
  });
  // The first number of a list and restarts are imported; only the number format is not.
  return item(
    'docx.listFormatting',
    'list item(s) with another number format',
    total,
    'imported as plain bullet or decimal-numbered items; the number format is not kept',
    kinds
  );
}

function bookmarks(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  walk(root, (node, tag) => {
    if (tag === 'w:bookmarkStart' && !(attr(node, 'w:name') ?? '').startsWith('_GoBack')) {
      total++;
      bump(kinds, 'bookmarks');
    } else if (tag === 'w:hyperlink' && attr(node, 'w:anchor')) {
      total++;
      bump(kinds, 'internal links');
    }
  });
  return item('docx.bookmarks', 'bookmark(s) and internal link(s)', total, 'bookmarks are dropped and internal links become plain text', kinds);
}

function sectionDetectors(root: XmlNode[], evenAndOdd: boolean): Item[] {
  const variantKinds = new Map<string, number>();
  const propKinds = new Map<string, number>();
  let variantSections = 0;
  let propSections = 0;
  walk(root, (sectPr, tag, path) => {
    if (tag !== 'w:sectPr' || inRevisionProps(path)) return;
    const kids = childrenOf(sectPr);
    const reasons: string[] = [];
    const titlePg = find(kids, 'w:titlePg');
    if (titlePg && isOn(titlePg)) reasons.push('first-page header/footer (titlePg)');
    const refs = kids.filter((k) => tagOf(k) === 'w:headerReference' || tagOf(k) === 'w:footerReference');
    for (const ref of refs) {
      const type = attr(ref, 'w:type');
      if (type === 'first' && !titlePg) reasons.push('first-page reference');
      if (type === 'even') reasons.push('even-page reference');
    }
    if (evenAndOdd && refs.length > 0) reasons.push('even/odd headers setting');
    if (reasons.length > 0) {
      variantSections++;
      for (const r of new Set(reasons)) bump(variantKinds, r);
    }

    const props: string[] = [];
    const pgNum = find(kids, 'w:pgNumType');
    if (pgNum && (attr(pgNum, 'w:start') || attr(pgNum, 'w:fmt') || attr(pgNum, 'w:chapStyle'))) props.push('page numbering');
    const vAlign = find(kids, 'w:vAlign');
    if (vAlign && attr(vAlign, 'w:val') && attr(vAlign, 'w:val') !== 'top') props.push('vertical alignment');
    if (props.length > 0) {
      propSections++;
      for (const p of props) bump(propKinds, p);
    }
  });
  return [
    ...item(
      'docx.headerFooterVariants',
      'section(s) with first-page or even/odd headers and footers',
      variantSections,
      'only the first header/footer reference of each section is used, as its default header/footer',
      variantKinds
    ),
    ...item(
      'docx.sectionProperties',
      'section(s) with page-number format/restart or vertical alignment',
      propSections,
      'not imported',
      propKinds
    ),
  ];
}

/** Field instructions in document order: complex fields (begin..end, nested ones included) and fldSimple. */
export function fieldInstructions(root: XmlNode[]): string[] {
  const out: string[] = [];
  const stack: { instr: string; separated: boolean }[] = [];
  walk(root, (node, tag) => {
    if (tag === 'w:fldChar') {
      const type = attr(node, 'w:fldCharType');
      if (type === 'begin') stack.push({ instr: '', separated: false });
      else if (type === 'separate') {
        if (stack.length) stack[stack.length - 1].separated = true;
      } else if (type === 'end') {
        const f = stack.pop();
        if (f) out.push(f.instr);
      }
    } else if (tag === 'w:instrText') {
      const text = childrenOf(node)
        .map((c) => String((c as Record<string, unknown>)['#text'] ?? ''))
        .join('');
      const f = stack[stack.length - 1];
      if (!f) out.push(text); // instruction outside any begin..end: damaged file, still count it
      else if (!f.separated) f.instr += text;
    } else if (tag === 'w:fldSimple') {
      out.push(attr(node, 'w:instr') ?? '');
    }
  });
  // A begin without an end (truncated or hand-edited file) is still a field.
  for (const f of stack.reverse()) out.push(f.instr);
  return out;
}

/**
 * Fields whose displayed value is not kept as a field: a table of contents or index keeps its text but not the field
 * (it is static), and legacy form fields lose their value. Every other field keeps its instruction and cached result.
 * A complex field is counted once even when its instruction is split over several instrText runs.
 */
const STATIC_FIELD_KINDS = new Set(['TOC', 'TOA', 'INDEX']);
const LEGACY_FORM_KINDS = new Set(['FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN']);

function fields(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  for (const raw of fieldInstructions(root)) {
    const kind = raw.trim().toUpperCase().split(/[\s\\"]/)[0];
    if (!STATIC_FIELD_KINDS.has(kind) && !LEGACY_FORM_KINDS.has(kind)) continue;
    total++;
    bump(kinds, kind);
  }
  return item(
    'docx.fields',
    'field(s) (table of contents, index, legacy form fields)',
    total,
    'tables of contents and indexes are kept as static text, legacy form fields lose their value',
    kinds,
    20
  );
}

/** `full` adds the formatting/structure detectors and the per-kind field breakdown; otherwise only the fields item (counts only). */
// ---------------------------------------------------------------------------
// Round 2: long-tail constructs and drawing counts
// ---------------------------------------------------------------------------

const isSectPrChild = (path: readonly string[]) => path[path.length - 1] === 'w:sectPr' && !inRevisionProps(path);

function trackedMoves(root: XmlNode[]): Item[] {
  let total = 0;
  walk(root, (_n, tag) => {
    if (tag === 'w:moveFrom' || tag === 'w:moveTo') total++;
  });
  return item('docx.trackedMoves', 'tracked move(s)', total, 'their text is not imported');
}

function formattingRevisions(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  walk(root, (_n, tag) => {
    if (tag === 'w:rPrChange' || tag === 'w:pPrChange' || tag === 'w:tblPrChange' || tag === 'w:sectPrChange') {
      total++;
      bump(kinds, tag.slice(2));
    }
  });
  return item('docx.formattingRevisions', 'tracked formatting change(s)', total, 'the change history is dropped; the current formatting is imported', kinds);
}

function sectionLevel(root: XmlNode[]): Item[] {
  let borders = 0;
  let lineNumbers = 0;
  walk(root, (_n, tag, path) => {
    if (!isSectPrChild(path)) return;
    if (tag === 'w:pgBorders') borders++;
    else if (tag === 'w:lnNumType') lineNumbers++;
  });
  return [
    ...item('docx.pageBorders', 'section(s) with page borders', borders, 'not imported'),
    ...item('docx.lineNumbers', 'section(s) with line numbering', lineNumbers, 'not imported'),
  ];
}

function dropCaps(root: XmlNode[]): Item[] {
  let total = 0;
  walk(root, (node, tag, path) => {
    if (tag === 'w:framePr' && path[path.length - 1] === 'w:pPr' && !inRevisionProps(path)) {
      const v = attr(node, 'w:dropCap');
      if (v === 'drop' || v === 'margin') total++;
    }
  });
  return item('docx.dropCaps', 'drop cap(s)', total, 'shown as a normal paragraph');
}

function tabStops(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  walk(root, (node, tag, path) => {
    if (tag !== 'w:tab' || path[path.length - 1] !== 'w:tabs' || inRevisionProps(path)) return;
    if (attr(node, 'w:val') === 'clear') return;
    total++;
    const leader = attr(node, 'w:leader');
    bump(kinds, leader && leader !== 'none' ? `${leader} leader` : 'no leader');
  });
  return item('docx.tabStops', 'custom tab stop(s)', total, 'not imported (tabs use default stops)', kinds);
}

function rtl(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  walk(root, (node, tag, path) => {
    if (inRevisionProps(path) || !isOn(node)) return;
    const parent = path[path.length - 1];
    let k: string | undefined;
    if (tag === 'w:bidi' && parent === 'w:pPr') k = 'paragraph direction';
    else if (tag === 'w:rtl' && parent === 'w:rPr') k = 'run direction';
    else if (tag === 'w:bidiVisual' && parent === 'w:tblPr') k = 'table direction';
    if (!k) return;
    total++;
    bump(kinds, k);
  });
  return item('docx.rtl', 'right-to-left setting(s)', total, 'the direction is not kept', kinds);
}

function countElements(root: XmlNode[], tag: string, id: string, label: string, effect: string): Item[] {
  let total = 0;
  walk(root, (_n, t) => {
    if (t === tag) total++;
  });
  return item(id, label, total, effect);
}

function legacyForms(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  for (const raw of fieldInstructions(root)) {
    const kind = raw.trim().toUpperCase().split(/[\s\\]/)[0];
    if (kind !== 'FORMTEXT' && kind !== 'FORMCHECKBOX' && kind !== 'FORMDROPDOWN') continue;
    total++;
    bump(kinds, kind);
  }
  return item('docx.legacyForms', 'legacy form field(s)', total, 'their displayed value is dropped (also counted under docx.fields)', kinds);
}

function watermarks(headerFooters: string[]): Item[] {
  let total = 0;
  for (const xml of headerFooters) {
    walk(parseXml(xml), (node, tag) => {
      if (tag === 'v:shape' && /watermark/i.test(attr(node, 'id') ?? '')) total++;
    });
  }
  return item('docx.watermarks', 'watermark(s)', total, 'not imported');
}

function hyphenation(settings: string): Item[] {
  const el = settings ? find(childrenOf(find(parseXml(settings), 'w:settings') ?? {}), 'w:autoHyphenation') : undefined;
  return item('docx.hyphenation', 'automatic hyphenation setting', el && isOn(el) ? 1 : 0, 'not applied (lines are broken without hyphenation)');
}

/**
 * Drawings counted as objects, not as parts: a chart or SmartArt is one object however many XML parts back it,
 * and only anchored pictures are floating images (anchored shapes, text boxes and charts are other things).
 */
function drawingCounts(root: XmlNode[]): { charts: Item[]; floating: Item[] } {
  let charts = 0;
  let floating = 0;
  walk(root, (drawing, tag) => {
    if (tag !== 'w:drawing') return;
    const kind = drawingKind(drawing);
    if (kind === 'chart' || kind === 'diagram') charts++;
    else if (kind === 'picture' && childrenOf(drawing).some((c) => tagOf(c) === 'wp:anchor')) floating++;
  });
  return {
    charts: item('docx.charts', 'chart(s)/diagram(s)', charts, 'not imported'),
    floating: item('docx.floatingImages', 'floating image(s)', floating, 'not imported (only inline images are)'),
  };
}

export function scanDocx(input: DocxScanInput, full: boolean): Item[] {
  let root: XmlNode[];
  try {
    root = parseXml(input.document);
  } catch {
    return [];
  }
  const evenAndOdd = (() => {
    try {
      const settings = input.settings ? find(parseXml(input.settings), 'w:settings') : undefined;
      const el = settings ? find(childrenOf(settings), 'w:evenAndOddHeaders') : undefined;
      return !!el && isOn(el);
    } catch {
      return false;
    }
  })();
  const safe = (fn: () => Item[]): Item[] => {
    try {
      return fn();
    } catch {
      return [];
    }
  };
  const drawings = safe(() => {
    const { charts, floating } = drawingCounts(root);
    return [...floating, ...charts];
  });
  if (!full) return [...drawings, ...safe(() => fields(root)).map(({ where: _where, ...rest }) => rest)];
  return [
    ...drawings,
    ...safe(() => extraSections(root)),
    ...safe(() => tableDetectors(root)),
    ...safe(() => listFormatting(root, numberingLookup(input.numbering))),
    ...safe(() => bookmarks(root)),
    ...safe(() => sectionDetectors(root, evenAndOdd)),
    ...safe(() => fields(root)),
    ...safe(() => trackedMoves(root)),
    ...safe(() => formattingRevisions(root)),
    ...safe(() => dropCaps(root)),
    ...safe(() => watermarks(input.headerFooters)),
    ...safe(() => sectionLevel(root)),
    ...safe(() => hyphenation(input.settings)),
    ...safe(() => tabStops(root)),
    ...safe(() => rtl(root)),
    ...safe(() => countElements(root, 'w:ruby', 'docx.ruby', 'ruby annotation(s)', 'the base text is kept; the reading marks are not imported')),
    ...item('docx.embeddedFonts', 'embedded font(s)', input.fontFiles, 'not imported (text uses installed fonts)'),
    ...safe(() => countElements(root, 'w:altChunk', 'docx.altChunk', 'embedded alternative-format chunk(s)', 'not imported')),
    ...safe(() => legacyForms(root)),
  ];
}
