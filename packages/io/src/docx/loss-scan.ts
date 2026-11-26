import type { ImportLossItem } from '@b32nio/spindle-shared';
import { unsupportedControls } from './content-controls';
import { findTocRanges, tocLists } from './toc-field';
import { drawingKind } from './drawing-kind';
import { attrsOf, childrenOf, find, findAll, findDeep, parseXml, tagOf, walk, type XmlNode } from './xml-tree';

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

/** Section properties still not kept: chapter numbers in page numbers (page number start and format, vertical alignment, headers and footers of every variant are imported). */
function sectionDetectors(root: XmlNode[]): Item[] {
  let sections = 0;
  walk(root, (sectPr, tag, path) => {
    if (tag !== 'w:sectPr' || inRevisionProps(path)) return;
    const pgNum = find(childrenOf(sectPr), 'w:pgNumType');
    if (pgNum && attr(pgNum, 'w:chapStyle')) sections++;
  });
  return item(
    'docx.sectionProperties',
    'section(s) with chapter numbers in the page numbers',
    sections,
    'the chapter number is not part of the page number',
    new Map(sections ? [['chapter numbering', sections] as const] : [])
  );
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
 * Fields whose displayed value is not kept as a field: an index keeps its text but not the field (it is static), as does a
 * table of contents that is not read as a block (see `findTocRanges`). Legacy form fields keep their text too and are reported as `docx.legacyForms`. Every other field keeps its instruction and cached result.
 * A complex field is counted once even when its instruction is split over several instrText runs.
 */
const STATIC_FIELD_KINDS = new Set(['TOC', 'TOA', 'INDEX']);

function fields(root: XmlNode[]): Item[] {
  const kinds = new Map<string, number>();
  let total = 0;
  let tocs = 0;
  for (const raw of fieldInstructions(root)) {
    const kind = raw.trim().toUpperCase().split(/[\s\\"]/)[0];
    if (kind === 'TOC') tocs++;
    if (!STATIC_FIELD_KINDS.has(kind)) continue;
    total++;
    bump(kinds, kind);
  }
  // A table of contents the importer reads as a block (its field spans paragraphs of the body or of a content control) is kept.
  const document = find(root, 'w:document');
  const body = document ? find(childrenOf(document), 'w:body') : undefined;
  const kept = body ? tocLists(childrenOf(body)).reduce((n, list) => n + findTocRanges(list).length, 0) : 0;
  const unread = Math.max(0, tocs - kept);
  total -= tocs - unread;
  if (unread > 0) kinds.set('TOC', unread);
  else kinds.delete('TOC');
  return item(
    'docx.fields',
    'field(s) (table of contents not read as one, index)',
    total,
    'a table of contents that could not be read as a block and an index are kept as static text',
    kinds,
    20
  );
}

/** `full` adds the formatting/structure detectors and the per-kind field breakdown; otherwise only the fields item (counts only). */
// ---------------------------------------------------------------------------
// Round 2: long-tail constructs and drawing counts
// ---------------------------------------------------------------------------

const isSectPrChild = (path: readonly string[]) => path[path.length - 1] === 'w:sectPr' && !inRevisionProps(path);

/** Tracked changes the importer does not keep, from `revisionLosses`. */
interface RevisionLosses {
  /** Formatting changes by element name (`rPrChange` of a paragraph mark, `tblPrChange`, `sectPrChange`). */
  formatting: Map<string, number>;
}

/** Revision elements of tables, rows, cells and sections: none is kept (an inserted, deleted or merged cell is imported as an ordinary cell). */
const CELL_AND_TABLE_REVISIONS = ['w:tblPrChange', 'w:sectPrChange', 'w:trPrChange', 'w:tcPrChange', 'w:cellIns', 'w:cellDel', 'w:cellMerge'];

/** Paragraph properties a paragraph change does not carry the old value of. */
const PPR_NOT_PREVIOUS = ['w:numPr', 'w:tabs', 'w:pBdr', 'w:shd'];

/**
 * The tracked changes that are lost. Insertions, deletions and moves around runs, on paragraph marks (also of a table
 * cell's only paragraph) and on table rows, and formatting changes of runs and paragraphs are kept. What is left:
 * formatting changes of a paragraph mark's own formatting, and changes of tables, table rows and cells, and sections.
 * Text boxes are reported as such and not counted here.
 */

function revisionLosses(root: XmlNode[]): RevisionLosses {
  const out: RevisionLosses = { formatting: new Map() };
  walk(root, (node, tag, path) => {
    if (path.includes('w:txbxContent')) return;
    if (CELL_AND_TABLE_REVISIONS.includes(tag)) {
      bump(out.formatting, tag.slice(2));
      return;
    }
    // `previous` of a paragraph change carries alignment, spacing, indents and keep flags; a change of numbering, tab stops,
    // borders or shading is not in it, so it cannot be restored by rejecting the change.
    if (tag === 'w:pPr' && !inRevisionProps(path)) {
      const change = find(childrenOf(node), 'w:pPrChange');
      const before = change ? find(childrenOf(change), 'w:pPr') : undefined;
      if (change && PPR_NOT_PREVIOUS.some((name) => JSON.stringify(find(childrenOf(node), name)) !== JSON.stringify(before && find(childrenOf(before), name)))) {
        bump(out.formatting, 'pPrChange (numbering, tabs, borders or shading)');
      }
      return;
    }
    const inParagraphMark = path[path.length - 1] === 'w:rPr' && path[path.length - 2] === 'w:pPr' && !inRevisionProps(path);
    if (tag === 'w:rPrChange' && inParagraphMark) bump(out.formatting, 'rPrChange');
  });
  return out;
}

const CHANGE_NOT_KEPT_RUN_CONTENT = ['w:fldChar', 'w:fldSimple', 'w:drawing', 'mc:AlternateContent', 'w:footnoteReference', 'w:endnoteReference'];

/** The elements that wrap runs in a tracked change; a move is an insertion (`moveTo`) or a deletion (`moveFrom`) as far as the content goes. */
const RUN_CHANGE_TAGS = ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo'];
const DELETING_TAGS = ['w:del', 'w:moveFrom'];

/**
 * Tracked changes around runs that the importer cannot keep as a change: a change around a field (a deleted field is dropped),
 * a deleted image or note reference (dropped), an inserted image, field or note reference (kept as ordinary content), and
 * deleted text inside a field's result (dropped). The reader also warns about each of these.
 */
function trackedChangesNotKept(root: XmlNode[]): Item[] {
  let count = 0;
  // Field nesting in document order: true once a field is past its separator (inside its result).
  const fields: boolean[] = [];
  const kinds = new Map<string, number>();
  walk(root, (node, tag, path) => {
    if (tag === 'w:fldChar') {
      const type = attr(node, 'w:fldCharType');
      if (type === 'begin') fields.push(false);
      else if (type === 'separate' && fields.length > 0) fields[fields.length - 1] = true;
      else if (type === 'end') fields.pop();
      return;
    }
    if (!RUN_CHANGE_TAGS.includes(tag) || inRevisionProps(path) || path.includes('w:txbxContent')) return;
    const parent = path[path.length - 1];
    if (parent === 'w:rPr' || parent === 'w:trPr' || parent === 'w:body' || parent === 'w:tc') return;
    // A change inside another one (the writer nests a deletion in an insertion) is one change with its outer element.
    const nested = path.some((t) => RUN_CHANGE_TAGS.includes(t));
    const found = nested ? undefined : CHANGE_NOT_KEPT_RUN_CONTENT.find((t) => findDeep(node, t));
    if (found) {
      count++;
      bump(kinds, found === 'w:fldChar' || found === 'w:fldSimple' ? 'field' : found === 'w:footnoteReference' || found === 'w:endnoteReference' ? 'note reference' : 'image');
    } else if (DELETING_TAGS.includes(tag) && !path.some((t) => DELETING_TAGS.includes(t)) && fields.some(Boolean) && findDeep(node, 'w:delText')) {
      count++;
      bump(kinds, 'text in a field result');
    }
  });
  return item('docx.trackedChangesNotKept', 'tracked change(s) around a field, image or note reference, or inside a field result', count, 'a deleted field, image or note reference is dropped and an inserted one is kept as ordinary content; deleted text in a field result is dropped; the change is not recorded', kinds);
}

/**
 * Tracked changes the importer does not keep because of where they are: around whole paragraphs or tables (a direct child of
 * the body or a table cell, whose content is not imported), and in headers and footers (deleted text is dropped, inserted
 * text is kept as ordinary text).
 */
function trackedChangesElsewhere(root: XmlNode[], headerFooters: string[]): Item[] {
  let blocks = 0;
  walk(root, (_n, tag, path) => {
    if ((tag !== 'w:ins' && tag !== 'w:del' && tag !== 'w:moveFrom' && tag !== 'w:moveTo') || inRevisionProps(path) || path.includes('w:txbxContent')) return;
    const parent = path[path.length - 1];
    if (parent === 'w:body' || parent === 'w:tc') blocks++;
  });
  // A paragraph that is read as a page break or a horizontal rule has no place for a change of its paragraph mark.
  let structural = 0;
  walk(root, (node, tag, path) => {
    if (tag !== 'w:p' || path.includes('w:tc') || path.includes('w:txbxContent')) return;
    const pPr = find(childrenOf(node), 'w:pPr');
    if (!pPr) return;
    const pPrKids = childrenOf(pPr);
    const markProps = find(pPrKids, 'w:rPr');
    const tracked = !!find(pPrKids, 'w:pPrChange') || (!!markProps && childrenOf(markProps).some((k) => ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo'].includes(tagOf(k) ?? '')));
    if (!tracked || ['w:t', 'w:drawing', 'w:pict', 'mc:AlternateContent', 'w:fldChar', 'w:fldSimple', 'w:tab'].some((t) => findDeep(node, t))) return;
    const pageBreak = findAll(childrenOf(node), 'w:r').some((r) => childrenOf(r).some((k) => tagOf(k) === 'w:br' && attrsOf(k)['@_w:type'] === 'page'));
    const pBdr = find(pPrKids, 'w:pBdr');
    if (pageBreak || (pBdr && find(childrenOf(pBdr), 'w:bottom'))) structural++;
  });
  let headers = 0;
  for (const xml of headerFooters) {
    walk(parseXml(xml), (_n, tag, path) => {
      if ((tag === 'w:ins' || tag === 'w:del' || tag === 'w:moveFrom' || tag === 'w:moveTo' || tag === 'w:rPrChange' || tag === 'w:pPrChange') && !inRevisionProps(path)) headers++;
    });
  }
  return [
    ...item('docx.trackedBlockChanges', 'tracked change(s) around whole paragraphs or tables', blocks, 'the paragraphs or tables inside are not imported'),
    ...item('docx.trackedStructureChanges', 'tracked change(s) of the paragraph mark of a page break or horizontal rule', structural, 'the page break or rule is imported without the change'),
    ...item('docx.trackedHeaderFooterChanges', 'tracked change(s) in headers and footers', headers, 'deleted text is dropped and inserted text is kept as ordinary text; the change is not recorded'),
  ];
}

function formattingRevisions(losses: RevisionLosses): Item[] {
  const total = [...losses.formatting.values()].reduce((sum, n) => sum + n, 0);
  return item('docx.formattingRevisions', 'tracked formatting change(s)', total, 'the change history is dropped; the current formatting is imported', losses.formatting);
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

/** Content controls that are read as their content only (see `unsupportedControls`); the others are kept. */
function contentControls(root: XmlNode[]): Item[] {
  const reasons = unsupportedControls(root);
  const total = [...reasons.values()].reduce((sum, n) => sum + n, 0);
  return item('docx.contentControls', 'content control(s)', total, 'their content is kept; the control itself (repeating section, data binding, row or cell wrapper, ...) is lost', reasons);
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
  return item('docx.legacyForms', 'legacy form field(s)', total, 'their displayed text is kept as plain text; the form field (checkbox state, list, text limits) is lost', kinds);
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
  const safe = (fn: () => Item[]): Item[] => {
    try {
      return fn();
    } catch {
      return [];
    }
  };
  let losses: RevisionLosses = { formatting: new Map() };
  try {
    losses = revisionLosses(root);
  } catch {
    // A detector that cannot read the tree reports nothing (like the others).
  }
  const drawings = safe(() => {
    const { charts, floating } = drawingCounts(root);
    return [...floating, ...charts];
  });
  if (!full) return [...drawings, ...safe(() => trackedChangesNotKept(root)), ...safe(() => trackedChangesElsewhere(root, input.headerFooters)), ...safe(() => contentControls(root)).map(({ where: _where, ...rest }) => rest), ...safe(() => fields(root)).map(({ where: _where, ...rest }) => rest)];
  return [
    ...drawings,
    ...safe(() => trackedChangesNotKept(root)),
    ...safe(() => trackedChangesElsewhere(root, input.headerFooters)),
    ...safe(() => contentControls(root)),
    ...safe(() => listFormatting(root, numberingLookup(input.numbering))),
    ...safe(() => sectionDetectors(root)),
    ...safe(() => fields(root)),
    ...safe(() => formattingRevisions(losses)),
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
