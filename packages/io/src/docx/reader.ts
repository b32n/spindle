import JSZip from 'jszip';
import type {
  DocBlock,
  DocHeaderFooter,
  DocHeaderFooterParagraph,
  DocImageRun,
  DocInline,
  DocParagraphStyle,
  DocSection,
  DocRevision,
  DocRevisionMeta,
  DocTextRun,
  DocBorders,
  DocCellTextDirection,
  DocColumns,
  DocPageNumbering,
  DocRunBase,
  DocSectionBreakType,
  DocSectionVerticalAlign,
  DocTableRow,
  DocumentImportResult,
} from '@b32nio/spindle-shared/export/docx';

import { createCommentTracker, parseComments, type CommentTracker } from './comments';
import { isKeptBlockControl, parseSdt, sdtContent } from './content-controls';
import { parseNotes, type NoteKind } from './notes';
import { findTocRanges, isEmptyParagraph, withoutMarkers, type TocRange } from './toc-field';
import { drawingKind, type DrawingKind } from './drawing-kind';
import { ListCounter, listTypeFor, parseNumbering, type NumberingInfo } from './numbering';
import { parsePPr, StyleSheet, toDocParagraphStyle } from './styles';
import { applicableParts, parseBorders, parseLook, parseShading, parseVAlign, pxFromTwips } from './table-props';
import { parseTheme, type Theme } from './theme';
import { attrsOf, childrenOf, find, findAll, findDeep, parseXml, tagOf, type XmlNode } from './xml-tree';

/** Depth-first concatenation of every #text leaf under a node (a run's <w:t>, a field's cached-result text, ...). */
function textOf(node: XmlNode): string {
  let out = '';
  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (tag === '#text') out += String((child as Record<string, unknown>)['#text'] ?? '');
    else out += textOf(child);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Relationships (word/_rels/document.xml.rels and friends): rId -> target.
// A hyperlink's URL, and a header/footer's part filename, are never inline —
// always one level of indirection through here.
// ---------------------------------------------------------------------------

function parseRels(xml: string | undefined): Map<string, string> {
  const rels = new Map<string, string>();
  if (!xml) return rels;
  const root = parseXml(xml);
  const relationships = find(root, 'Relationships');
  if (!relationships) return rels;
  for (const rel of findAll(childrenOf(relationships), 'Relationship')) {
    const a = attrsOf(rel);
    if (a['@_Id'] && a['@_Target']) rels.set(a['@_Id'], a['@_Target']);
  }
  return rels;
}

/** Everything paragraph/table/header/footer parsing needs to resolve a part's cross-references: hyperlink and header/footer targets, plus embedded images pre-loaded as data: URIs (keyed by rId, since loading them is async and the parse tree walk below it isn't). */
interface PartContext {
  rels: Map<string, string>;
  media: Map<string, string>;
  styles: StyleSheet;
  /** The theme, for the colours of table borders and shading. */
  theme?: Theme;
  /** Text properties of the table style that reach the cell being read (under the paragraph and run formatting). */
  tableText?: DocRunBase;
  /** Follows list numbering through the body (absent for headers/footers). */
  lists?: ListCounter;
  /** Comment ranges open at the current position of the body walk (absent for headers/footers). */
  comments?: CommentTracker;
  /** Content controls are kept as wrappers (the body); headers and footers read them as their content only. */
  controls?: boolean;
  /** The text of the notes the body refers to (absent in notes, headers and footers, where a note reference is not kept). */
  notes?: NoteLookup;
  /** Numbers the tracked changes that have no id of their own, across the body and the notes of one import (absent in headers and footers, which keep no changes). */
  anonymousRevisions?: { next: number };
}

/**
 * The note a reference points at (by its id in the file), read on first use: its text, and the id it gets in the
 * model, which counts the references of its kind in document order. Undefined when the document has no such note.
 */
type NoteLookup = (kind: NoteKind, id: string) => { id: string; content: DocInline[] } | undefined;

const MEDIA_EXT_TO_MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

async function buildMediaMap(zip: JSZip, rels: Map<string, string>): Promise<Map<string, string>> {
  const media = new Map<string, string>();
  for (const [rId, target] of rels) {
    if (!/(^|\/)media\//i.test(target)) continue;
    const extMatch = target.match(/\.([a-zA-Z]+)$/);
    const mime = extMatch ? MEDIA_EXT_TO_MIME[extMatch[1].toLowerCase()] : undefined;
    if (!mime) continue;
    const file = zip.file(`word/${target}`);
    if (!file) continue;
    const base64 = await file.async('base64');
    media.set(rId, `data:${mime};base64,${base64}`);
  }
  return media;
}

// ---------------------------------------------------------------------------
// Runs and inline content
// ---------------------------------------------------------------------------

const NON_BREAKING_HYPHEN = '\u2011';
const SOFT_HYPHEN = '\u00ad';

/** Content controls and custom XML wrap block or inline content: replace each wrapper by what it contains (its properties are dropped). */
function unwrapContainer(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (tag === 'w:sdt') return sdtContent(node);
  return childrenOf(node).filter((k) => tagOf(k) !== 'w:customXmlPr');
}

/**
 * Block-level children with w:sdt / w:customXml wrappers replaced by their content, at any depth. With `keepControls`
 * (the body and table cells) a content control that is kept as a wrapper stays in place (see `isKeptBlockControl`).
 */
function flattenBlocks(nodes: XmlNode[], keepControls = false): XmlNode[] {
  return nodes.flatMap((n) => {
    const tag = tagOf(n);
    if (tag === 'w:sdt' && keepControls && isKeptBlockControl(n)) return [n];
    return tag === 'w:sdt' || tag === 'w:customXml' ? flattenBlocks(unwrapContainer(n), keepControls) : [n];
  });
}

/** Paragraph-level children with w:customXml and w:smartTag wrappers replaced by their content, at any depth. Content controls and tracked changes (w:ins, w:del) are left to the caller. */
function flattenInline(nodes: XmlNode[]): XmlNode[] {
  return nodes.flatMap((n) => {
    const tag = tagOf(n);
    if (tag === 'w:customXml') return flattenInline(unwrapContainer(n));
    if (tag === 'w:smartTag') return flattenInline(childrenOf(n).filter((k) => tagOf(k) !== 'w:smartTagPr'));
    return [n];
  });
}

/** The tracked change a `w:ins` / `w:del` element records. A change without an id gets the next number of its import. */
function revisionMeta(node: XmlNode, fallbackId: () => string): DocRevisionMeta {
  const attrs = attrsOf(node);
  const date = attrs['@_w:date'];
  return { id: attrs['@_w:id'] || fallbackId(), author: attrs['@_w:author'] || 'Unknown', ...(date && { date }) };
}

/**
 * Walk a paragraph's direct children, producing inline content. Handles
 * plain runs, hyperlinks (resolved through rels) and field-code sequences:
 * PAGE/NUMPAGES keep today's page-number mapping, HYPERLINK fields become
 * links, tables of contents and indexes keep their result as static text,
 * legacy form fields are dropped with a warning, and every other field is
 * kept as a generic `field` carrying its instruction and the result Word
 * last computed (nested fields end up as the outer field's cached content).
 *
 * A field is a fldChar "begin", an instrText carrying the instruction, an
 * optional "separate" marker, optional cached display text, then a fldChar
 * "end" — spread across sibling <w:r> elements per the OOXML spec, but our
 * own writer (and possibly other producers) packs the whole sequence into
 * a single <w:r>'s children instead. Walking element-by-element within
 * each run, with field state carried across runs, handles both shapes.
 * <w:fldSimple> is the shorthand form some producers use in place of the
 * begin/end sequence. A run's <w:drawing> (inline image) is resolved
 * through ctx.media, pre-loaded as a data: URI since loading it is async
 * and this tree walk isn't.
 */
function paragraphInlineContent(pNode: XmlNode, ctx: PartContext, warnings: string[]): DocInline[] {
  let out: DocInline[] = [];
  const pStyleId = paragraphStyleId(pNode);
  ctx.comments?.paragraphStart();
  /** Tracked changes (`w:ins`, `w:del`) the content being read is inside of. */
  let revision: DocRevision | undefined;
  const deleted = () => revision?.deletion !== undefined;
  const fields: OpenField[] = []; // begin/.../end sequences in progress, outermost first
  const top = (): OpenField | undefined => fields[fields.length - 1];
  const inInstruction = () => top()?.phase === 'instruction';

  /** An image goes into the cached result of the field it is in, else into the paragraph. */
  const place = (image: DocInline) => {
    const open = top();
    if (!open) {
      if (revision?.insertion) warnings.push('A tracked insertion of an image was not kept as a change.');
      out.push(image);
    } else if (open.phase === 'result' && image.type === 'image') open.result.push(image);
  };

  /** A bookmark is kept as a point where it starts; none inside a field (the field keeps only its result). */
  const placeBookmark = (node: XmlNode) => {
    const name = bookmarkName(node);
    if (name && !top()) out.push({ type: 'bookmark', name });
  };

  /** A field whose end was reached: nested in another field it becomes part of that field, else it is written to the paragraph. */
  const finishField = (done: OpenField) => {
    if (done.revision) {
      // A change around a field is not kept as a change: an inserted field stays as it is, a deleted one is dropped (as it was before changes were kept).
      warnings.push(`A tracked ${done.revision.deletion ? 'deletion' : 'insertion'} around the field "${done.instruction.trim()}" was not kept as a change${done.revision.deletion ? '; the field was dropped' : ''}.`);
      if (done.revision.deletion) return;
    }
    const parent = top();
    if (!parent) {
      emitField(done, out);
    } else if (parent.phase === 'result') {
      parent.result.push(...done.result);
    } else {
      // A field inside another field's instruction (IF { PAGE } > 1 ...) cannot be kept as such: its cached value stands in.
      parent.instruction += resultText(done.result);
      warnings.push(`A field nested in the instruction of "${parent.instruction.trim()}" was kept as its cached value.`);
    }
  };

  const consumeRun = (runNode: XmlNode, target: string | undefined) => {
    const rPr = find(childrenOf(runNode), 'w:rPr');
    const rStyleEl = rPr ? find(childrenOf(rPr), 'w:rStyle') : undefined;
    // `Hyperlink` on a link is Word's implicit link style, which the writer applies to every link again.
    const characterStyle = ctx.styles.namedCharacterStyle(rStyleEl ? attrsOf(rStyleEl)['@_w:val'] : undefined);
    const namedStyleId = target && characterStyle === 'Hyperlink' ? undefined : characterStyle;
    const props = { ...ctx.styles.run(rPr, pStyleId, ctx.tableText), ...(namedStyleId && { namedStyleId }), ...(revision && !top() && { revision }) };
    let text = '';
    /** A reference followed by its custom mark (`w:customMarkFollows`): the text after it in the run is that mark. */
    let markOf: Extract<DocInline, { type: 'noteRef' }> | undefined;
    const flush = () => {
      if (!text) return;
      const commentThreadId = ctx.comments?.anchorText(text);
      const open = top();
      if (open) {
        // Inside a field the text is its cached result (nothing inside an instruction is displayed). A change does not
        // reach into a result: deleted text is dropped, inserted text stays as plain result text.
        if (open.phase === 'result') {
          if (deleted()) {
            if (!open.warnedDeleted) warnings.push(`Tracked deleted text inside the field "${open.instruction.trim()}" was not kept; it was dropped.`);
            open.warnedDeleted = true;
          } else {
            const { revision: _revision, ...plain } = props;
            open.result.push({ type: 'text', text, ...plain });
          }
        }
        text = '';
        return;
      }
      const runProps = commentThreadId ? { ...props, commentThreadId } : props;
      if (target) out.push({ type: 'link', text, href: target, ...runProps });
      else out.push({ type: 'text', text, ...runProps });
      text = '';
    };
    for (const piece of childrenOf(runNode)) {
      const pieceTag = tagOf(piece);
      if (pieceTag === 'w:fldChar') {
        const type = attrsOf(piece)['@_w:fldCharType'];
        flush();
        if (type === 'begin') {
          const dirty = attrsOf(piece)['@_w:dirty'];
          fields.push({ instruction: '', phase: 'instruction', result: [], dirty: dirty === 'true' || dirty === '1', revision });
        } else if (type === 'separate') {
          const open = top();
          if (open) open.phase = 'result';
        } else if (type === 'end') {
          const done = fields.pop();
          if (done) finishField(done);
        }
      } else if (pieceTag === 'w:instrText' || pieceTag === 'w:delInstrText') {
        const open = top();
        if (open?.phase === 'instruction') open.instruction += textOf(piece);
      } else if (pieceTag === 'w:t' || pieceTag === 'w:delText') {
        if (markOf) markOf.customMark = (markOf.customMark ?? '') + textOf(piece);
        else if (!inInstruction()) text += textOf(piece);
      } else if (pieceTag === 'w:footnoteReference' || pieceTag === 'w:endnoteReference') {
        flush();
        const kind: NoteKind = pieceTag === 'w:footnoteReference' ? 'footnote' : 'endnote';
        const id = attrsOf(piece)['@_w:id'];
        // A reference in a field's result, in a header or footer, or inside a note is not kept.
        if (!ctx.notes) {
          warnings.push(`A ${kind} reference in a header, footer or note was not imported.`);
        } else if (deleted()) {
          warnings.push(`A deleted ${kind} reference was not kept as a tracked deletion; it was dropped.`);
        } else if (!top() && id !== undefined) {
          const note = ctx.notes(kind, id);
          if (note) {
            if (revision?.insertion) warnings.push(`A tracked insertion of a ${kind} reference was not kept as a change.`);
            const ref: Extract<DocInline, { type: 'noteRef' }> = { type: 'noteRef', kind, id: note.id, content: JSON.parse(JSON.stringify(note.content)) as DocInline[] };
            out.push(ref);
            const follows = attrsOf(piece)['@_w:customMarkFollows'];
            if (follows === '1' || follows === 'true') markOf = ref;
          } else {
            warnings.push(`A ${kind} reference points to a ${kind} that does not exist.`);
          }
        }
      } else if (pieceTag === 'w:tab' || pieceTag === 'w:ptab') {
        if (!inInstruction()) text += '\t';
      } else if (pieceTag === 'w:br') {
        // A page break is a structural block (see pageBreakPlacement); every other break is a line break.
        if (!inInstruction() && attrsOf(piece)['@_w:type'] !== 'page') text += '\n';
      } else if (pieceTag === 'w:cr') {
        if (!inInstruction()) text += '\n';
      } else if (pieceTag === 'w:noBreakHyphen') {
        if (!inInstruction()) text += NON_BREAKING_HYPHEN;
      } else if (pieceTag === 'w:softHyphen') {
        if (!inInstruction()) text += SOFT_HYPHEN;
      } else if (pieceTag === 'w:ruby') {
        // Keep the base text; the reading marks (w:rt) are not imported.
        flush();
        const base = find(childrenOf(piece), 'w:rubyBase');
        if (base) consumeRuns(childrenOf(base), target);
      } else if ((pieceTag === 'w:drawing' || pieceTag === 'mc:AlternateContent') && deleted()) {
        warnings.push('A deleted image was not kept as a tracked deletion; it was dropped.');
      } else if (pieceTag === 'w:drawing') {
        const image = inInstruction() ? null : parseDrawing(piece, ctx.media, warnings);
        if (image) place(image);
      } else if (pieceTag === 'mc:AlternateContent') {
        // LibreOffice (and Word for some shapes) wrap the drawing in
        // <mc:AlternateContent><mc:Choice><w:drawing>; the Fallback is legacy VML.
        const drawing = findDeep(piece, 'w:drawing');
        const image = drawing && !inInstruction() ? parseDrawing(drawing, ctx.media, warnings) : null;
        if (image) place(image);
      }
    }
    flush();
    // A note with a custom mark starts with that mark as text (not with the reference mark): it is not part of the note's text.
    if (markOf?.customMark) dropCustomMark(markOf.content, markOf.customMark);
  };

  /** A run-level content control wraps what it contains; one inside a field's result, or in a header or footer, is read as its content only. */
  const consumeControl = (sdt: XmlNode, target: string | undefined) => {
    const parsed = ctx.controls && !top() ? parseSdt(sdt) : undefined;
    if (!parsed || !('control' in parsed)) {
      consumeRuns(sdtContent(sdt), target);
      return;
    }
    const outer = out;
    out = [];
    consumeRuns(sdtContent(sdt), target);
    const content = out;
    out = outer;
    out.push({ type: 'contentControl', control: parsed.control, content });
  };

  const consumeRuns = (runs: XmlNode[], target: string | undefined) => {
    for (const child of flattenInline(runs)) {
      const tag = tagOf(child);
      if (tag === 'w:ins' || tag === 'w:del') {
        const outer = revision;
        const meta = revisionMeta(child, () => `r${++(ctx.anonymousRevisions ??= { next: 0 }).next}`);
        revision = { ...outer, ...(tag === 'w:ins' ? { insertion: meta } : { deletion: meta }) };
        consumeRuns(childrenOf(child), target);
        revision = outer;
      } else if (tag === 'w:fldSimple') {
        const dirty = attrsOf(child)['@_w:dirty'];
        fields.push({ instruction: attrsOf(child)['@_w:instr'] ?? '', phase: 'result', result: [], dirty: dirty === 'true' || dirty === '1', revision });
        consumeRuns(childrenOf(child), target);
        const done = fields.pop();
        if (done) finishField(done);
      } else if (tag === 'w:hyperlink') {
        consumeRuns(childrenOf(child), resolveHyperlinkTarget(child, ctx.rels, warnings));
      } else if (tag === 'w:sdt') consumeControl(child, target);
      else if (tag === 'w:r') consumeRun(child, target);
      else if (tag === 'w:bookmarkStart') placeBookmark(child);
      else ctx.comments?.mark(child);
    }
  };

  consumeRuns(childrenOf(pNode), undefined);

  // A field still open here continues in a later paragraph (a table of contents, for instance): what it displayed
  // so far stays as plain text. Anything open inside it is part of that text.
  while (fields.length > 1) {
    const inner = fields.pop() as OpenField;
    const outer = top() as OpenField;
    if (outer.phase === 'result') outer.result.push(...inner.result);
  }
  if (fields[0]?.phase === 'result') out.push(...fields[0].result);

  return out;
}

const EMU_TO_PX = 1 / 9525; // 914400 EMU/inch, 96px/inch

const DRAWING_KIND_WARNINGS: Record<Exclude<DrawingKind, 'picture'>, string> = {
  chart: 'A chart was not imported — charts are not supported yet.',
  diagram: 'A SmartArt diagram was not imported — diagrams are not supported yet.',
  shape: 'A shape or text box was not imported — shapes and text boxes are not supported yet.',
  embeddedObject: 'An embedded object preview was not imported — embedded objects are not supported yet.',
  unknown: 'A drawing object of an unsupported kind was not imported.',
};

/**
 * A <w:drawing> run child. Only inline pictures (<wp:inline> + <a:blip r:embed>) become images. Anchored
 * pictures aren't positioned the way our block model supports, so they're dropped with a warning; every
 * other kind (chart, SmartArt, shape/text box, embedded-object preview) is reported as what it is rather
 * than as a missing image. Linked pictures (<a:blip r:link>) point outside the package and are never fetched.
 */
function parseDrawing(drawingNode: XmlNode, media: Map<string, string>, warnings: string[]): DocInline | null {
  const kind = drawingKind(drawingNode);
  if (kind !== 'picture') {
    warnings.push(DRAWING_KIND_WARNINGS[kind]);
    return null;
  }
  const inline = findDeep(drawingNode, 'wp:inline');
  if (!inline) {
    warnings.push('A floating (anchored) image was skipped — only inline images are supported.');
    return null;
  }
  const extent = findDeep(inline, 'wp:extent');
  const extentAttrs = extent ? attrsOf(extent) : {};
  const width = extentAttrs['@_cx'] ? Math.round(Number(extentAttrs['@_cx']) * EMU_TO_PX) : 0;
  const height = extentAttrs['@_cy'] ? Math.round(Number(extentAttrs['@_cy']) * EMU_TO_PX) : 0;
  const docPr = findDeep(inline, 'wp:docPr');
  const docPrAttrs = docPr ? attrsOf(docPr) : {};
  const alt = docPrAttrs['@_descr'] || docPrAttrs['@_name'] || undefined;
  const blip = findDeep(inline, 'a:blip');
  const blipAttrs = blip ? attrsOf(blip) : {};
  const embedId = blipAttrs['@_r:embed'];
  const src = embedId ? media.get(embedId) : undefined;
  if (!src) {
    if (!embedId && blipAttrs['@_r:link']) {
      warnings.push('A linked picture was skipped — images linked from outside the document are never fetched.');
    } else {
      warnings.push('An embedded image could not be resolved and was skipped.');
    }
    return null;
  }
  return { type: 'image', src, width, height, ...(alt && { alt }) };
}

/** A field being read: `phase` says whether its instruction or its cached result is being collected. */
interface OpenField {
  instruction: string;
  phase: 'instruction' | 'result';
  result: Array<DocTextRun | DocImageRun>;
  dirty: boolean;
  /** The tracked changes the field began inside of. */
  revision?: DocRevision;
  /** Set once deleted text in its result has been warned about, so a result of several runs warns once. */
  warnedDeleted?: boolean;
}

/** Fields whose result is block content (a table of contents, an index): the text stays, the field does not. */
const STATIC_RESULT_FIELDS = new Set(['TOC', 'TOA', 'INDEX']);
/** Legacy form fields (`FORMTEXT` and the like): the text they display stays, as plain text; the form behaviour (`w:ffData`) does not. */
const LEGACY_FORM_FIELDS = new Set(['FORMTEXT', 'FORMCHECKBOX', 'FORMDROPDOWN']);

function resultText(result: Array<DocTextRun | DocImageRun>): string {
  return result.map((r) => (r.type === 'text' ? r.text : '')).join('');
}

/** The field type: the first word of the instruction, upper case. */
function fieldKind(instruction: string): string {
  return instruction.trim().toUpperCase().split(/[\s\\"]/)[0];
}

/** `HYPERLINK "url" \l "anchor"`: the target of a link field (`#anchor` when it has only an anchor), or undefined for anything else. */
function hyperlinkFieldTarget(instruction: string): string | undefined {
  const args = [...instruction.trim().matchAll(/"([^"]*)"|(\S+)/g)].map((m) => ({ value: m[1] ?? m[2], quoted: m[1] !== undefined }));
  const rest = args.slice(1);
  let url: string | undefined;
  let anchor: string | undefined;
  for (let i = 0; i < rest.length; i++) {
    const arg = rest[i];
    if (!arg.quoted && arg.value.startsWith('\\')) {
      // Switches that take a value: \l (location) \t (target frame) \m (map); \n and the rest are flags.
      const takesValue = /^\\[ltm]$/i.test(arg.value);
      if (takesValue && rest[i + 1]) {
        if (/^\\l$/i.test(arg.value)) anchor = rest[i + 1].value;
        i++;
      }
    } else if (url === undefined) {
      url = arg.value;
    }
  }
  if (!url) return anchor ? `#${anchor}` : undefined;
  return anchor ? `${url}#${anchor}` : url;
}

/** Whether a cached result run carries nothing but its text. */
function isPlainText(run: DocTextRun | DocImageRun): run is DocTextRun {
  return run.type === 'text' && Object.keys(run).every((k) => k === 'type' || k === 'text');
}

/** Writes a finished, outermost field into the paragraph (see paragraphInlineContent for what each kind becomes). */
function emitField(field: OpenField, out: DocInline[]): void {
  const instruction = field.instruction.trim();
  const kind = fieldKind(instruction);
  if (!kind) {
    out.push(...field.result);
    return;
  }
  if (kind === 'PAGE') {
    out.push({ type: 'pageNumber' });
    return;
  }
  if (kind === 'NUMPAGES') {
    out.push({ type: 'totalPages' });
    return;
  }
  if (STATIC_RESULT_FIELDS.has(kind) || LEGACY_FORM_FIELDS.has(kind)) {
    out.push(...field.result);
    return;
  }
  if (kind === 'HYPERLINK') {
    const href = hyperlinkFieldTarget(instruction);
    if (href && resultText(field.result)) {
      for (const run of field.result) {
        if (run.type === 'text') {
          const { namedStyleId, ...props } = run;
          out.push({ ...props, type: 'link', href, ...(namedStyleId && namedStyleId !== 'Hyperlink' && { namedStyleId }) });
        } else out.push(run);
      }
      return;
    }
  }
  const plain = field.result.every(isPlainText);
  out.push({
    type: 'field',
    instruction,
    kind,
    result: plain ? resultText(field.result) : field.result,
    ...(field.dirty && { dirty: true }),
  });
}

/** A bookmark's name, or undefined for Word's own cursor marker (`_GoBack`) and an unnamed one. */
function bookmarkName(node: XmlNode): string | undefined {
  const name = attrsOf(node)['@_w:name'];
  return name && name !== '_GoBack' ? name : undefined;
}

/**
 * Where a `w:hyperlink` points: its relationship's target (resolved through rels), `#anchor` for a link to a
 * bookmark in the document, or the target with the anchor appended when it has both.
 */
function resolveHyperlinkTarget(hyperlinkNode: XmlNode, rels: Map<string, string>, warnings: string[]): string | undefined {
  const attrs = attrsOf(hyperlinkNode);
  const rId = attrs['@_r:id'];
  const anchor = attrs['@_w:anchor'];
  let target: string | undefined;
  if (rId) {
    target = rels.get(rId);
    if (!target) warnings.push(`Hyperlink relationship "${rId}" could not be resolved.`);
  }
  if (!anchor) return target;
  return target ? `${target}#${anchor}` : `#${anchor}`;
}

// ---------------------------------------------------------------------------
// Paragraphs (-> paragraph / heading / listItem blocks)
// ---------------------------------------------------------------------------

/** The id of a paragraph's own style (`w:pStyle`), if it names one. */
function paragraphStyleId(pNode: XmlNode): string | undefined {
  const pPr = find(childrenOf(pNode), 'w:pPr');
  const el = pPr ? find(childrenOf(pPr), 'w:pStyle') : undefined;
  return el ? attrsOf(el)['@_w:val'] : undefined;
}

function hasPageBreakRun(pNode: XmlNode): boolean {
  return childrenOf(pNode).some((child) => {
    if (tagOf(child) !== 'w:r') return false;
    const br = find(childrenOf(child), 'w:br');
    return !!br && attrsOf(br)['@_w:type'] === 'page';
  });
}

/**
 * Where a page-break run sits relative to the paragraph's other content:
 * 'before' (nothing but empty runs precede it), 'after' otherwise, or null when
 * the paragraph has no page-break run. LibreOffice stores the break after the
 * content it follows (`<w:r><w:br w:type="page"/></w:r>` at the end of the
 * previous paragraph); Word's "page break before" style puts it first.
 */
function pageBreakPlacement(pNode: XmlNode): 'before' | 'after' | null {
  let sawContent = false;
  for (const child of childrenOf(pNode)) {
    const tag = tagOf(child);
    if (tag === 'w:r') {
      const runKids = childrenOf(child);
      const br = find(runKids, 'w:br');
      if (br && attrsOf(br)['@_w:type'] === 'page') return sawContent ? 'after' : 'before';
      if (runKids.some((k) => ['w:t', 'w:drawing', 'w:pict', 'w:tab'].includes(tagOf(k) ?? ''))) sawContent = true;
    } else if (tag === 'w:hyperlink' || tag === 'w:ins' || tag === 'w:del' || tag === 'w:sdt' || tag === 'w:smartTag') {
      sawContent = true;
    }
  }
  return null;
}

/**
 * Each numbered list item carries the number it shows (`start`), but the model keeps `start` only where a list
 * begins at something other than 1 or restarts: on the first item of a run of numbered items, and on a later
 * top-level item whose number is not simply the next one (a new list instance). Elsewhere the number follows from
 * the order.
 */
function settleListStarts(blocks: DocBlock[]): DocBlock[] {
  let inRun = false;
  let lastTopLevel = 0;
  for (const block of blocks) {
    if (block.type !== 'listItem' || block.listType !== 'numbered') {
      inRun = false;
      continue;
    }
    if (block.level === 0 && block.start !== undefined) {
      const n = block.start;
      const keep = inRun ? n !== lastTopLevel + 1 : n !== 1;
      lastTopLevel = n;
      if (!keep) delete block.start;
    } else {
      delete block.start;
    }
    inRun = true;
  }
  return blocks;
}

/** parseParagraph plus the page break it may carry alongside its content. */
function parseParagraphBlocks(
  pNode: XmlNode,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[]
): DocBlock[] {
  const block = parseParagraph(pNode, numbering, ctx, warnings);
  if (block.type === 'pageBreak') return [block];
  const placement = pageBreakPlacement(pNode);
  if (!placement) return [block];
  return placement === 'before' ? [{ type: 'pageBreak' }, block] : [block, { type: 'pageBreak' }];
}

function hasBottomBorder(pPrKids: XmlNode[]): boolean {
  const pBdr = find(pPrKids, 'w:pBdr');
  return !!pBdr && !!find(childrenOf(pBdr), 'w:bottom');
}

/** `plain`: a paragraph of a table cell, which stays a paragraph (inline image, no page break or rule block) rather than being promoted to a structural block. */
function parseParagraph(
  pNode: XmlNode,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[],
  plain = false
): DocBlock {
  const kids = childrenOf(pNode);
  const pPr = find(kids, 'w:pPr');
  const pPrKids = pPr ? childrenOf(pPr) : [];
  const content = paragraphInlineContent(pNode, ctx, warnings);
  // Effective formatting: docDefaults, then the paragraph's style chain, then its own properties.
  const direct = parsePPr(pPr);
  const effective = ctx.styles.paragraph(paragraphStyleId(pNode), direct);
  const namedStyleId = ctx.styles.namedParagraphStyle(paragraphStyleId(pNode));
  const style: DocParagraphStyle = { ...toDocParagraphStyle(effective), ...(namedStyleId && { namedStyleId }) };

  // An otherwise-empty paragraph carrying only a page-break run or a
  // bottom-border (our own writer's idiom for a horizontal rule) maps back
  // to the corresponding structural block rather than an empty paragraph.
  const onlyBookmarks = content.every((c) => c.type === 'bookmark');
  if (!plain && onlyBookmarks && hasPageBreakRun(pNode)) {
    return { type: 'pageBreak' };
  }
  if (!plain && content.length === 0 && hasBottomBorder(pPrKids)) {
    return { type: 'horizontalRule' };
  }

  // A paragraph whose only content is a single image is how our own writer
  // (and most producers) represents a block-level image — promote it
  // rather than returning a "paragraph" block wrapping one inline image.
  if (!plain && content.length === 1 && content[0].type === 'image') {
    const image = content[0];
    return {
      type: 'image',
      src: image.src,
      width: image.width,
      height: image.height,
      ...(image.alt && { alt: image.alt }),
      ...(style.alignment && style.alignment !== 'justify' ? { alignment: style.alignment } : {}),
    };
  }

  // A list item is a paragraph numbered directly, or through its style. A heading that is numbered through its
  // style (numbered headings) stays a heading.
  const level = effective.ilvl ?? 0;
  const listItem = (numId: string | undefined): DocBlock => {
    const listType = numId ? listTypeFor(numbering, numId, level) : 'bullet';
    // The number the item shows; importDocx keeps it as `start` only where a list begins or restarts.
    const shown = numId && listType === 'numbered' ? ctx.lists?.next(numId, level) : undefined;
    // `ListParagraph` is Word's implicit list style, which the writer applies to every list item again.
    const { namedStyleId: listStyle, ...rest } = style;
    return { type: 'listItem', listType, level, content, ...(shown !== undefined && { start: shown }), ...rest, ...(listStyle && listStyle !== 'ListParagraph' && { namedStyleId: listStyle }) };
  };
  if (find(pPrKids, 'w:numPr') && direct.numId !== '0') return listItem(direct.numId);
  if (effective.headingLevel) return { type: 'heading', level: effective.headingLevel as 1 | 2 | 3 | 4 | 5 | 6, content, ...style };
  if (effective.numId && effective.numId !== '0') return listItem(effective.numId);

  return { type: 'paragraph', content, ...style };
}

// ---------------------------------------------------------------------------
// Tables
// ---------------------------------------------------------------------------

const isOnEl = (el: XmlNode | undefined): boolean => {
  if (!el) return false;
  const v = attrsOf(el)['@_w:val'];
  return !(v === '0' || v === 'false' || v === 'off');
};

const gridWidthPx = (el: XmlNode | undefined): number | undefined => {
  const w = el ? Number(attrsOf(el)['@_w:w']) : NaN;
  return Number.isFinite(w) && w > 0 ? pxFromTwips(w) : undefined;
};

/** Nesting depth beyond which a table inside a table is flattened to its text (a guard against hostile files). */
const MAX_TABLE_DEPTH = 20;

function parseTable(tblNode: XmlNode, outerCtx: PartContext, numbering: NumberingInfo, warnings: string[], depth = 0): DocBlock {
  // The conditional text formatting of an enclosing table's style does not reach a table nested in one of its cells.
  const ctx = outerCtx.tableText ? { ...outerCtx, tableText: undefined } : outerCtx;
  const tblKids = childrenOf(tblNode);
  const tblPr = find(tblKids, 'w:tblPr');
  const tblPrKids = tblPr ? childrenOf(tblPr) : [];
  const styleEl = find(tblPrKids, 'w:tblStyle');
  const styleId = styleEl ? attrsOf(styleEl)['@_w:val'] : undefined;
  const style = ctx.styles.tableStyle(styleId);
  const namedStyleId = ctx.styles.namedTableStyle(styleId);
  const look = parseLook(find(tblPrKids, 'w:tblLook'));
  // A table with no border line anywhere has `borders: {}` (the model's way of saying "no lines"), unlike an
  // editor-made table, which has none set and shows the default grid.
  // A side that is not listed has no line, so sides set to none are not kept (an exported "no lines" reads back as `{}`).
  const tableBorders = Object.fromEntries(
    Object.entries({ ...style?.borders, ...parseBorders(find(tblPrKids, 'w:tblBorders'), ctx.theme) }).filter(([, border]) => border.style !== 'none')
  ) as DocBorders;

  const grid = findAll(childrenOf(find(tblKids, 'w:tblGrid') ?? {}), 'w:gridCol').map((g) => gridWidthPx(g) ?? 0);
  const trNodes = flattenBlocks(tblKids).filter((n) => tagOf(n) === 'w:tr');
  const spanOf = (tc: XmlNode): number => {
    const pr = find(childrenOf(tc), 'w:tcPr');
    const g = pr ? find(childrenOf(pr), 'w:gridSpan') : undefined;
    const n = g ? Number(attrsOf(g)['@_w:val']) : 1;
    return Number.isFinite(n) && n > 0 ? n : 1;
  };
  const gridBeforeOf = (tr: XmlNode): number => {
    const pr = find(childrenOf(tr), 'w:trPr');
    const g = pr ? find(childrenOf(pr), 'w:gridBefore') : undefined;
    const n = g ? Number(attrsOf(g)['@_w:val']) : 0;
    return Number.isFinite(n) && n > 0 ? n : 0;
  };
  const cols = grid.length > 0
    ? grid.length
    : Math.max(0, ...trNodes.map((tr) => gridBeforeOf(tr) + flattenBlocks(childrenOf(tr)).filter((n) => tagOf(n) === 'w:tc').reduce((sum, tc) => sum + spanOf(tc), 0)));

  const rows: DocTableRow[] = [];
  // Column index -> the DocTableCell object currently accumulating a
  // vertical merge, so a later row's continuation cell can bump its rowspan
  // instead of emitting a new cell.
  const openVMerge = new Map<number, { cell: DocTableRow['cells'][number] }>();

  for (const trChild of flattenBlocks(tblKids)) {
    if (tagOf(trChild) !== 'w:tr') {
      if (tagOf(trChild) !== 'w:tblPr' && tagOf(trChild) !== 'w:tblGrid') ctx.comments?.mark(trChild);
      continue;
    }
    const trNode = trChild;
    const rowIndex = trNodes.indexOf(trNode);
    const trPr = find(childrenOf(trNode), 'w:trPr');
    const trPrKids = trPr ? childrenOf(trPr) : [];
    const rowCells: DocTableRow['cells'] = [];
    const gridBefore = gridBeforeOf(trNode);
    let col = gridBefore;
    // Grid columns this row continues a vertical merge in, or opens one in.
    const mergedHere = new Set<number>();
    const extended = new Set<DocTableRow['cells'][number]>();
    // A row that starts after skipped grid columns (`w:gridBefore`) keeps its cells where they are with an empty, line-less cell in front.
    if (gridBefore > 0) {
      const noLine = { style: 'none' } as const;
      const widths = grid.length >= gridBefore ? grid.slice(0, gridBefore) : [];
      rowCells.push({
        content: [],
        ...(gridBefore > 1 && { colspan: gridBefore }),
        borders: { top: noLine, right: noLine, bottom: noLine, left: noLine },
        ...(widths.length > 0 && widths.every((w) => w > 0) && { colwidthPx: widths }),
      });
    }
    for (const tcNode of flattenBlocks(childrenOf(trNode))) {
      if (tagOf(tcNode) !== 'w:tc') {
        if (tagOf(tcNode) !== 'w:trPr' && tagOf(tcNode) !== 'w:tblPrEx') ctx.comments?.mark(tcNode);
        continue;
      }
      const tcKids = childrenOf(tcNode);
      const tcPr = find(tcKids, 'w:tcPr');
      const tcPrKids = tcPr ? childrenOf(tcPr) : [];
      const colspan = spanOf(tcNode) > 1 ? spanOf(tcNode) : undefined;
      const vMergeEl = find(tcPrKids, 'w:vMerge');
      const vMergeVal = vMergeEl ? (attrsOf(vMergeEl)['@_w:val'] ?? 'continue') : undefined;

      const open = vMergeVal === 'continue' ? openVMerge.get(col) : undefined;
      if (open) {
        // Several continuation cells under one wider merged cell extend it once.
        if (!extended.has(open.cell)) open.cell.rowspan = (open.cell.rowspan ?? 1) + 1;
        extended.add(open.cell);
        for (let i = 0; i < (colspan ?? 1); i++) mergedHere.add(col + i);
        col += colspan ?? 1;
        ctx.comments?.mark(tcNode);
        continue; // this cell contributes no content of its own
      }
      // A continuation with no merge open above it is an ordinary cell, so that its content is not lost.

      // The parts of the table style that reach this cell (header row, banding...), lowest precedence first.
      const parts = style ? applicableParts(style, look, { row: rowIndex, rows: trNodes.length, col, colspan: colspan ?? 1, cols }) : [];
      const tableText = parts.reduce<DocRunBase>((acc, part) => (part.text ? { ...acc, ...part.text } : acc), {});
      const cellCtx = Object.keys(tableText).length > 0 ? { ...ctx, tableText } : ctx;
      const { content, blocks } = parseCellContent(tcNode, cellCtx, numbering, warnings, depth);

      const shading = parseShading(find(tcPrKids, 'w:shd'), ctx.theme) ?? [...parts].reverse().find((p) => p.backgroundColor)?.backgroundColor;
      const cellBorders = [...parts.map((p) => p.borders), parseBorders(find(tcPrKids, 'w:tcBorders'), ctx.theme)].reduce<DocBorders | undefined>(
        (acc, b) => (b ? { ...acc, ...b } : acc),
        undefined
      );
      const partAlign = [...parts].reverse().find((p) => p.verticalAlign)?.verticalAlign;
      const verticalAlign = parseVAlign(find(tcPrKids, 'w:vAlign')) ?? partAlign;
      const span = colspan ?? 1;
      const tcW = find(tcPrKids, 'w:tcW');
      const widthTwips = tcW && attrsOf(tcW)['@_w:type'] !== 'pct' && attrsOf(tcW)['@_w:type'] !== 'auto' ? Number(attrsOf(tcW)['@_w:w']) : NaN;
      const fromGrid = grid.length >= col + span ? grid.slice(col, col + span) : [];
      const colwidthPx =
        fromGrid.length > 0 && fromGrid.every((w) => w > 0)
          ? fromGrid
          : Number.isFinite(widthTwips) && widthTwips > 0
            ? Array.from({ length: span }, () => pxFromTwips(widthTwips / span))
            : undefined;

      const textDirection = parseTextDirection(find(tcPrKids, 'w:textDirection'));
      const cell: DocTableRow['cells'][number] = {
        content,
        ...(blocks && { blocks }),
        ...(colspan && colspan > 1 ? { colspan } : {}),
        ...(shading && { backgroundColor: shading }),
        ...(cellBorders && { borders: cellBorders }),
        ...(verticalAlign && (verticalAlign !== 'top' || partAlign) && { verticalAlign }),
        ...(textDirection && { textDirection }),
        ...(colwidthPx && { colwidthPx }),
      };
      rowCells.push(cell);
      for (let i = 0; i < span; i++) {
        if (vMergeVal === 'restart') {
          openVMerge.set(col + i, { cell });
          mergedHere.add(col + i);
        } else {
          openVMerge.delete(col + i);
        }
      }
      col += colspan ?? 1;
    }
    // A merge that this row neither continues nor opens ends here: a later continuation cell is not part of it.
    for (const openCol of [...openVMerge.keys()]) if (!mergedHere.has(openCol)) openVMerge.delete(openCol);
    const repeatHeader = isOnEl(find(trPrKids, 'w:tblHeader'));
    const cantSplit = isOnEl(find(trPrKids, 'w:cantSplit'));
    rows.push({ cells: rowCells, ...(repeatHeader && { repeatHeader }), ...(cantSplit && { cantSplit }) });
  }

  const colWidthsPx = grid.length > 0 && grid.every((w) => w > 0) ? grid : undefined;
  return {
    type: 'table',
    rows,
    ...(colWidthsPx && { colWidthsPx }),
    borders: tableBorders,
    ...(look && { look }),
    ...(namedStyleId && { namedStyleId }),
  };
}

/** `w:textDirection` values that turn the text of a cell: bottom to top, and top to bottom (right to left). */
function parseTextDirection(el: XmlNode | undefined): DocCellTextDirection | undefined {
  const v = el ? attrsOf(el)['@_w:val'] : undefined;
  if (v === 'btLr') return 'btLr';
  if (v === 'tbRl' || v === 'tbRlV') return 'tbRl';
  return undefined;
}

/**
 * A table of contents (a `TOC` field) as a block: its entries are the paragraphs the field spans, read without the field's
 * own markup. A paragraph that only held the end of the field is not an entry, and a link around the tab and the page number
 * is split at the tab, so the tab is plain text the editor can draw as a leader.
 */
function parseTocBlock(
  range: TocRange,
  paragraphs: XmlNode[],
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[]
): DocBlock {
  const entries: DocBlock[] = [];
  paragraphs.forEach((paragraph) => {
    const stripped = withoutMarkers(paragraph, range.markers);
    if (stripped !== paragraph && isEmptyParagraph(stripped)) return;
    const parsed = parseParagraph(stripped, numbering, ctx, warnings, true);
    if (parsed.type === 'paragraph' || parsed.type === 'heading' || parsed.type === 'listItem') {
      // An entry is a paragraph, whatever its style says (a numbered or heading style does not make it a list item or a heading).
      const { type: _type, level: _level, listType: _listType, start: _start, content, ...format } = parsed as unknown as Record<string, unknown> & { content: DocInline[] };
      entries.push({ ...format, type: 'paragraph', content: splitTabsOutOfLinks(content) } as DocBlock);
    }
  });
  return { type: 'toc', instruction: range.instruction, blocks: entries };
}

/** A link whose text holds a tab becomes the link, the tab as plain text, and the link again. */
function splitTabsOutOfLinks(content: DocInline[]): DocInline[] {
  return content.flatMap((item): DocInline[] => {
    if (item.type !== 'link' || !item.text.includes('\t')) return [item];
    const { text, ...props } = item;
    const out: DocInline[] = [];
    text.split('\t').forEach((part, i) => {
      if (i > 0) out.push({ type: 'text', text: '\t' });
      if (part) out.push({ ...props, text: part });
    });
    return out;
  });
}

/**
 * The elements of a list of blocks with each table of contents (a `TOC` field spanning paragraphs) replaced by a marker
 * naming the field and its paragraphs, so the caller can read it as one block.
 */
type TocItem = { kind: 'node'; node: XmlNode } | { kind: 'toc'; range: TocRange; paragraphs: XmlNode[] };
function groupTocFields(nodes: XmlNode[]): TocItem[] {
  const ranges = findTocRanges(nodes);
  const items: TocItem[] = [];
  let i = 0;
  while (i < nodes.length) {
    const range = ranges.find((r) => r.start === i);
    if (range) {
      items.push({ kind: 'toc', range, paragraphs: nodes.slice(range.start, range.end + 1) });
      i = range.end + 1;
    } else {
      items.push({ kind: 'node', node: nodes[i] });
      i++;
    }
  }
  return items;
}

/** The paragraphs, headings and list items among blocks, looking inside content controls. */
function textBlocks(blocks: DocBlock[]): DocBlock[] {
  return blocks.flatMap((b) => (b.type === 'contentControl' ? textBlocks(b.blocks) : b.type === 'paragraph' || b.type === 'heading' || b.type === 'listItem' ? [b] : []));
}

/**
 * A block-level content control: its blocks (paragraphs, tables and nested controls) wrapped with its properties.
 * `plain` and `tableDepth` are those of the surrounding context (a table cell keeps its paragraphs as paragraphs);
 * `pending` holds bookmarks that started before the control and go to its first paragraph.
 */
function parseControlBlock(
  sdt: XmlNode,
  numbering: NumberingInfo,
  ctx: PartContext,
  warnings: string[],
  plain: boolean,
  tableDepth: number,
  pending: DocInline[]
): DocBlock {
  const parsed = parseSdt(sdt);
  const control = 'control' in parsed ? parsed.control : { type: 'richText' as const };
  const blocks: DocBlock[] = [];
  const isText = (b: DocBlock) => b.type === 'paragraph' || b.type === 'heading' || b.type === 'listItem';
  const place = (added: DocBlock[]) => {
    if (pending.length > 0) {
      const first = added.find(isText);
      if (first && 'content' in first) first.content.unshift(...pending.splice(0));
    }
    blocks.push(...added);
  };
  for (const item of groupTocFields(flattenBlocks(sdtContent(sdt), true))) {
    if (item.kind === 'toc') {
      place([parseTocBlock(item.range, item.paragraphs, numbering, ctx, warnings)]);
      continue;
    }
    const child = item.node;
    const tag = tagOf(child);
    if (tag === 'w:p') {
      place(plain ? [parseParagraph(child, numbering, ctx, warnings, true)] : parseParagraphBlocks(child, numbering, ctx, warnings));
    } else if (tag === 'w:tbl') {
      if (tableDepth >= MAX_TABLE_DEPTH) {
        warnings.push('A table nested too deeply was not imported.');
        ctx.comments?.mark(child);
      } else {
        place([parseTable(child, ctx, numbering, warnings, tableDepth)]);
      }
    } else if (tag === 'w:sdt') {
      place([parseControlBlock(child, numbering, ctx, warnings, plain, tableDepth, pending)]);
    } else if (tag === 'w:bookmarkStart') {
      const name = bookmarkName(child);
      if (name) pending.push({ type: 'bookmark', name });
    } else if (tag !== 'w:sdtPr' && tag !== 'w:sdtEndPr') {
      ctx.comments?.mark(child);
    }
  }
  // A bookmark left in `pending` (nothing after it in this control) goes to the next paragraph outside it.
  settleListStarts(blocks);
  return { type: 'contentControl', control, blocks };
}

/**
 * The content of a table cell: its paragraphs, lists and nested tables. `content` is the inline content of the first
 * paragraph that has any (what a cell of one paragraph holds); `blocks` is the whole cell, set unless the cell is one
 * plain paragraph.
 */
function parseCellContent(
  tcNode: XmlNode,
  ctx: PartContext,
  numbering: NumberingInfo,
  warnings: string[],
  depth: number
): { content: DocInline[]; blocks?: DocBlock[] } {
  const blocks: DocBlock[] = [];
  // A bookmark that starts between the cell's paragraphs is kept at the start of the next one.
  const pendingBookmarks: DocInline[] = [];
  const addParagraph = (block: DocBlock) => {
    if (pendingBookmarks.length > 0 && 'content' in block) block.content.unshift(...pendingBookmarks.splice(0));
    blocks.push(block);
  };
  for (const child of flattenBlocks(childrenOf(tcNode), !!ctx.controls)) {
    const tag = tagOf(child);
    if (tag === 'w:p') {
      addParagraph(parseParagraph(child, numbering, ctx, warnings, true));
    } else if (tag === 'w:sdt') {
      blocks.push(parseControlBlock(child, numbering, ctx, warnings, true, depth + 1, pendingBookmarks));
    } else if (tag === 'w:tbl') {
      if (depth + 1 >= MAX_TABLE_DEPTH) {
        warnings.push('A table nested too deeply was not imported.');
        ctx.comments?.mark(child);
      } else {
        blocks.push(parseTable(child, ctx, numbering, warnings, depth + 1));
      }
    } else if (tag === 'w:bookmarkStart') {
      const name = bookmarkName(child);
      if (name) pendingBookmarks.push({ type: 'bookmark', name });
    } else if (tag !== 'w:tcPr') {
      // Content that is not imported can still open or close a comment range.
      ctx.comments?.mark(child);
    }
  }
  // A bookmark with no paragraph after it goes to the end of the last paragraph, or to a paragraph of its own.
  if (pendingBookmarks.length > 0) {
    const last = blocks[blocks.length - 1];
    if (last && 'content' in last) last.content.push(...pendingBookmarks);
    else blocks.push({ type: 'paragraph', content: pendingBookmarks });
  }
  settleListStarts(blocks);

  const inline = (block: DocBlock | undefined): DocInline[] => (block && 'content' in block ? block.content : []);
  const paragraphs = textBlocks(blocks);
  const firstIndex = paragraphs.findIndex((b) => inline(b).some((c) => c.type !== 'bookmark'));
  // Bookmarks of the paragraphs before the first one with content stay in front of it
  const content = firstIndex < 0 ? inline(paragraphs[0]) : [...paragraphs.slice(0, firstIndex).flatMap(inline), ...inline(paragraphs[firstIndex])];
  const plain = blocks.length === 1 && blocks[0].type === 'paragraph';
  return { content, ...(!plain && blocks.length > 0 && { blocks }) };
}

// ---------------------------------------------------------------------------
// Section properties (page size/margins/orientation, header/footer refs)
// ---------------------------------------------------------------------------

type HeaderFooterVariant = 'default' | 'first' | 'even';

interface ParsedSectPr {
  pageWidthPx: number;
  pageHeightPx: number;
  marginTopPx: number;
  marginRightPx: number;
  marginBottomPx: number;
  marginLeftPx: number;
  orientation: 'portrait' | 'landscape';
  /** Relationship id of the header/footer of each variant (`w:type`; a reference with no type is the default one). */
  headerRIds: Partial<Record<HeaderFooterVariant, string>>;
  footerRIds: Partial<Record<HeaderFooterVariant, string>>;
  /** The first page has its own header/footer (`w:titlePg`). */
  titlePg: boolean;
  columns?: DocColumns;
  pageNumbering?: DocPageNumbering;
  verticalAlign?: DocSectionVerticalAlign;
  breakType?: DocSectionBreakType;
}

const TWIPS_TO_PX = 96 / 1440;

function parseSectPr(sectPrNode: XmlNode): ParsedSectPr {
  const kids = childrenOf(sectPrNode);
  const pgSz = find(kids, 'w:pgSz');
  const pgSzAttrs = pgSz ? attrsOf(pgSz) : {};
  const rawWidthTwips = Number(pgSzAttrs['@_w:w'] ?? 12240);
  const rawHeightTwips = Number(pgSzAttrs['@_w:h'] ?? 15840);
  // Some producers write landscape pages with w < h (the natural size) and
  // rely on w:orient; others omit w:orient and only swap w/h. Treat either
  // signal as landscape and report the on-page size (width >= height).
  const landscape = pgSzAttrs['@_w:orient'] === 'landscape' || rawWidthTwips > rawHeightTwips;
  const orientation: 'portrait' | 'landscape' = landscape ? 'landscape' : 'portrait';
  const widthTwips = landscape ? Math.max(rawWidthTwips, rawHeightTwips) : rawWidthTwips;
  const heightTwips = landscape ? Math.min(rawWidthTwips, rawHeightTwips) : rawHeightTwips;

  const pgMar = find(kids, 'w:pgMar');
  const marAttrs = pgMar ? attrsOf(pgMar) : {};
  const top = Number(marAttrs['@_w:top'] ?? 1440);
  const right = Number(marAttrs['@_w:right'] ?? 1440);
  const bottom = Number(marAttrs['@_w:bottom'] ?? 1440);
  const left = Number(marAttrs['@_w:left'] ?? 1440);

  const references = (tag: 'w:headerReference' | 'w:footerReference'): Partial<Record<HeaderFooterVariant, string>> => {
    const out: Partial<Record<HeaderFooterVariant, string>> = {};
    for (const el of kids.filter((k) => tagOf(k) === tag)) {
      const type = attrsOf(el)['@_w:type'];
      const variant: HeaderFooterVariant = type === 'first' || type === 'even' ? type : 'default';
      const rId = attrsOf(el)['@_r:id'];
      if (rId && !out[variant]) out[variant] = rId;
    }
    return out;
  };

  const titlePg = find(kids, 'w:titlePg');
  const columns = parseColumns(find(kids, 'w:cols'));
  const pageNumbering = parsePageNumbering(find(kids, 'w:pgNumType'));
  const vAlign = find(kids, 'w:vAlign');
  const vAlignVal = vAlign ? attrsOf(vAlign)['@_w:val'] : undefined;
  const typeEl = find(kids, 'w:type');
  const typeVal = typeEl ? attrsOf(typeEl)['@_w:val'] : undefined;

  return {
    pageWidthPx: Math.round(widthTwips * TWIPS_TO_PX),
    pageHeightPx: Math.round(heightTwips * TWIPS_TO_PX),
    marginTopPx: Math.round(top * TWIPS_TO_PX),
    marginRightPx: Math.round(right * TWIPS_TO_PX),
    marginBottomPx: Math.round(bottom * TWIPS_TO_PX),
    marginLeftPx: Math.round(left * TWIPS_TO_PX),
    orientation,
    headerRIds: references('w:headerReference'),
    footerRIds: references('w:footerReference'),
    titlePg: !!titlePg && isOnEl(titlePg),
    ...(columns && { columns }),
    ...(pageNumbering && { pageNumbering }),
    ...(vAlignVal === 'center' || vAlignVal === 'bottom' || vAlignVal === 'both' ? { verticalAlign: vAlignVal } : {}),
    // A new page (also what a column break amounts to with one column) is the default and is not kept.
    ...(typeVal === 'continuous' || typeVal === 'evenPage' || typeVal === 'oddPage' ? { breakType: typeVal } : {}),
  };
}

/** Text columns (`w:cols`); a single column is the default and is not kept. */
function parseColumns(cols: XmlNode | undefined): DocColumns | undefined {
  if (!cols) return undefined;
  const attrs = attrsOf(cols);
  const children = findAll(childrenOf(cols), 'w:col');
  const num = Number(attrs['@_w:num']);
  const count = Number.isInteger(num) && num > 0 ? num : Math.max(1, children.length);
  if (count < 2) return undefined;
  const space = Number(attrs['@_w:space']);
  const equal = attrs['@_w:equalWidth'];
  const widths = children.map((c) => Number(attrsOf(c)['@_w:w']));
  const unequal = (equal === '0' || equal === 'false' || equal === 'off') && widths.length === count && widths.every((w) => Number.isFinite(w) && w > 0);
  const sep = attrs['@_w:sep'];
  return {
    count,
    ...(Number.isFinite(space) && space >= 0 && attrs['@_w:space'] !== undefined && { spacingPx: pxFromTwips(space) }),
    ...((sep === '1' || sep === 'true' || sep === 'on') && { separator: true }),
    ...(unequal && { widthsPx: widths.map(pxFromTwips) }),
  };
}

/** Page numbering (`w:pgNumType`): the first page's number and the number format; chapter numbering is not kept. */
function parsePageNumbering(el: XmlNode | undefined): DocPageNumbering | undefined {
  if (!el) return undefined;
  const attrs = attrsOf(el);
  const start = Number(attrs['@_w:start']);
  const fmt = attrs['@_w:fmt'];
  const numbering: DocPageNumbering = {
    ...(attrs['@_w:start'] !== undefined && Number.isInteger(start) && { start }),
    ...(fmt && fmt !== 'decimal' && /^[A-Za-z0-9]{1,40}$/.test(fmt) && { format: fmt }),
  };
  return Object.keys(numbering).length > 0 ? numbering : undefined;
}

// ---------------------------------------------------------------------------
// Headers / footers
// ---------------------------------------------------------------------------

function parseHeaderFooter(xml: string, ctx: PartContext, warnings: string[]): DocHeaderFooter {
  const root = parseXml(xml);
  const rootEl = find(root, 'w:hdr') ?? find(root, 'w:ftr');
  const paragraphs: DocHeaderFooterParagraph[] = [];
  if (rootEl) {
    for (const p of findAll(flattenBlocks(childrenOf(rootEl)), 'w:p')) {
      const pPr = find(childrenOf(p), 'w:pPr');
      const alignment = ctx.styles.paragraph(paragraphStyleId(p), parsePPr(pPr)).alignment;
      // A header or footer shows one value on every page, so a cached field result (STYLEREF, DATE...) would be wrong
      // on most of them: only the page-number fields are kept there.
      let dropped = 0;
      let flattened = 0;
      const content = paragraphInlineContent(p, ctx, warnings)
        .filter((item) => {
          if (item.type === 'bookmark') return false; // headers and footers keep no bookmarks
          if (item.type !== 'field') return true;
          warnings.push(`Unsupported field code "${item.instruction}" — its cached value was dropped.`);
          return false;
        })
        .flatMap((item): DocInline[] => {
          // Headers and footers keep no tracked changes: deleted text is dropped, inserted text stays as ordinary text.
          if ((item.type !== 'text' && item.type !== 'link') || !item.revision) return [item];
          if (item.revision.deletion) {
            dropped++;
            return [];
          }
          flattened++;
          const { revision: _revision, ...plain } = item;
          return [plain];
        });
      if (dropped > 0) warnings.push('Tracked deleted text in a header or footer was not kept; it was dropped.');
      if (flattened > 0) warnings.push('A tracked insertion in a header or footer was not kept as a change.');
      paragraphs.push({ content, ...(alignment && alignment !== 'justify' ? { alignment } : {}) });
    }
  }
  return { paragraphs };
}

// ---------------------------------------------------------------------------
// Top-level orchestration
// ---------------------------------------------------------------------------

async function readZipText(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.file(path);
  return file ? file.async('string') : undefined;
}

export async function importDocx(bytes: Uint8Array): Promise<DocumentImportResult> {
  const warnings: string[] = [];
  const zip = await JSZip.loadAsync(bytes);

  const [documentXml, docRelsXml, stylesXml, numberingXml, commentsXml, commentsExtendedXml, peopleXml, settingsXml] = await Promise.all([
    readZipText(zip, 'word/document.xml'),
    readZipText(zip, 'word/_rels/document.xml.rels'),
    readZipText(zip, 'word/styles.xml'),
    readZipText(zip, 'word/numbering.xml'),
    readZipText(zip, 'word/comments.xml'),
    readZipText(zip, 'word/commentsExtended.xml'),
    readZipText(zip, 'word/people.xml'),
    readZipText(zip, 'word/settings.xml'),
  ]);

  if (!documentXml) {
    return { title: '', sections: [], warnings: ['word/document.xml is missing — this file may not be a valid .docx.'] };
  }

  const rels = parseRels(docRelsXml);
  const media = await buildMediaMap(zip, rels);
  // `w:evenAndOddHeaders` makes the even pages of every section show their own header and footer.
  const evenAndOddHeaders = (() => {
    try {
      const settings = settingsXml ? find(parseXml(settingsXml), 'w:settings') : undefined;
      const el = settings ? find(childrenOf(settings), 'w:evenAndOddHeaders') : undefined;
      return !!el && isOnEl(el);
    } catch {
      return false;
    }
  })();
  const parsedComments = parseComments({ comments: commentsXml, extended: commentsExtendedXml, people: peopleXml }, warnings);
  const comments = createCommentTracker(parsedComments, documentXml);
  const themeXml = await readZipText(zip, 'word/theme/theme1.xml');
  const theme = parseTheme(themeXml);
  const styles = new StyleSheet(stylesXml, theme);
  const numbering = parseNumbering(numberingXml);
  const anonymousRevisions = { next: 0 };
  const ctx: PartContext = { rels, media, styles, theme, lists: new ListCounter(numbering), comments, controls: true, anonymousRevisions };
  ctx.notes = await createNoteLookup(zip, rels, media, styles, theme, warnings, anonymousRevisions);

  const root = parseXml(documentXml);
  const documentEl = find(root, 'w:document');
  const bodyEl = documentEl ? find(childrenOf(documentEl), 'w:body') : undefined;
  const bodyChildren = bodyEl ? flattenBlocks(childrenOf(bodyEl), true) : [];

  const sections: DocSection[] = [];
  let currentBlocks: DocBlock[] = [];

  const flushSection = async (sectPrNode: XmlNode) => {
    const parsed = parseSectPr(sectPrNode);
    const header = await loadHeaderFooterVariants(zip, rels, parsed.headerRIds, parsed.titlePg, evenAndOddHeaders, styles, warnings);
    const footer = await loadHeaderFooterVariants(zip, rels, parsed.footerRIds, parsed.titlePg, evenAndOddHeaders, styles, warnings);
    sections.push({
      pageWidthPx: parsed.pageWidthPx,
      pageHeightPx: parsed.pageHeightPx,
      marginTopPx: parsed.marginTopPx,
      marginRightPx: parsed.marginRightPx,
      marginBottomPx: parsed.marginBottomPx,
      marginLeftPx: parsed.marginLeftPx,
      orientation: parsed.orientation,
      blocks: settleListStarts(currentBlocks),
      ...(header && { header }),
      ...(footer && { footer }),
      ...(parsed.columns && { columns: parsed.columns }),
      ...(parsed.pageNumbering && { pageNumbering: parsed.pageNumbering }),
      ...(parsed.verticalAlign && { verticalAlign: parsed.verticalAlign }),
      ...(parsed.breakType && { breakType: parsed.breakType }),
    });
    currentBlocks = [];
  };

  // A bookmark that starts between paragraphs is kept at the start of the next one.
  const pendingBookmarks: DocInline[] = [];
  const paragraphBlocks = (node: XmlNode): DocBlock[] => {
    const blocks = parseParagraphBlocks(node, numbering, ctx, warnings);
    if (pendingBookmarks.length > 0) {
      const first = blocks.find((b) => b.type === 'paragraph' || b.type === 'heading' || b.type === 'listItem');
      if (first && 'content' in first) first.content.unshift(...pendingBookmarks.splice(0));
    }
    return blocks;
  };

  for (const item of groupTocFields(bodyChildren)) {
    if (item.kind === 'toc') {
      currentBlocks.push(parseTocBlock(item.range, item.paragraphs, numbering, ctx, warnings));
      continue;
    }
    const node = item.node;
    const tag = tagOf(node);
    if (tag === 'w:bookmarkStart') {
      const name = bookmarkName(node);
      if (name) pendingBookmarks.push({ type: 'bookmark', name });
    } else if (tag === 'w:p') {
      const pPr = find(childrenOf(node), 'w:pPr');
      const inlineSectPr = pPr ? find(childrenOf(pPr), 'w:sectPr') : undefined;
      if (inlineSectPr) {
        // A sectPr nested in a paragraph's pPr ends that section, and per
        // OOXML the paragraph is the section's last and keeps its content
        // (Word and docx >= 9.12 emit it this way; older docx used an empty
        // carrier paragraph).
        currentBlocks.push(...paragraphBlocks(node));
        await flushSection(inlineSectPr);
        continue;
      }
      currentBlocks.push(...paragraphBlocks(node));
    } else if (tag === 'w:tbl') {
      currentBlocks.push(parseTable(node, ctx, numbering, warnings));
    } else if (tag === 'w:sdt') {
      currentBlocks.push(parseControlBlock(node, numbering, ctx, warnings, false, 0, pendingBookmarks));
    } else if (tag === 'w:sectPr') {
      // The trailing sectPr for the document's last section.
      await flushSection(node);
    } else {
      ctx.comments?.mark(node);
    }
    // Other body-level markers carry no block content — skipped.
  }

  // A document with content but no trailing sectPr (malformed, but don't
  // silently drop the content) still gets a section with sane defaults.
  if (currentBlocks.length > 0 || sections.length === 0) {
    sections.push({
      pageWidthPx: Math.round(12240 * TWIPS_TO_PX),
      pageHeightPx: Math.round(15840 * TWIPS_TO_PX),
      marginTopPx: 96,
      marginRightPx: 96,
      marginBottomPx: 96,
      marginLeftPx: 96,
      orientation: 'portrait',
      blocks: settleListStarts(currentBlocks),
    });
  }

  const coreXml = await readZipText(zip, 'docProps/core.xml');
  const title = coreXml ? parseTitle(coreXml) : '';

  const threads = comments.threads();
  const namedStyles = styles.definitions(usedStyleIds(sections));
  return {
    title,
    sections,
    warnings,
    ...(threads.length > 0 && { comments: threads }),
    ...(Object.keys(namedStyles).length > 0 && { namedStyles }),
  };
}

/** Take the custom mark, and the space after it, off the start of a note's text. */
function dropCustomMark(content: DocInline[], mark: string): void {
  const first = content[0];
  if (!first || first.type !== 'text' || !first.text.startsWith(mark)) return;
  first.text = first.text.slice(mark.length);
  if (first.text === '') content.shift();
  const next = content[0];
  if (next && next.type === 'text' && next.text.startsWith(' ')) {
    next.text = next.text.slice(1);
    if (next.text === '') content.shift();
  }
}

/** The paragraphs of a note: each one's inline content, joined by line breaks, without the leading space Word writes after the reference mark. */
function noteContent(blocks: XmlNode[], ctx: PartContext, warnings: string[]): DocInline[] {
  const content: DocInline[] = [];
  let tables = 0;
  for (const block of flattenBlocks(blocks)) {
    const tag = tagOf(block);
    if (tag === 'w:p') {
      if (content.length > 0) {
        // The line break joins the paragraphs; it takes the formatting of the text before it.
        const last = content[content.length - 1];
        if (last.type === 'text') last.text += '\n';
        else content.push({ type: 'text', text: '\n' });
      }
      content.push(...paragraphInlineContent(block, ctx, warnings));
    } else if (tag === 'w:tbl') {
      tables++;
    }
  }
  if (tables > 0) warnings.push('A table in a note was not imported.');
  // The mark is followed by a space, in a run of its own or at the start of the text.
  const first = content[0];
  if (first && first.type === 'text' && first.text.startsWith(' ')) {
    first.text = first.text.slice(1);
    if (first.text === '') content.shift();
  }
  return content;
}

/**
 * The text of the notes a body refers to, read when first asked for (a note nobody refers to is not read: the loss
 * detector reports it). A note's links and pictures are resolved through its own part's relationships, then the
 * document's.
 */
async function createNoteLookup(
  zip: JSZip,
  docRels: Map<string, string>,
  docMedia: Map<string, string>,
  styles: StyleSheet,
  theme: Theme | undefined,
  warnings: string[],
  anonymousRevisions: { next: number }
): Promise<NoteLookup> {
  const kinds: NoteKind[] = ['footnote', 'endnote'];
  const parts = new Map<NoteKind, { notes: Map<string, XmlNode[]>; ctx: PartContext }>();
  for (const kind of kinds) {
    const file = `word/${kind}s.xml`;
    const notes = parseNotes(await readZipText(zip, file), kind);
    if (notes.size === 0) continue;
    const rels = parseRels(await readZipText(zip, `word/_rels/${kind}s.xml.rels`));
    const media = await buildMediaMap(zip, rels);
    parts.set(kind, { notes, ctx: { rels: new Map([...docRels, ...rels]), media: new Map([...docMedia, ...media]), styles, theme, anonymousRevisions } });
  }
  const cache = new Map<string, DocInline[]>();
  const counts: Record<NoteKind, number> = { footnote: 0, endnote: 0 };
  return (kind, id) => {
    const part = parts.get(kind);
    const blocks = part?.notes.get(id);
    if (!part || !blocks) return undefined;
    const key = `${kind}:${id}`;
    let content = cache.get(key);
    if (!content) {
      content = noteContent(blocks, part.ctx, warnings);
      cache.set(key, content);
    }
    return { id: String(++counts[kind]), content };
  };
}

/** Every named style id the imported blocks and runs refer to. */
function usedStyleIds(sections: DocSection[]): Set<string> {
  const used = new Set<string>();
  const inlines = (content: DocInline[]) => {
    for (const inline of content) {
      if ((inline.type === 'text' || inline.type === 'link') && inline.namedStyleId) used.add(inline.namedStyleId);
      else if (inline.type === 'contentControl' || inline.type === 'noteRef') inlines(inline.content);
      else if (inline.type === 'field' && typeof inline.result !== 'string') inlines(inline.result);
    }
  };
  const hf = (part: DocHeaderFooter | undefined) => {
    for (const p of [...(part?.paragraphs ?? []), ...(part?.firstPageParagraphs ?? []), ...(part?.evenPageParagraphs ?? [])]) inlines(p.content);
  };
  const visit = (block: DocBlock) => {
    if (block.type === 'table') {
      if (block.namedStyleId) used.add(block.namedStyleId);
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (cell.blocks) cell.blocks.forEach(visit);
          else inlines(cell.content);
        }
      }
    } else if (block.type === 'contentControl' || block.type === 'toc') {
      block.blocks.forEach(visit);
    } else if ('content' in block) {
      inlines(block.content);
      if ('namedStyleId' in block && block.namedStyleId) used.add(block.namedStyleId);
    }
  };
  for (const section of sections) {
    section.blocks.forEach(visit);
    hf(section.header);
    hf(section.footer);
  }
  return used;
}

async function loadHeaderFooter(
  zip: JSZip,
  rels: Map<string, string>,
  rId: string,
  styles: StyleSheet,
  warnings: string[]
): Promise<DocHeaderFooter | undefined> {
  const target = rels.get(rId);
  if (!target) {
    warnings.push(`Header/footer relationship "${rId}" could not be resolved.`);
    return undefined;
  }
  const partXml = await readZipText(zip, `word/${target}`);
  if (!partXml) return undefined;
  const partRelsXml = await readZipText(zip, `word/_rels/${target}.rels`);
  const partRels = parseRels(partRelsXml);
  const partMedia = await buildMediaMap(zip, partRels);
  return parseHeaderFooter(partXml, { rels: partRels, media: partMedia, styles }, warnings);
}

/**
 * The header or footer of a section with its first-page and even-page variants. `differentFirstPage` follows the
 * section's `w:titlePg` and `differentOddEven` the document's `w:evenAndOddHeaders`; a variant that is switched on
 * without a part of its own is blank (empty paragraphs), as in Word. A variant part present while its switch is off is
 * kept (inert), so the file writes back as it was.
 */
async function loadHeaderFooterVariants(
  zip: JSZip,
  rels: Map<string, string>,
  rIds: Partial<Record<HeaderFooterVariant, string>>,
  titlePg: boolean,
  evenAndOdd: boolean,
  styles: StyleSheet,
  warnings: string[]
): Promise<DocHeaderFooter | undefined> {
  const load = (variant: HeaderFooterVariant) => (rIds[variant] ? loadHeaderFooter(zip, rels, rIds[variant]!, styles, warnings) : Promise.resolve(undefined));
  const base = await load('default');
  const first = await load('first');
  const even = await load('even');
  if (!base && !first && !even) return undefined;
  return {
    paragraphs: base?.paragraphs ?? [],
    ...(titlePg && { differentFirstPage: true }),
    ...(first && { firstPageParagraphs: first.paragraphs }),
    ...(evenAndOdd && { differentOddEven: true }),
    ...(even && { evenPageParagraphs: even.paragraphs }),
  };
}

function parseTitle(coreXml: string): string {
  const root = parseXml(coreXml);
  const coreProps = find(root, 'cp:coreProperties');
  if (!coreProps) return '';
  const titleEl = find(childrenOf(coreProps), 'dc:title');
  return titleEl ? textOf(titleEl) : '';
}
