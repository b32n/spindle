import type {
  XlsxCalcProperties,
  XlsxCommentParts,
  XlsxDefinedName,
  XlsxRelationship,
  XlsxSheetRef,
  XlsxSheetState,
  XlsxWorkbookInfo,
} from '@b32nio/spindle-shared/export/xlsx';
import { ImportError } from '../../common/import-error';
import { isTrue, kid, kids, type XmlElement } from '../xml';
import { abbreviate, CappedWarnings, flat, truncate } from './capped';
import { partsWithContentType } from './content-types';
import type { XlsxPackage } from './package';
import { relsOfKind } from './relationships';
import { scanXml } from './sax';
import { TreeBuilder, type ParsedTree } from './tree';
import { boolAttr, intAttr, numAttr, put } from './values';

// ---------------------------------------------------------------------------
// Locating the workbook and everything it points at: the workbook part (from the package relationships, else from the
// content types, else `xl/workbook.xml`), then `sheets`, `definedNames`, `calcPr`, `workbookPr` (date1904), `bookViews`, and the
// parts for strings, styles, theme and comments. Names and ids are as the file has them; deciding what a sheet is called
// when the file repeats a name is for the consumer.
// ---------------------------------------------------------------------------

/** A workbook content type: the plain, template and macro-enabled forms (`.xlsx`, `.xltx`, `.xlsm`, `.xltm`). */
const WORKBOOK_TYPE = /(?:spreadsheetml\.(?:sheet|template)|ms-excel\.(?:sheet|template)\.macroEnabled)\.main\+xml$/i;

/** Locate the workbook part, or throw the `ImportError` for a package that has none. */
export async function findWorkbookPart(pkg: XlsxPackage): Promise<string> {
  for (const rel of relsOfKind(await pkg.rels(''), 'officeDocument')) {
    const exact = rel.external ? undefined : pkg.resolve(rel.target);
    if (exact !== undefined) return exact;
  }
  for (const name of partsWithContentType(pkg.contentTypes, (t) => WORKBOOK_TYPE.test(t))) {
    const exact = pkg.resolve(name);
    if (exact !== undefined) return exact;
  }
  const conventional = pkg.resolve('xl/workbook.xml');
  if (conventional !== undefined) return conventional;
  throw new ImportError('corrupt', 'The file has no workbook part (it is not a complete Excel workbook).');
}

/**
 * Whether `part` (as `findWorkbookPart` returned it) is a binary workbook (`.xlsb`: `xl/workbook.bin`), which is not XML and is not
 * read. `findWorkbookPart` still returns it, so the caller can refuse with a message of its own (`importXlsx` does) rather than see an empty workbook.
 */
export function isBinaryWorkbook(pkg: XlsxPackage, part: string): boolean {
  return /\.bin$/i.test(part) || /ms-excel\.sheet\.binary\.macroEnabled\.main$/i.test(pkg.contentType(part) ?? '');
}

/** Sheet kinds by lower-cased relationship kind; a `Map` because the kind is the file's text (`constructor` is not a kind). */
const SHEET_KINDS: ReadonlyMap<string, string> = new Map([['worksheet', 'worksheet'], ['chartsheet', 'chartsheet'], ['dialogsheet', 'dialogsheet'], ['xlmacrosheet', 'macrosheet'], ['xlintlmacrosheet', 'macrosheet']]);

/** The workbook's `<sheet>` entries read are this many times `maxSheets`: entries that are not read as worksheets (chart sheets, a part repeated, no part) count, so naming many cannot make the reader do work the sheet limit does not bound. */
export const SHEET_ENTRY_FACTOR = 4;

/**
 * Longest sheet name kept (characters). Excel allows 31, and Excel-written names never exceed it; other producers are tolerated up to
 * this many, which is more than any real sheet name needs, and the rest of a longer name is cut with a warning, so a name of
 * megabytes is never carried through the import (into the result, the editor's tabs, or each message that names the sheet).
 */
export const MAX_SHEET_NAME_CHARS = 255;

const STATES: ReadonlySet<string> = new Set<XlsxSheetState>(['visible', 'hidden', 'veryHidden']);

/** `state` as the file writes it; anything else (or nothing) is visible. */
const sheetState = (v: string | undefined): XlsxSheetState => (v !== undefined && STATES.has(v) ? (v as XlsxSheetState) : 'visible');

/** What the sheet entries of a workbook share: the workbook's relationships by id, the comment parts found per sheet part (a part named by many entries is looked at once), and the capped warnings. */
interface SheetScan {
  byId: ReadonlyMap<string, XlsxRelationship>;
  comments: Map<string, XlsxCommentParts>;
  problems: CappedWarnings;
}

function commentParts(rels: readonly XlsxRelationship[], pkg: XlsxPackage, sheetName: string, problems: CappedWarnings): XlsxCommentParts {
  const found = (kind: string) =>
    relsOfKind(rels, kind).flatMap((r) => {
      if (r.external) return [];
      const part = pkg.resolve(r.target);
      if (part === undefined) problems.add(`The sheet "${abbreviate(sheetName)}" has ${kind} at ${abbreviate(r.target)}, which the package does not contain.`);
      return part ?? [];
    });
  const out: XlsxCommentParts = { threadedComments: found('threadedComment') };
  put(out, 'comments', found('comments')[0]);
  put(out, 'vmlDrawing', found('vmlDrawing')[0]);
  return out;
}

async function readSheetRef(el: XmlElement, index: number, scan: SheetScan, pkg: XlsxPackage): Promise<XlsxSheetRef> {
  let name = el.attrs.name ?? `Sheet${index + 1}`;
  if (name.length > MAX_SHEET_NAME_CHARS) {
    scan.problems.add(`The sheet name "${abbreviate(name)}" is ${name.length} characters long (Excel allows 31); it was cut to ${MAX_SHEET_NAME_CHARS}.`);
    name = truncate(name, MAX_SHEET_NAME_CHARS);
    pkg.loss.add('sheetNames', name);
  }
  const relId = el.attrs.id ?? '';
  const rel = scan.byId.get(relId);
  let part: string | undefined;
  if (rel && !rel.external) {
    part = pkg.resolve(rel.target);
    if (part === undefined) scan.problems.add(`The sheet "${abbreviate(name)}" points at ${abbreviate(rel.target)}, which the package does not contain.`);
  } else scan.problems.add(`The sheet "${abbreviate(name)}" has no usable relationship (${relId === '' ? 'no r:id' : abbreviate(relId)}).`);
  const rels = part === undefined ? [] : await pkg.rels(part);
  let comments = part === undefined ? undefined : scan.comments.get(part);
  if (comments === undefined) {
    comments = commentParts(rels, pkg, name, scan.problems);
    if (part !== undefined) scan.comments.set(part, comments);
  }
  const sheetId = intAttr(el.attrs.sheetId);
  return {
    name,
    sheetId: sheetId ?? index + 1,
    state: sheetState(el.attrs.state),
    relId,
    kind: rel ? (SHEET_KINDS.get(rel.kind.toLowerCase()) ?? rel.kind) : 'worksheet',
    ...(part === undefined ? {} : { part }),
    rels,
    commentParts: comments,
  };
}

/** A `<definedName>` as a name: `undefined` for one without a name. */
function toDefinedName(attrs: Readonly<Record<string, string>>, text: string): XlsxDefinedName | undefined {
  const name = attrs.name;
  if (name === undefined || name === '') return undefined;
  const dn: XlsxDefinedName = { name, ref: flat(text.trim().replace(/^=/, '')), builtin: name.startsWith('_xlnm.') };
  put(dn, 'localSheetId', intAttr(attrs.localSheetId));
  put(dn, 'hidden', boolAttr(attrs.hidden));
  put(dn, 'comment', attrs.comment);
  return dn;
}

interface WorkbookParse extends ParsedTree {
  definedNames: XlsxDefinedName[];
  /** How many names past `maxNames` were in the part (the first ones were kept). */
  namesLeftOut: number;
}

/**
 * The workbook part in one pass: an element tree of everything but the first `<definedNames>` (under the element budget
 * `maxElements`), and that element's names as a list of events (at most `maxNames` kept): a workbook of a million names is
 * data, not a document to refuse, and its memory is the names kept.
 */
function parseWorkbook(text: string, maxElements: number, maxNames: number): WorkbookParse {
  const builder = new TreeBuilder({ left: maxElements });
  const definedNames: XlsxDefinedName[] = [];
  let depth = 0;
  let listSeen = false;
  let inList = false;
  let attrs: Readonly<Record<string, string>> | undefined;
  let nameText = '';
  let namesLeftOut = 0;
  const wellFormed = scanXml(text, {
    open(name, a) {
      depth++;
      if (depth === 2 && name === 'definedNames' && !listSeen) listSeen = inList = true;
      else if (inList) {
        if (depth === 3 && name === 'definedName') {
          attrs = a;
          nameText = '';
        }
      } else builder.start(name, a);
    },
    close() {
      if (!inList) builder.end();
      else if (depth === 3 && attrs !== undefined) {
        const dn = toDefinedName(attrs, nameText);
        attrs = undefined;
        if (dn !== undefined) {
          if (definedNames.length < maxNames) definedNames.push(dn);
          else namesLeftOut++;
        }
      } else if (depth === 2) inList = false;
      depth--;
    },
    text(chunk) {
      if (inList) {
        if (depth === 3 && attrs !== undefined) nameText += chunk;
      } else builder.text(chunk);
    },
    halted: () => builder.dropped,
  });
  if (builder.dropped) return { root: undefined, overLimit: true, definedNames, namesLeftOut };
  return { root: wellFormed ? builder.root : undefined, overLimit: false, definedNames, namesLeftOut };
}

const CALC_MODES = new Set(['auto', 'manual', 'autoNoTable']);

function readCalc(el: XmlElement | undefined): XlsxCalcProperties {
  const calc: XlsxCalcProperties = {};
  if (!el) return calc;
  const a = el.attrs;
  put(calc, 'calcId', intAttr(a.calcId));
  if (a.calcMode !== undefined && CALC_MODES.has(a.calcMode)) calc.calcMode = a.calcMode as XlsxCalcProperties['calcMode'];
  put(calc, 'fullCalcOnLoad', boolAttr(a.fullCalcOnLoad));
  if (a.refMode === 'A1' || a.refMode === 'R1C1') calc.refMode = a.refMode;
  put(calc, 'iterate', boolAttr(a.iterate));
  put(calc, 'iterateCount', intAttr(a.iterateCount));
  put(calc, 'iterateDelta', numAttr(a.iterateDelta));
  put(calc, 'fullPrecision', boolAttr(a.fullPrecision));
  put(calc, 'calcOnSave', boolAttr(a.calcOnSave));
  put(calc, 'concurrentCalc', boolAttr(a.concurrentCalc));
  put(calc, 'forceFullCalc', boolAttr(a.forceFullCalc));
  return calc;
}

/** The part a workbook-level relationship of `kind` names: its relationship, else a part with the content type, else the conventional name. */
function locate(pkg: XlsxPackage, rels: readonly XlsxRelationship[], kind: string, contentType: RegExp, conventional?: string): string | undefined {
  for (const rel of relsOfKind(rels, kind)) {
    const exact = rel.external ? undefined : pkg.resolve(rel.target);
    if (exact !== undefined) return exact;
  }
  for (const name of partsWithContentType(pkg.contentTypes, (t) => contentType.test(t))) {
    const exact = pkg.resolve(name);
    if (exact !== undefined) return exact;
  }
  return conventional === undefined ? undefined : pkg.resolve(conventional);
}

/**
 * Read the workbook part and locate what it refers to. Throws only `ImportError('corrupt')` when no workbook part can be found;
 * a workbook part that cannot be parsed gives an empty workbook (no sheets) and a warning.
 */
export async function readWorkbook(pkg: XlsxPackage, options: { strict?: boolean } = {}): Promise<XlsxWorkbookInfo> {
  const part = await findWorkbookPart(pkg);
  const text = await pkg.text(part);
  const parsed = text === undefined ? undefined : parseWorkbook(text, pkg.limits.maxPartElements, pkg.limits.maxPartElements);
  const root = parsed === undefined ? undefined : pkg.report(part, parsed);
  if (root === undefined && options.strict)
    throw new ImportError(
      'corrupt',
      parsed?.overLimit
        ? `The workbook part of the file has too many elements (more than the import limit maxPartElements, ${pkg.limits.maxPartElements}; raise it with the limits option).`
        : 'The workbook part of the file is damaged (it is not well-formed XML).'
    );
  if (parsed && parsed.namesLeftOut > 0) {
    pkg.loss.add('definedNames', part, parsed.namesLeftOut);
    pkg.warnings.push(`The workbook has more than ${pkg.limits.maxPartElements} defined names (the import limit maxPartElements); ${parsed.namesLeftOut} were not read.`);
  }
  const entries = kids(kid(root, 'sheets'), 'sheet');
  const limit = pkg.limits.maxSheets * SHEET_ENTRY_FACTOR;
  // the relationships kept are those of the kinds the reader uses and those the sheet entries name (their kind is the sheet's kind)
  const rels = await pkg.rels(part, new Set(entries.slice(0, limit).map((el) => el.attrs.id ?? '')));
  const workbookPr = kid(root, 'workbookPr');
  const view = kid(kid(root, 'bookViews'), 'workbookView');
  const sheets: XlsxSheetRef[] = [];
  // the workbook's relationships by id (the first of a repeated id), and each sheet part's comment parts once: a workbook that
  // names a million sheets must not scan its relationships once per sheet
  const byId = new Map<string, XlsxRelationship>();
  for (const r of rels) if (!byId.has(r.id)) byId.set(r.id, r);
  const problems = new CappedWarnings((m) => pkg.warnings.push(m), 'sheet warnings');
  const scan: SheetScan = { byId, comments: new Map(), problems };
  for (const [i, el] of entries.slice(0, limit).entries()) sheets.push(await readSheetRef(el, i, scan, pkg));
  problems.close();
  if (entries.length > limit) {
    pkg.loss.add('sheets', 'the sheet entries of the workbook', entries.length - limit);
    pkg.warnings.push(`The workbook lists ${entries.length} sheets; only the first ${limit} entries were read (${SHEET_ENTRY_FACTOR} times the import limit maxSheets).`);
  }

  const info: XlsxWorkbookInfo = {
    part,
    date1904: isTrue(workbookPr?.attrs.date1904),
    calc: readCalc(kid(root, 'calcPr')),
    sheets,
    definedNames: root === undefined ? [] : (parsed as WorkbookParse).definedNames,
    parts: {},
  };
  put(info, 'codeName', workbookPr?.attrs.codeName);
  put(info, 'appName', kid(root, 'fileVersion')?.attrs.appName);
  put(info, 'activeTab', intAttr(view?.attrs.activeTab));
  put(info, 'firstSheet', intAttr(view?.attrs.firstSheet));
  put(info.parts, 'sharedStrings', locate(pkg, rels, 'sharedStrings', /spreadsheetml\.sharedStrings\+xml$/i, 'xl/sharedStrings.xml'));
  put(info.parts, 'styles', locate(pkg, rels, 'styles', /spreadsheetml\.styles\+xml$/i, 'xl/styles.xml'));
  put(info.parts, 'theme', locate(pkg, rels, 'theme', /officedocument\.theme\+xml$/i, 'xl/theme/theme1.xml'));
  put(info.parts, 'persons', locate(pkg, rels, 'person', /ms-excel\.person\+xml$/i));
  return info;
}
