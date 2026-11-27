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
import { partsWithContentType } from './content-types';
import type { XlsxPackage } from './package';
import { relsOfKind } from './relationships';
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

const STATES: ReadonlySet<string> = new Set<XlsxSheetState>(['visible', 'hidden', 'veryHidden']);

/** `state` as the file writes it; anything else (or nothing) is visible. */
const sheetState = (v: string | undefined): XlsxSheetState => (v !== undefined && STATES.has(v) ? (v as XlsxSheetState) : 'visible');

function commentParts(rels: readonly XlsxRelationship[], pkg: XlsxPackage, sheetName: string): XlsxCommentParts {
  const found = (kind: string) =>
    relsOfKind(rels, kind).flatMap((r) => {
      if (r.external) return [];
      const part = pkg.resolve(r.target);
      if (part === undefined) pkg.warnings.push(`The sheet "${sheetName}" has ${kind} at ${r.target}, which the package does not contain.`);
      return part ?? [];
    });
  const out: XlsxCommentParts = { threadedComments: found('threadedComment') };
  put(out, 'comments', found('comments')[0]);
  put(out, 'vmlDrawing', found('vmlDrawing')[0]);
  return out;
}

async function readSheetRef(el: XmlElement, index: number, workbookRels: readonly XlsxRelationship[], pkg: XlsxPackage): Promise<XlsxSheetRef> {
  const relId = el.attrs.id ?? '';
  const rel = workbookRels.find((r) => r.id === relId);
  let part: string | undefined;
  if (rel && !rel.external) {
    part = pkg.resolve(rel.target);
    if (part === undefined) pkg.warnings.push(`The sheet "${el.attrs.name ?? ''}" points at ${rel.target}, which the package does not contain.`);
  } else pkg.warnings.push(`The sheet "${el.attrs.name ?? ''}" has no usable relationship (${relId === '' ? 'no r:id' : relId}).`);
  const rels = part === undefined ? [] : await pkg.rels(part);
  const sheetId = intAttr(el.attrs.sheetId);
  return {
    name: el.attrs.name ?? `Sheet${index + 1}`,
    sheetId: sheetId ?? index + 1,
    state: sheetState(el.attrs.state),
    relId,
    kind: rel ? (SHEET_KINDS.get(rel.kind.toLowerCase()) ?? rel.kind) : 'worksheet',
    ...(part === undefined ? {} : { part }),
    rels,
    commentParts: commentParts(rels, pkg, el.attrs.name ?? ''),
  };
}

function readDefinedNames(root: XmlElement | undefined): XlsxDefinedName[] {
  const out: XlsxDefinedName[] = [];
  for (const el of kids(kid(root, 'definedNames'), 'definedName')) {
    const name = el.attrs.name;
    if (name === undefined || name === '') continue;
    const dn: XlsxDefinedName = { name, ref: el.text.trim().replace(/^=/, ''), builtin: name.startsWith('_xlnm.') };
    put(dn, 'localSheetId', intAttr(el.attrs.localSheetId));
    put(dn, 'hidden', boolAttr(el.attrs.hidden));
    put(dn, 'comment', el.attrs.comment);
    out.push(dn);
  }
  return out;
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
export async function readWorkbook(pkg: XlsxPackage): Promise<XlsxWorkbookInfo> {
  const part = await findWorkbookPart(pkg);
  const root = await pkg.xml(part);
  const rels = await pkg.rels(part);
  const workbookPr = kid(root, 'workbookPr');
  const view = kid(kid(root, 'bookViews'), 'workbookView');
  const sheets: XlsxSheetRef[] = [];
  for (const [i, el] of kids(kid(root, 'sheets'), 'sheet').entries()) sheets.push(await readSheetRef(el, i, rels, pkg));

  const info: XlsxWorkbookInfo = {
    part,
    date1904: isTrue(workbookPr?.attrs.date1904),
    calc: readCalc(kid(root, 'calcPr')),
    sheets,
    definedNames: readDefinedNames(root),
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
