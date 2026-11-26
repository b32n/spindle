/**
 * The `values` kind: what the engine computes for every formula cell after import, against the value the file
 * itself caches for that cell (what Excel, LibreOffice or the producing library stored). This is the plan's
 * probe 3 as a golden. The expectations are read straight from the package XML, not from the importer, so they
 * are right even where the importer mangles a value, and so they exist for files whose import throws.
 */
import { relationshipTargets, type OoxmlParts } from '../fidelity-core/ooxml-package';
import type { Json, ValuesTotals } from '../fidelity-core/types';
import { isObj } from '../fidelity-core/diff';
import { cellAddress, stableNumber } from './normalize';

export type Scalar = string | number | boolean | null;

export interface ExpectedCell {
  /** Sheet name and which sheet of that name this is (names may repeat in a damaged file). */
  sheet: string;
  occurrence: number;
  address: string;
  expected: Scalar;
  /** The formula had no cached value, or contains a volatile function. */
  skip: 'uncached' | 'volatile' | null;
}

export interface FileFormulas {
  /** Sheet names in workbook order. */
  sheets: string[];
  cells: ExpectedCell[];
  /** Sheets whose part is missing, unreadable or over the size cap. */
  unreadableSheets: number;
}

const VOLATILE = /\b(?:NOW|TODAY|RAND|RANDBETWEEN|RANDARRAY)\s*\(/i;

const ENTITIES: Record<string, string> = { lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' };

/** One pass, so `&amp;#65;` stays the text `&#65;` instead of being decoded twice. */
export function unescapeXml(text: string): string {
  return text.replace(
    /&(lt|gt|quot|apos|amp);|&#x([0-9a-fA-F]+);|&#(\d+);|_x([0-9a-fA-F]{4})_/g,
    (_, name?: string, hex?: string, dec?: string, esc?: string) =>
      name !== undefined
        ? ENTITIES[name]
        : safeChar(parseInt((hex ?? esc ?? dec) as string, dec !== undefined ? 10 : 16))
  );
}

const safeChar = (cp: number): string =>
  cp >= 0 && cp <= 0x10ffff ? String.fromCodePoint(cp) : '�';

/** The text of every `<t>` below an element's XML (rich-text runs are concatenated; phonetic runs skipped). */
function textOf(xml: string): string {
  const withoutPhonetic = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
  return [...withoutPhonetic.matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)]
    .map((m) => unescapeXml(m[1] ?? ''))
    .join('');
}

export function sharedStrings(xml: string | undefined): string[] {
  if (!xml) return [];
  return [...xml.matchAll(/<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g)].map((m) => textOf(m[1] ?? ''));
}

const attr = (tag: string, name: string): string | undefined =>
  new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];

/** `A1` -> [0, 0]; null for anything that is not a plain cell address. */
export function parseAddress(ref: string): [number, number] | null {
  const m = /^([A-Z]{1,3})(\d+)$/.exec(ref);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return [Number(m[2]) - 1, col - 1];
}

function cachedValue(
  type: string | undefined,
  body: string,
  strings: string[]
): Scalar | undefined {
  if (type === 'inlineStr') return textOf(/<is\b[^>]*>([\s\S]*?)<\/is>/.exec(body)?.[1] ?? '');
  const v = /<v\b[^>]*?(?:\/>|>([\s\S]*?)<\/v>)/.exec(body);
  if (!v) return undefined;
  const raw = unescapeXml(v[1] ?? '');
  if (type === 'b') return raw === '1';
  if (type === 's') return strings[Number(raw)] ?? undefined;
  if (type === 'str' || type === 'e' || type === 'd') return raw;
  const n = Number(raw);
  return raw === '' || Number.isNaN(n) ? undefined : n;
}

/**
 * Formula cells of one worksheet part, with their cached values. The cell and formula regexes are lazy and can be
 * quadratic on a sheet with many unclosed `<c>` elements; that is bounded by the 64 MB part cap and the corpus is trusted.
 */
export function readSheetFormulas(
  xml: string,
  strings: string[],
  sheet: string,
  occurrence: number
): ExpectedCell[] {
  const out: ExpectedCell[] = [];
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const body = m[2] ?? '';
    const f = /<f\b[^>]*?(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
    const pos = parseAddress(attr(m[1], 'r') ?? '');
    if (!f || !pos) continue;
    const cached = cachedValue(attr(m[1], 't'), body, strings);
    const volatile = VOLATILE.test(unescapeXml(f[1] ?? ''));
    out.push({
      sheet,
      occurrence,
      address: cellAddress(pos[0], pos[1]),
      expected: cached ?? null,
      skip: volatile ? 'volatile' : cached === undefined ? 'uncached' : null,
    });
  }
  return out;
}

/** Every formula cell of every worksheet of a package, in workbook order. */
export function readFileFormulas(parts: Pick<OoxmlParts, 'xml'>): FileFormulas {
  const workbook = parts.xml['xl/workbook.xml'] ?? '';
  const rels = relationshipTargets(
    'xl/_rels/workbook.xml.rels',
    parts.xml['xl/_rels/workbook.xml.rels'] ?? ''
  );
  const strings = sharedStrings(parts.xml['xl/sharedStrings.xml']);
  const sheets: string[] = [];
  const seen = new Map<string, number>();
  const cells: ExpectedCell[] = [];
  let unreadableSheets = 0;
  for (const m of workbook.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unescapeXml(attr(m[0], 'name') ?? '');
    const occurrence = seen.get(name) ?? 0;
    seen.set(name, occurrence + 1);
    sheets.push(name);
    const target = rels.find((r) => r.id === attr(m[0], 'r:id'))?.target;
    const xml = target ? parts.xml[target] : undefined;
    if (xml === undefined) unreadableSheets++;
    else cells.push(...readSheetFormulas(xml, strings, name, occurrence));
  }
  return { sheets, cells, unreadableSheets };
}

/** Relative tolerance (9 significant digits) with a floor that only absorbs floating-point noise around zero. */
const TOLERANCE = 1e-9;
const ABSOLUTE_FLOOR = 1e-15;

const ERROR =
  /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|CIRCULAR!)/;

/** The error code a value shows (`#NAME?` for `#NAME? Function PI not found`), or null for anything else. */
export const errorCode = (v: Scalar): string | null =>
  typeof v === 'string' ? (ERROR.exec(v)?.[0] ?? null) : null;

/**
 * Numbers agree to 9 significant digits (cached values are written with at most 15); errors agree when they are the
 * same error code (the engine appends a diagnostic to its message); everything else must be equal.
 */
export function sameValue(expected: Scalar, actual: Scalar): boolean {
  if (typeof expected === 'number' && typeof actual === 'number')
    return (
      Math.abs(expected - actual) <=
      Math.max(ABSOLUTE_FLOOR, TOLERANCE * Math.max(Math.abs(expected), Math.abs(actual)))
    );
  const [codeA, codeB] = [errorCode(expected), errorCode(actual)];
  if (codeA !== null && codeB !== null) return codeA === codeB;
  return expected === actual;
}

function kindOf(v: Scalar): string {
  if (v === null) return 'empty';
  if (typeof v === 'string') return errorCode(v) === null ? 'text' : 'error';
  return typeof v;
}

export const show = (v: Scalar): string =>
  v === null
    ? '(empty)'
    : typeof v === 'string'
      ? JSON.stringify(v)
      : typeof v === 'boolean'
        ? v
          ? 'TRUE'
          : 'FALSE'
        : String(stableNumber(v));

export interface ValuesSnapshot extends ValuesTotals {
  /** `ok`, or why nothing was compared. */
  status: 'ok' | 'skipped-import-error' | 'unreadable-package';
  /** Formula cells in the file. */
  formulaCells: number;
  /** Formulas without a cached value, and formulas with a volatile function: not compared. */
  uncached: number;
  volatile: number;
  /** Sheets of the file whose part could not be read. */
  unreadableSheets: number;
  /** Mismatches counted by what was expected and what the engine produced (`number -> error`). */
  mismatchClasses: { [kind: string]: number };
  firstMismatches: Array<{ sheet: string; cell: string; expected: string; actual: string }>;
}

export const MAX_MISMATCHES = 20;

const emptyValues = (
  status: ValuesSnapshot['status'],
  formulaCells = 0,
  extra: Partial<ValuesSnapshot> = {}
): ValuesSnapshot => ({
  status,
  formulaCells,
  compared: 0,
  matched: 0,
  mismatched: 0,
  uncached: 0,
  volatile: 0,
  unreadableSheets: 0,
  mismatchClasses: {},
  firstMismatches: [],
  ...extra,
});

/** Compare the engine's values (read from a normalised model) with the file's cached values. */
export function compareValues(
  file: FileFormulas,
  model: Json | null,
  importFailed: boolean
): ValuesSnapshot {
  if (importFailed || model === null)
    return emptyValues('skipped-import-error', file.cells.length, {
      unreadableSheets: file.unreadableSheets,
    });
  const modelSheets = isObj(model) && Array.isArray(model.sheets) ? model.sheets : [];
  const out = emptyValues('ok', file.cells.length, { unreadableSheets: file.unreadableSheets });
  for (const cell of file.cells) {
    if (cell.skip === 'uncached') {
      out.uncached++;
      continue;
    }
    if (cell.skip === 'volatile') {
      out.volatile++;
      continue;
    }
    const sheet = modelSheets.filter((s) => isObj(s) && s.name === cell.sheet)[cell.occurrence];
    const mc = isObj(sheet) && isObj(sheet.cells) ? sheet.cells[cell.address] : undefined;
    const actual = (isObj(mc) && mc.value !== undefined ? mc.value : null) as Scalar;
    out.compared++;
    if (sameValue(cell.expected, actual)) {
      out.matched++;
      continue;
    }
    out.mismatched++;
    const cls = `${kindOf(cell.expected)} -> ${kindOf(actual)}`;
    out.mismatchClasses[cls] = (out.mismatchClasses[cls] ?? 0) + 1;
    if (out.firstMismatches.length < MAX_MISMATCHES)
      out.firstMismatches.push({
        sheet: cell.sheet,
        cell: cell.address,
        expected: show(cell.expected),
        actual: show(actual),
      });
  }
  return out;
}

/** A values snapshot for a package that cannot be read at all. */
export const unreadableValues = (): ValuesSnapshot => emptyValues('unreadable-package');
