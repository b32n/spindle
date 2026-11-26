/**
 * Turns a sheets-core `WorkbookData` into a stable JSON model for the goldens: random ids and pool keys are
 * replaced by positions (`A1`, `s1`, `f1`), the stable-id formula AST is dropped (the formula text stays),
 * binary payloads become digests, and nothing time- or machine-dependent is kept.
 */
import { createHash } from 'node:crypto';
import { isObj } from '../fidelity-core/diff';
import type { Json, JsonObject, ModelCounts } from '../fidelity-core/types';

/** `0` -> `A`, `26` -> `AA`. */
export function colName(index: number): string {
  let n = index;
  let name = '';
  do {
    name = String.fromCharCode(65 + (n % 26)) + name;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return name;
}

export const cellAddress = (row: number, col: number): string => `${colName(col)}${row + 1}`;

const sha256 = (data: Uint8Array | string): string =>
  createHash('sha256').update(data).digest('hex');

const DATA_URL = /^data:([^;,]*)(?:;[^,]*)?;base64,([\s\S]*)$/;

/**
 * Replace binary payloads (typed arrays, buffers and base64 data URLs) with their size and SHA-256, so a model
 * that carries images stays small and stable. Dates become ISO strings, maps and sets arrays. Everything else
 * is returned unchanged.
 */
export function digestBinaries(value: unknown): Json {
  if (value instanceof Uint8Array) return { bytes: value.length, sha256: sha256(value) };
  if (value instanceof ArrayBuffer) return digestBinaries(new Uint8Array(value));
  if (typeof value === 'string') {
    const m = DATA_URL.exec(value);
    if (!m) return value;
    const bytes = Buffer.from(m[2], 'base64');
    return { dataUrl: m[1], bytes: bytes.length, sha256: sha256(bytes) };
  }
  if (value instanceof Date)
    return Number.isNaN(value.getTime()) ? 'Invalid Date' : value.toISOString();
  if (value instanceof Map) return digestBinaries([...value.entries()]);
  if (value instanceof Set) return digestBinaries([...value]);
  if (Array.isArray(value)) return value.map(digestBinaries);
  if (value && typeof value === 'object') {
    const out: JsonObject = {};
    for (const [k, v] of Object.entries(value)) if (v !== undefined) out[k] = digestBinaries(v);
    return out;
  }
  if (typeof value === 'number' && !Number.isFinite(value)) return `#number:${String(value)}`;
  return (value ?? null) as Json;
}

/** Keys (anywhere inside a comment thread) whose values change from run to run. */
const VOLATILE_THREAD_KEYS = /^(id|createdAt|updatedAt|resolvedAt|timestamp)$|Id$/;

function scrubThread(value: unknown): Json {
  if (Array.isArray(value)) return value.map(scrubThread);
  if (value && typeof value === 'object') {
    const out: JsonObject = {};
    for (const [k, v] of Object.entries(value))
      if (v !== undefined && !VOLATILE_THREAD_KEYS.test(k)) out[k] = scrubThread(v);
    return out;
  }
  return digestBinaries(value);
}

type Loose = Record<string, unknown>;
const asArray = (v: unknown): unknown[] => (Array.isArray(v) ? v : []);
const asObj = (v: unknown): Loose => (v && typeof v === 'object' ? (v as Loose) : {});

/** Stable form of one workbook as returned by `getData()`. Never throws on a malformed shape. */
export function normalizeWorkbook(data: unknown): JsonObject {
  const wb = asObj(data);
  const styles = new Map<string, string>(); // pool id -> s1, s2 ... in first-use order
  const formats = new Map<string, string>();
  const stylePool = asObj(wb.stylePool);
  const formatPool = asObj(wb.formatPool);
  const outStyles: JsonObject = {};
  const outFormats: JsonObject = {};
  const refOf = (
    id: unknown,
    pool: Loose,
    ordinals: Map<string, string>,
    prefix: string,
    out: JsonObject
  ): string => {
    const key = String(id);
    let ref = ordinals.get(key);
    if (!ref) {
      ref = `${prefix}${ordinals.size + 1}`;
      ordinals.set(key, ref);
      out[ref] = key in pool ? digestBinaries(pool[key]) : { unresolved: true };
    }
    return ref;
  };

  const sheetsIn = asArray(wb.sheets).map(asObj);
  const sheets: Json[] = sheetsIn.map((sheet) => {
    const rowIndex = new Map<string, number>();
    const colIndex = new Map<string, number>();
    for (const [i, id] of asArray(sheet.rowOrder) as Array<[number, string]>) rowIndex.set(id, i);
    for (const [i, id] of asArray(sheet.colOrder) as Array<[number, string]>) colIndex.set(id, i);
    const rowOf = (id: unknown) => rowIndex.get(String(id));
    const colOf = (id: unknown) => colIndex.get(String(id));

    const placed: Array<{ row: number; col: number; cell: Loose }> = [];
    for (const entry of asArray(sheet.cells).map(asObj)) {
      const [rowId, colId] = String(entry.key).split(':');
      const row = rowOf(rowId);
      const col = colOf(colId);
      if (row === undefined || col === undefined) continue; // a cell without a position cannot be addressed
      placed.push({ row, col, cell: asObj(entry.cell) });
    }
    // Position order, so pool ordinals (s1, s2 ...) are assigned the same way on every run.
    placed.sort((a, b) => a.row - b.row || a.col - b.col);
    const cells = placed.map(({ row, col, cell }) => {
      const json: JsonObject = { value: digestBinaries(cell.value) };
      if (typeof cell.formula === 'string') json.formula = cell.formula;
      if (cell.styleId) json.style = refOf(cell.styleId, stylePool, styles, 's', outStyles);
      if (cell.formatId) json.format = refOf(cell.formatId, formatPool, formats, 'f', outFormats);
      if (cell.hyperlink !== undefined) json.hyperlink = digestBinaries(cell.hyperlink);
      return { row, col, json };
    });

    const config = asObj(sheet.config);
    const byRow = (pairs: unknown) =>
      Object.fromEntries(
        (asArray(pairs) as Array<[string, unknown]>)
          .map(([id, v]) => [rowOf(id), v] as const)
          .filter(([r]) => r !== undefined)
          .sort((a, b) => (a[0] as number) - (b[0] as number))
          .map(([r, v]) => [String((r as number) + 1), digestBinaries(v)])
      );
    const byCol = (pairs: unknown) =>
      Object.fromEntries(
        (asArray(pairs) as Array<[string, unknown]>)
          .map(([id, v]) => [colOf(id), v] as const)
          .filter(([c]) => c !== undefined)
          .sort((a, b) => (a[0] as number) - (b[0] as number))
          .map(([c, v]) => [colName(c as number), digestBinaries(v)])
      );
    const rowsList = (ids: unknown) =>
      asArray(ids)
        .map(rowOf)
        .filter((r): r is number => r !== undefined)
        .sort((a, b) => a - b)
        .map((r) => r + 1);
    const colsList = (ids: unknown) =>
      asArray(ids)
        .map(colOf)
        .filter((c): c is number => c !== undefined)
        .sort((a, b) => a - b)
        .map(colName);
    const merges = asArray(config.mergedRegions)
      .map(asObj)
      .map((r) => [rowOf(r.startRowId), colOf(r.startColId), rowOf(r.endRowId), colOf(r.endColId)])
      .filter((p): p is [number, number, number, number] => p.every((n) => n !== undefined))
      .sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2] || a[3] - b[3])
      .map(([r1, c1, r2, c2]) => `${cellAddress(r1, c1)}:${cellAddress(r2, c2)}`);

    const outConfig: JsonObject = {
      frozenRows: (config.frozenRows as number | undefined) ?? 0,
      frozenCols: (config.frozenCols as number | undefined) ?? 0,
      showGridLines: (config.showGridLines as boolean | undefined) ?? true,
      rowHeights: byRow(config.rowHeights),
      colWidths: byCol(config.colWidths),
      hiddenRows: rowsList(config.hiddenRows),
      hiddenCols: colsList(config.hiddenCols),
      merges,
      filters: byCol(config.filters),
      sortOrder: digestBinaries(config.sortOrder ?? []),
    };
    if (config.defaultRowHeight !== undefined)
      outConfig.defaultRowHeight = config.defaultRowHeight as number;
    if (config.defaultColWidth !== undefined)
      outConfig.defaultColWidth = config.defaultColWidth as number;

    const threads = asArray(sheet.threads)
      .map(asObj)
      .map((t) => {
        const anchor = asObj(t.anchor);
        const row = rowOf(anchor.rowId);
        const col = colOf(anchor.colId);
        const rest = scrubThread({ ...t, anchor: undefined });
        return {
          anchor: row === undefined || col === undefined ? 'orphaned' : cellAddress(row, col),
          ...(rest as JsonObject),
        } as Json;
      });

    return {
      name: String(sheet.name),
      rowCount: Number(sheet.rowCount),
      colCount: Number(sheet.colCount),
      config: outConfig,
      cells: Object.fromEntries(cells.map((c) => [cellAddress(c.row, c.col), c.json])),
      threads,
    };
  });

  const activeIndex = sheetsIn.findIndex((s) => s.id === wb.activeSheetId);
  return {
    activeSheet: activeIndex < 0 ? 0 : activeIndex,
    defaultRowHeight: Number(wb.defaultRowHeight),
    defaultColWidth: Number(wb.defaultColWidth),
    styles: outStyles,
    formats: outFormats,
    sheets,
  };
}

const DYNAMIC_REFS = new Set(['style', 'format']);

/** Inline the style and format pools into each cell (so a pool renumbering is not a difference) for diffing. */
export function resolvePools(model: Json): Json {
  if (!isObj(model)) return model;
  const styles = isObj(model.styles) ? model.styles : {};
  const formats = isObj(model.formats) ? model.formats : {};
  const sheets = Array.isArray(model.sheets) ? model.sheets : [];
  const { styles: _s, formats: _f, ...rest } = model;
  void _s;
  void _f;
  return {
    ...rest,
    sheets: sheets.map((sheet) => {
      if (!isObj(sheet) || !isObj(sheet.cells)) return sheet;
      const cells: JsonObject = {};
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        if (!isObj(cell)) {
          cells[addr] = cell;
          continue;
        }
        const next: JsonObject = { ...cell };
        for (const key of DYNAMIC_REFS) {
          const ref = cell[key];
          if (typeof ref === 'string')
            next[key] = (key === 'style' ? styles : formats)[ref] ?? { unresolved: true };
        }
        cells[addr] = next;
      }
      return { ...sheet, cells };
    }),
  };
}

/** Size counters of a normalised model: sheets, cells, formulas, merges, pooled styles and formats. */
export function countWorkbook(model: Json): ModelCounts {
  const counts: ModelCounts = {
    sheets: 0,
    cells: 0,
    formulas: 0,
    merges: 0,
    styles: 0,
    formats: 0,
  };
  if (!isObj(model)) return counts;
  counts.styles = isObj(model.styles) ? Object.keys(model.styles).length : 0;
  counts.formats = isObj(model.formats) ? Object.keys(model.formats).length : 0;
  const sheets = Array.isArray(model.sheets) ? model.sheets : [];
  counts.sheets = sheets.length;
  for (const sheet of sheets) {
    if (!isObj(sheet)) continue;
    const cells = isObj(sheet.cells) ? Object.values(sheet.cells) : [];
    counts.cells += cells.length;
    counts.formulas += cells.filter((c) => isObj(c) && typeof c.formula === 'string').length;
    const config = isObj(sheet.config) ? sheet.config : {};
    counts.merges += Array.isArray(config.merges) ? config.merges.length : 0;
  }
  return counts;
}

/** Folds cell addresses and row/column keys out of a diff path so `byField` stays readable. */
export function diffField(path: string): string {
  return path
    .replace(/\[\d+\]/g, '[]')
    .replace(/\.cells\.[A-Z]+\d+/g, '.cells.*')
    .replace(/\.(rowHeights|colWidths|filters)\.[A-Z0-9]+/g, '.$1.*');
}

/** Replace absolute repository paths in an error message so goldens do not depend on the checkout location. */
export function scrubMessage(message: string, roots: string[]): string {
  return roots
    .filter(Boolean)
    .sort((a, b) => b.length - a.length)
    .reduce((m, root) => m.split(root).join('<repo>'), message);
}
