// Y.Doc schema for collaborative workbooks.
//
// Layout:
//   meta:       Y.Map  — workbook-level fields (id, name, activeSheetId, defaults)
//   stylePool:  Y.Map  — styleId  → CellStyle JSON
//   formatPool: Y.Map  — formatId → CellFormat JSON
//   sheetIds:   Y.Array<string>  — sheet display order
//   sheets:     Y.Map  — sheetId → per-sheet Y.Map (see SheetYTypes below)
//
// meta.schemaVersion is the version of this layout the room was written with (missing = 1, the layout described
// here). attachCollabToWorkbook refuses a room whose version is newer than COLLAB_SCHEMA_VERSION (see
// assertSchemaVersion), so an older client never re-saves a room and drops keys it does not know.
//
// Each sheet's Y.Map contains:
//   meta:       Y.Map  — id, name, rowCount, colCount, frozenRows/Cols, showGridLines, sortOrder JSON
//   rowOrder:   Y.Array<rowId>
//   colOrder:   Y.Array<colId>
//   cells:      Y.Map<"rowId:colId", Cell JSON>
//   rowHeights: Y.Map<rowId, number>
//   colWidths:  Y.Map<colId, number>
//   hiddenRows: Y.Map<rowId, true>   — used as a set
//   hiddenCols: Y.Map<colId, true>
//   colStyles:  Y.Map<colId, LineFormat>  — default style/format ids of a column (schema 3; absent in sheets written before)
//   rowStyles:  Y.Map<rowId, LineFormat>  — the same for a row (schema 3)
//   filters:    Y.Map<colId, ColumnFilter JSON>
//
// Cells are stored as plain JSON inside the Y.Map. Concurrent writes to
// DIFFERENT cells merge cleanly via Y.Map; same-cell concurrent writes are
// last-writer-wins, which is acceptable for v1. Going finer-grained (per
// cell-field Y.Map) is a future option.

import * as Y from 'yjs';
import type {
  Cell,
  CellFormat,
  CellStyle,
  ColumnFilter,
  LineFormat,
  MergedRegion,
  Selection,
  SheetData,
  SortOrder,
  WorkbookData,
} from '../types';
import type { SheetCommentThread } from '../comments';

/**
 * The Y layout version this client reads and writes. 1 is the layout described above (rooms written before the key
 * existed have no `schemaVersion` and read as 1). Raised by any later change an older client could not read.
 *
 * 2: S4 formula interchange. A cell's JSON may carry `cachedOnly: true` (an imported formula the engine cannot evaluate:
 * `formula` is the file's text and `value` the result the file saved, never recalculated) and `arraySpan`
 * (`{ rows, cols, spill? }`, the extent of an array formula; `spill: true` for a dynamic-array one). A v1 client would recalculate such a cell to `#NAME?` or
 * re-evaluate its text without the mark, and on its next write of the cell keep or drop the keys unseen, so it must not join
 * a v2 room.
 *
 * 3: S6 styles. A sheet's Y.Map may carry `colStyles` and `rowStyles` (`colId`/`rowId` -> `{ styleId?, formatId? }`, the default
 * format of a column or row: a cell typed there starts with it), and a style in the pool may carry the fields `underlineKind`,
 * `strikethrough`, `indent`, `shrinkToFit`, `fillPattern`, `patternColor` and `border*Kind`. A v2 client would not see the column
 * and row formats, and on its next structural write drop them (it rewrites only the maps it knows), so it must not join a v3 room.
 * A sheet written before schema 3 has no such maps: readers treat them as empty, and the first client that needs them adds them
 * (`ensureLineStyleMaps`).
 */
export const COLLAB_SCHEMA_VERSION = 3;

/** Thrown when a room was written with a newer Y workbook layout than this client understands. */
export class CollabSchemaVersionError extends Error {
  readonly roomVersion: number;
  readonly supportedVersion: number;
  constructor(roomVersion: number, supportedVersion: number = COLLAB_SCHEMA_VERSION) {
    super(
      `This collaboration room uses workbook schema version ${roomVersion}, but this editor only supports up to ` +
        `version ${supportedVersion}. Update the editor to join the room.`,
    );
    this.name = 'CollabSchemaVersionError';
    this.roomVersion = roomVersion;
    this.supportedVersion = supportedVersion;
  }
}

/** The schema version a Y.Doc was written with (1 when the key is absent or not a number). */
export function readSchemaVersion(ydoc: Y.Doc): number {
  const v = ydoc.getMap<unknown>('meta').get('schemaVersion');
  return typeof v === 'number' && Number.isFinite(v) && v >= 1 ? v : 1;
}

/** Throw a CollabSchemaVersionError when the Y.Doc was written with a newer schema than this client supports. */
export function assertSchemaVersion(ydoc: Y.Doc, supported: number = COLLAB_SCHEMA_VERSION): void {
  const version = readSchemaVersion(ydoc);
  if (version > supported) throw new CollabSchemaVersionError(version, supported);
}

export interface SheetYTypes {
  root: Y.Map<unknown>;
  meta: Y.Map<unknown>;
  /**
   * rowId → numeric display index. Storing the reverse direction lets each
   * peer write to a different rowId without colliding (rowIds are random
   * per peer). When two peers concurrently shove different rowIds at the
   * same index, both entries persist and deserialization tie-breaks by
   * (index, rowId) — both peers converge to the same final order.
   */
  rowOrder: Y.Map<number>;
  colOrder: Y.Map<number>;
  cells: Y.Map<Cell>;
  rowHeights: Y.Map<number>;
  colWidths: Y.Map<number>;
  hiddenRows: Y.Map<true>;
  hiddenCols: Y.Map<true>;
  filters: Y.Map<ColumnFilter>;
  /** Default format per column / row. `undefined` in a sheet written before schema 3 (see `ensureLineStyleMaps`). */
  colStyles: Y.Map<LineFormat> | undefined;
  rowStyles: Y.Map<LineFormat> | undefined;
  /**
   * Merged cell regions, keyed by the composite stable-ID corner key
   * (see `mergedRegionKey`). Per-region keying lets concurrent merges of
   * different ranges converge — peers don't fight over a single array.
   */
  mergedRegions: Y.Map<MergedRegion>;
}

/** Composite key uniquely identifying a merged region by its four stable corners. */
/** Reject object keys that would poison a plain object's prototype when a key
 *  comes from peer-controlled collab data. */
function isUnsafeKey(k: string): boolean {
  return k === '__proto__' || k === 'constructor' || k === 'prototype';
}

export function mergedRegionKey(r: MergedRegion): string {
  return `${r.startRowId}:${r.startColId}:${r.endRowId}:${r.endColId}`;
}

export interface WorkbookYTypes {
  meta: Y.Map<unknown>;
  stylePool: Y.Map<CellStyle>;
  formatPool: Y.Map<CellFormat>;
  sheetIds: Y.Array<string>;
  sheets: Y.Map<Y.Map<unknown>>;
  /**
   * Comment threads, keyed `"sheetId/threadId"`. Deliberately a flat
   * top-level map rather than a child of `sheets` — the Y.UndoManager tracks
   * `sheets`, so nesting threads there would pull comment add/reply/resolve
   * into the document undo stack. Comments are intentionally not undoable.
   */
  threads: Y.Map<SheetCommentThread>;
}

/** Return the named top-level Y types. Idempotent. */
export function getWorkbookYTypes(ydoc: Y.Doc): WorkbookYTypes {
  return {
    meta: ydoc.getMap<unknown>('meta'),
    stylePool: ydoc.getMap<CellStyle>('stylePool'),
    formatPool: ydoc.getMap<CellFormat>('formatPool'),
    sheetIds: ydoc.getArray<string>('sheetIds'),
    sheets: ydoc.getMap<Y.Map<unknown>>('sheets'),
    threads: ydoc.getMap<SheetCommentThread>('threads'),
  };
}

/** Compose the flat thread key used in the top-level `threads` Y.Map. */
export function threadKey(sheetId: string, threadId: string): string {
  return `${sheetId}/${threadId}`;
}

/** Pull out the named sub-types from a sheet's Y.Map. Caller ensures the
 *  map is shaped correctly (created via createSheetYMap or hydrated). */
export function getSheetYTypes(sheetMap: Y.Map<unknown>): SheetYTypes {
  return {
    root: sheetMap,
    meta: sheetMap.get('meta') as Y.Map<unknown>,
    rowOrder: sheetMap.get('rowOrder') as Y.Map<number>,
    colOrder: sheetMap.get('colOrder') as Y.Map<number>,
    cells: sheetMap.get('cells') as Y.Map<Cell>,
    rowHeights: sheetMap.get('rowHeights') as Y.Map<number>,
    colWidths: sheetMap.get('colWidths') as Y.Map<number>,
    hiddenRows: sheetMap.get('hiddenRows') as Y.Map<true>,
    hiddenCols: sheetMap.get('hiddenCols') as Y.Map<true>,
    filters: sheetMap.get('filters') as Y.Map<ColumnFilter>,
    colStyles: sheetMap.get('colStyles') as Y.Map<LineFormat> | undefined,
    rowStyles: sheetMap.get('rowStyles') as Y.Map<LineFormat> | undefined,
    mergedRegions: sheetMap.get('mergedRegions') as Y.Map<MergedRegion>,
  };
}

/**
 * Get or create the per-sheet Y.Map under workbook.sheets[sheetId], wired
 * up with all its sub-Y.Maps. Must be called inside a Y transaction.
 *
 * The order matters: we integrate the parent sheetMap into workbook.sheets
 * BEFORE setting child Y.Maps on it. Y.js handles mutations on detached
 * types in most cases, but building a tree of nested Y.Maps and then
 * integrating the root can leave the child references in an inconsistent
 * state when read back via .get() — concretely, t.cells.set then throws
 * "Cannot read properties of undefined (reading 'set')". Integrating the
 * parent first sidesteps it entirely.
 */
export function ensureSheetYMap(
  parentSheets: Y.Map<Y.Map<unknown>>,
  sheetId: string,
): Y.Map<unknown> {
  let sheetMap = parentSheets.get(sheetId);
  if (sheetMap) return sheetMap;
  sheetMap = new Y.Map<unknown>();
  parentSheets.set(sheetId, sheetMap);
  sheetMap.set('meta', new Y.Map<unknown>());
  sheetMap.set('rowOrder', new Y.Map<number>());
  sheetMap.set('colOrder', new Y.Map<number>());
  sheetMap.set('cells', new Y.Map<Cell>());
  sheetMap.set('rowHeights', new Y.Map<number>());
  sheetMap.set('colWidths', new Y.Map<number>());
  sheetMap.set('hiddenRows', new Y.Map<true>());
  sheetMap.set('hiddenCols', new Y.Map<true>());
  sheetMap.set('filters', new Y.Map<ColumnFilter>());
  sheetMap.set('colStyles', new Y.Map<LineFormat>());
  sheetMap.set('rowStyles', new Y.Map<LineFormat>());
  sheetMap.set('mergedRegions', new Y.Map<MergedRegion>());
  return sheetMap;
}

/**
 * Add the `colStyles` / `rowStyles` maps to a sheet written before schema 3 and return its types. Call inside a transaction, only when
 * there is something to write (two clients adding the same missing map at once keep one of the two: the next structural write of the
 * other puts its entries back, since the maps are rebuilt from the local sheet).
 */
export function ensureLineStyleMaps(sheetMap: Y.Map<unknown>): SheetYTypes {
  if (!sheetMap.has('colStyles')) sheetMap.set('colStyles', new Y.Map<LineFormat>());
  if (!sheetMap.has('rowStyles')) sheetMap.set('rowStyles', new Y.Map<LineFormat>());
  return getSheetYTypes(sheetMap);
}

/** @deprecated kept as an alias during the transition; prefer ensureSheetYMap. */
export const createSheetYMap = (): Y.Map<unknown> => {
  // For callers that don't have the parent yet, fall back to the old shape.
  // Note: callers using this MUST integrate the returned map into a parent
  // BEFORE mutating its children, otherwise the sub-Y.Maps may be unreadable.
  const sheetMap = new Y.Map<unknown>();
  sheetMap.set('meta', new Y.Map<unknown>());
  sheetMap.set('rowOrder', new Y.Map<number>());
  sheetMap.set('colOrder', new Y.Map<number>());
  sheetMap.set('cells', new Y.Map<Cell>());
  sheetMap.set('rowHeights', new Y.Map<number>());
  sheetMap.set('colWidths', new Y.Map<number>());
  sheetMap.set('hiddenRows', new Y.Map<true>());
  sheetMap.set('hiddenCols', new Y.Map<true>());
  sheetMap.set('filters', new Y.Map<ColumnFilter>());
  sheetMap.set('colStyles', new Y.Map<LineFormat>());
  sheetMap.set('rowStyles', new Y.Map<LineFormat>());
  sheetMap.set('mergedRegions', new Y.Map<MergedRegion>());
  return sheetMap;
};

// ============================================================================
// WorkbookData → Y.Doc  (called once on attachCollab to seed peer-zero)
// ============================================================================

export function hydrateYDocFromData(ydoc: Y.Doc, data: WorkbookData): void {
  const types = getWorkbookYTypes(ydoc);

  ydoc.transact(() => {
    types.meta.set('schemaVersion', COLLAB_SCHEMA_VERSION);
    types.meta.set('id', data.id);
    types.meta.set('name', data.name);
    // NOTE: activeSheetId is deliberately NOT stored in the Y.Doc — which
    // sheet a user is viewing is per-user view state, not shared document
    // state. Syncing it would yank every peer onto whichever sheet anyone
    // last clicked. serializeYDocToData takes the local activeSheetId as
    // an argument instead.
    types.meta.set('defaultRowHeight', data.defaultRowHeight);
    types.meta.set('defaultColWidth', data.defaultColWidth);

    for (const [styleId, style] of Object.entries(data.stylePool)) {
      types.stylePool.set(styleId, style);
    }
    if (data.formatPool) {
      for (const [formatId, format] of Object.entries(data.formatPool)) {
        types.formatPool.set(formatId, format);
      }
    }

    for (const sheetData of data.sheets) {
      types.sheetIds.push([sheetData.id]);
      const sheetMap = ensureSheetYMap(types.sheets, sheetData.id);
      hydrateSheetYMap(sheetMap, sheetData);
      if (sheetData.threads) {
        for (const thread of sheetData.threads) {
          types.threads.set(threadKey(sheetData.id, thread.id), thread);
        }
      }
    }
  });
}

function hydrateSheetYMap(sheetMap: Y.Map<unknown>, data: SheetData): void {
  const t = getSheetYTypes(sheetMap);

  // meta
  t.meta.set('id', data.id);
  t.meta.set('name', data.name);
  t.meta.set('rowCount', data.rowCount);
  t.meta.set('colCount', data.colCount);
  if (data.config.frozenRows !== undefined) {
    t.meta.set('frozenRows', data.config.frozenRows);
  }
  if (data.config.frozenCols !== undefined) {
    t.meta.set('frozenCols', data.config.frozenCols);
  }
  if (data.config.showGridLines !== undefined) {
    t.meta.set('showGridLines', data.config.showGridLines);
  }
  if (data.config.defaultRowHeight !== undefined) {
    t.meta.set('defaultRowHeight', data.config.defaultRowHeight);
  }
  if (data.config.defaultColWidth !== undefined) {
    t.meta.set('defaultColWidth', data.config.defaultColWidth);
  }
  if (data.config.sortOrder) {
    t.meta.set('sortOrder', data.config.sortOrder);
  }
  if (data.config.mergedRegions) {
    for (const region of data.config.mergedRegions) {
      t.mergedRegions.set(mergedRegionKey(region), region);
    }
  }

  // rowOrder / colOrder: wire format is sparse Array<[idx, id]>; store as
  // Y.Map<rowId, idx>.
  if (data.rowOrder) {
    for (const [idx, id] of data.rowOrder) {
      t.rowOrder.set(id, idx);
    }
  }
  if (data.colOrder) {
    for (const [idx, id] of data.colOrder) {
      t.colOrder.set(id, idx);
    }
  }

  // cells — keys are already stable-form when rowOrder is present.
  for (const { key, cell } of data.cells) {
    t.cells.set(key, cell);
  }

  // ID-keyed config maps. These come through SheetData.config in
  // stable-form ([rowId, value] entries) when rowOrder is present.
  if (data.config.rowHeights) {
    for (const [key, h] of data.config.rowHeights) {
      if (typeof key === 'string') t.rowHeights.set(key, h);
    }
  }
  if (data.config.colWidths) {
    for (const [key, w] of data.config.colWidths) {
      if (typeof key === 'string') t.colWidths.set(key, w);
    }
  }
  if (data.config.hiddenRows) {
    for (const key of data.config.hiddenRows) {
      if (typeof key === 'string') t.hiddenRows.set(key, true);
    }
  }
  if (data.config.hiddenCols) {
    for (const key of data.config.hiddenCols) {
      if (typeof key === 'string') t.hiddenCols.set(key, true);
    }
  }
  if (data.config.filters) {
    for (const [key, filter] of data.config.filters) {
      if (typeof key === 'string') t.filters.set(key, filter);
    }
  }
  if (data.config.colStyles) {
    const maps = ensureLineStyleMaps(sheetMap);
    for (const [key, format] of data.config.colStyles) {
      if (typeof key === 'string') maps.colStyles?.set(key, format);
    }
  }
  if (data.config.rowStyles) {
    const maps = ensureLineStyleMaps(sheetMap);
    for (const [key, format] of data.config.rowStyles) {
      if (typeof key === 'string') maps.rowStyles?.set(key, format);
    }
  }
}

/**
 * Read a Y.Map<id, idx> back into the SheetData sparse Array<[idx, id]>
 * form. Sorting by (idx, id) ensures deterministic tie-breaks across peers
 * when two rowIds end up at the same display index (rare; happens only on
 * truly-concurrent insertions).
 */
function serializeOrderMap(map: Y.Map<number>): Array<[number, string]> {
  const entries: Array<[number, string]> = [];
  for (const [id, idx] of map.entries()) entries.push([idx, id]);
  entries.sort(([a, ia], [b, ib]) => a - b || ia.localeCompare(ib));
  return entries;
}

// ============================================================================
// Y.Doc → WorkbookData  (called on attach to seed a late joiner)
// ============================================================================

/**
 * Read the Y.Doc back into WorkbookData.
 *
 * `activeSheetId` is NOT in the Y.Doc (per-user view state) — the caller
 * passes its local value. It's resolved against the actual sheet list:
 * if the local active sheet has been deleted by a peer, the result snaps
 * to the first sheet so consumers never get a dangling reference.
 */
export function serializeYDocToData(
  ydoc: Y.Doc,
  selection?: Selection,
  activeSheetId?: string,
): WorkbookData {
  const t = getWorkbookYTypes(ydoc);

  const stylePool: Record<string, CellStyle> = {};
  for (const [k, v] of t.stylePool.entries()) {
    if (isUnsafeKey(k)) continue; // guard against __proto__/constructor from a peer
    stylePool[k] = v;
  }
  const formatPool: Record<string, CellFormat> = {};
  for (const [k, v] of t.formatPool.entries()) {
    if (isUnsafeKey(k)) continue;
    formatPool[k] = v;
  }

  const sheets: SheetData[] = [];
  for (const sheetId of t.sheetIds.toArray()) {
    const sheetMap = t.sheets.get(sheetId);
    if (!sheetMap) continue;
    sheets.push(serializeSheetYMap(sheetMap));
  }

  // Comment threads live in a flat top-level map keyed "sheetId/threadId";
  // demux them back onto each sheet.
  const threadsBySheet = new Map<string, SheetCommentThread[]>();
  for (const [key, thread] of t.threads.entries()) {
    const slash = key.indexOf('/');
    if (slash === -1) continue;
    const sid = key.slice(0, slash);
    let list = threadsBySheet.get(sid);
    if (!list) {
      list = [];
      threadsBySheet.set(sid, list);
    }
    list.push(thread);
  }
  for (const sheet of sheets) {
    const list = threadsBySheet.get(sheet.id);
    if (list && list.length > 0) sheet.threads = list;
  }

  let resolvedActive = activeSheetId ?? '';
  if (!resolvedActive || !sheets.some((s) => s.id === resolvedActive)) {
    resolvedActive = sheets[0]?.id ?? '';
  }

  return {
    id: (t.meta.get('id') as string) ?? '',
    name: (t.meta.get('name') as string) ?? '',
    activeSheetId: resolvedActive,
    defaultRowHeight: (t.meta.get('defaultRowHeight') as number) ?? 20,
    defaultColWidth: (t.meta.get('defaultColWidth') as number) ?? 100,
    stylePool,
    formatPool,
    sheets,
    selection,
  };
}

function serializeSheetYMap(sheetMap: Y.Map<unknown>): SheetData {
  const t = getSheetYTypes(sheetMap);

  const rowOrderEntries = serializeOrderMap(t.rowOrder);
  const colOrderEntries = serializeOrderMap(t.colOrder);

  const cells: Array<{ key: string; cell: Cell }> = [];
  for (const [key, cell] of t.cells.entries()) {
    cells.push({ key, cell });
  }

  const rowHeights: Array<[string, number]> = [];
  for (const [k, v] of t.rowHeights.entries()) rowHeights.push([k, v]);
  const colWidths: Array<[string, number]> = [];
  for (const [k, v] of t.colWidths.entries()) colWidths.push([k, v]);

  const hiddenRows: string[] = [];
  for (const k of t.hiddenRows.keys()) hiddenRows.push(k);
  const hiddenCols: string[] = [];
  for (const k of t.hiddenCols.keys()) hiddenCols.push(k);

  const filters: Array<[string, ColumnFilter]> = [];
  for (const [k, v] of t.filters.entries()) filters.push([k, v]);

  const colStyles: Array<[string, LineFormat]> = [];
  for (const [k, v] of t.colStyles?.entries() ?? []) colStyles.push([k, v]);
  const rowStyles: Array<[string, LineFormat]> = [];
  for (const [k, v] of t.rowStyles?.entries() ?? []) rowStyles.push([k, v]);

  // Read merged regions back as an array, sorted by composite key so every
  // peer agrees on iteration order (drives selection / first-match lookups).
  const mergedRegionEntries: Array<[string, MergedRegion]> = [];
  for (const [k, v] of t.mergedRegions.entries()) mergedRegionEntries.push([k, v]);
  mergedRegionEntries.sort(([a], [b]) => a.localeCompare(b));
  const mergedRegions = mergedRegionEntries.map(([, v]) => v);

  return {
    id: (t.meta.get('id') as string) ?? '',
    name: (t.meta.get('name') as string) ?? '',
    cells,
    rowOrder: rowOrderEntries,
    colOrder: colOrderEntries,
    config: {
      rowHeights: rowHeights.length ? rowHeights : undefined,
      colWidths: colWidths.length ? colWidths : undefined,
      hiddenRows: hiddenRows.length ? hiddenRows : undefined,
      hiddenCols: hiddenCols.length ? hiddenCols : undefined,
      frozenRows: t.meta.get('frozenRows') as number | undefined,
      frozenCols: t.meta.get('frozenCols') as number | undefined,
      showGridLines: t.meta.get('showGridLines') as boolean | undefined,
      defaultRowHeight: t.meta.get('defaultRowHeight') as number | undefined,
      defaultColWidth: t.meta.get('defaultColWidth') as number | undefined,
      sortOrder: t.meta.get('sortOrder') as SortOrder[] | undefined,
      filters: filters.length ? filters : undefined,
      colStyles: colStyles.length ? colStyles : undefined,
      rowStyles: rowStyles.length ? rowStyles : undefined,
      mergedRegions: mergedRegions.length ? mergedRegions : undefined,
    },
    rowCount: (t.meta.get('rowCount') as number) ?? 1000,
    colCount: (t.meta.get('colCount') as number) ?? 100,
  };
}
