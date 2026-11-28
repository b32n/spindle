import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from '../workbook';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  ensureLineStyleMaps,
  getSheetYTypes,
  getWorkbookYTypes,
  hydrateYDocFromData,
  readSchemaVersion,
  serializeYDocToData,
  createSheetYMap,
} from './y-schema';
import { getCollabBlockers } from './blockers';

// S6 added per-sheet colStyles / rowStyles (Y maps) and style fields, so the layout is at version 3.

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

function formatted(): WorkbookImpl {
  const wb = new WorkbookImpl('room', 'WB');
  wb.setCellValue(undefined, 0, 0, 'x');
  const sheet = wb.getSheet(undefined)!;
  const style = wb.getStylePool().getOrCreate({ bold: true, borderTop: '1px dashed #000000', borderTopKind: 'dashDot', fillPattern: 'darkGray', patternColor: '#112233', indent: 1 });
  const format = wb.getFormatPool().getOrCreate({ type: 'percentage', decimalPlaces: 1 });
  sheet.setColFormat(2, { styleId: style, formatId: format });
  sheet.setRowFormat(5, { styleId: style });
  return wb;
}

describe('column and row formats in the Y layout (schema version 3)', () => {
  afterEach(() => __resetInMemoryRooms());

  it('is at version 3, stamped on a hydrated room, and a version 2 client refuses it', () => {
    expect(COLLAB_SCHEMA_VERSION).toBe(3);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, formatted().getData());
    expect(readSchemaVersion(ydoc)).toBe(3);
    expect(() => assertSchemaVersion(ydoc, 2)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 3)).not.toThrow();
    try {
      assertSchemaVersion(ydoc, 2);
    } catch (e) {
      expect(e).toMatchObject({ roomVersion: 3, supportedVersion: 2 });
      expect((e as Error).message).toMatch(/version 3.*only supports up to version 2/);
    }
  });

  it('a workbook with column and row formats is not blocked from collaboration', () => {
    expect(getCollabBlockers(formatted().getData())).toEqual([]);
  });

  it('hydrates and serializes the formats and the style fields', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, formatted().getData());
    const copy = new WorkbookImpl('copy', 'Copy');
    copy.setData(serializeYDocToData(ydoc));
    const sheet = copy.getSheet(undefined)!;
    const col = sheet.getColFormat(2)!;
    expect(copy.getStylePool().get(col.styleId!)).toMatchObject({ borderTopKind: 'dashDot', fillPattern: 'darkGray', patternColor: '#112233', indent: 1 });
    expect(copy.getFormatPool().get(col.formatId!)).toMatchObject({ type: 'percentage' });
    expect(sheet.getRowFormat(5)?.styleId).toBe(col.styleId);
    expect(sheet.getColFormat(1)).toBeUndefined();
  });

  it('an old room (schema 2, sheets without the maps) loads, edits, and gets the maps when a format is set', async () => {
    const a = new WorkbookImpl('room', 'WB');
    a.setCellValue(undefined, 0, 0, 'hello');
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'old' });
    ha.ydoc.getMap('meta').set('schemaVersion', 2);
    const sheetMap = getWorkbookYTypes(ha.ydoc).sheets.get(a.getSheet(undefined)!.id)!;
    sheetMap.delete('colStyles');
    sheetMap.delete('rowStyles');
    expect(getSheetYTypes(sheetMap).colStyles).toBeUndefined();

    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'old' });
    expect(b.getCell(undefined, 0, 0)?.value).toBe('hello');
    b.setCellValue(undefined, 1, 0, 'edit');
    expect(a.getCell(undefined, 1, 0)?.value).toBe('edit');
    // setting a format on the old sheet adds the maps and reaches the other client
    const style = b.getStylePool().getOrCreate({ bold: true });
    b.getSheet(undefined)!.setColFormat(3, { styleId: style });
    expect(getSheetYTypes(getWorkbookYTypes(hb.ydoc).sheets.get(b.getSheet(undefined)!.id)!).colStyles?.size).toBe(1);
    expect(a.getSheet(undefined)!.getColFormat(3)?.styleId).toBe(style);
    expect(a.getStylePool().get(style)).toMatchObject({ bold: true });
    ha.detach();
    hb.detach();
  });

  it('two clients converge on column and row formats, including their removal and the style pool', async () => {
    const a = formatted();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    const sa = a.getSheet(undefined)!;
    const sb = b.getSheet(undefined)!;
    expect(sb.getColFormat(2)).toEqual(sa.getColFormat(2));
    expect(b.getStylePool().get(sb.getColFormat(2)!.styleId!)).toMatchObject({ bold: true, fillPattern: 'darkGray' });

    const other = b.getStylePool().getOrCreate({ italic: true });
    b.getSheet(undefined)!.setRowFormat(7, { styleId: other });
    a.getSheet(undefined)!.setColFormat(2, undefined);
    // (a remote change reloads the sheets, so they are fetched again)
    for (const [name, wb] of [['a', a], ['b', b]] as const) {
      const s = wb.getSheet(undefined)!;
      expect([name, s.getColFormat(2)]).toEqual([name, undefined]);
      expect([name, s.getRowFormat(7)?.styleId]).toEqual([name, other]);
      expect(s.getRowFormat(7)?.styleId).toBe(other);
      expect(s.getRowFormat(5)).toBeDefined();
    }
    expect(a.getStylePool().get(other)).toMatchObject({ italic: true });
    // a cell typed into the formatted row starts with its format on both
    b.setCellValue(undefined, 5, 1, 9);
    expect(a.getCell(undefined, 5, 1)?.styleId).toBe(b.getSheet(undefined)!.getRowFormat(5)!.styleId);
    ha.detach();
    hb.detach();
  });

  it('ensureLineStyleMaps adds only what is missing and createSheetYMap carries both maps', () => {
    const doc = new Y.Doc();
    const root = doc.getMap<Y.Map<unknown>>('sheets');
    const m = createSheetYMap();
    root.set('s', m);
    expect(m.get('colStyles')).toBeInstanceOf(Y.Map);
    const col = m.get('colStyles');
    m.delete('rowStyles');
    ensureLineStyleMaps(m);
    expect(m.get('colStyles')).toBe(col);
    expect(m.get('rowStyles')).toBeInstanceOf(Y.Map);
  });
});
