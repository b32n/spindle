import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from '../workbook';
import { importFromXlsx } from '../export/xlsx';
import { fcell, ncell, xlsxWith } from '../export/test-xlsx';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  hydrateYDocFromData,
  readSchemaVersion,
  serializeYDocToData,
} from './y-schema';

// S4 added `cachedOnly` and `arraySpan` to a cell's JSON, so the sheets layout is at version 2: an old room still loads, a v1
// client refuses a v2 room, and two current clients converge on the marks.

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });
const LET = '_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)';

async function importedWorkbook(): Promise<WorkbookImpl> {
  const bytes = await xlsxWith({
    1: [ncell('A1', 2), fcell('B1', LET, 99), fcell('C1', '_xlfn._xlws.SORT(A1:A2)', 1, 't="array" ref="C1:C2"', ' cm="1"')],
    2: [ncell('A2', 7), fcell('B2', 'A1*2', 'SAVED'), fcell('D2', '_xlfn.LET(1)', 5, 't="array" ref="D2:E2"'), ncell('E2', 8)],
  });
  const wb = new WorkbookImpl('room', 'WB');
  await importFromXlsx(bytes, wb);
  return wb;
}

describe('formula marks in the Y layout (schema version 2)', () => {
  afterEach(() => __resetInMemoryRooms());

  it('is at version 2 or later, stamped on a hydrated room (S6 raised it to 3)', async () => {
    expect(COLLAB_SCHEMA_VERSION).toBeGreaterThanOrEqual(2);
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, (await importedWorkbook()).getData());
    expect(readSchemaVersion(ydoc)).toBe(COLLAB_SCHEMA_VERSION);
  });

  it('carries cachedOnly and arraySpan through hydrate and serialize', async () => {
    const wb = await importedWorkbook();
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, wb.getData());
    const copy = new WorkbookImpl('copy', 'Copy');
    copy.setData(serializeYDocToData(ydoc));
    expect(copy.getCell(undefined, 0, 1)).toMatchObject({ formula: `=${LET}`, value: 99, cachedOnly: true });
    expect(copy.getCell(undefined, 1, 3)).toMatchObject({ cachedOnly: true, arraySpan: { rows: 1, cols: 2 } });
    expect(copy.getCell(undefined, 0, 2)).toMatchObject({ formula: '=SORT(A1:A2)', arraySpan: { rows: 2, cols: 1, spill: true } });
    expect(copy.getSpillRegion(undefined, 0, 2)).toMatchObject({ rows: 2, cols: 1 });
  });

  it('an old room (no schemaVersion, plain cells) still loads and edits', async () => {
    const a = new WorkbookImpl('room', 'WB');
    a.setCellValue(undefined, 0, 0, 'hello');
    a.setFormula(undefined, 0, 1, '=A1&"!"');
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'old' });
    ha.ydoc.getMap('meta').delete('schemaVersion');
    expect(readSchemaVersion(ha.ydoc)).toBe(1);

    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'old' });
    expect(b.getCell(undefined, 0, 1)).toMatchObject({ formula: '=A1&"!"', value: 'hello!' });
    b.setCellValue(undefined, 0, 0, 'bye');
    expect(a.getCell(undefined, 0, 1)?.value).toBe('bye!');
    ha.detach();
    hb.detach();
  });

  it('a version 1 client refuses a room that carries the marks; a version 2 client joins it', async () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, (await importedWorkbook()).getData());
    // a room as a version 2 client wrote it (the layout S6 added column and row formats to is the next one)
    ydoc.getMap('meta').set('schemaVersion', 2);
    expect(() => assertSchemaVersion(ydoc, 1)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, 2)).not.toThrow();
    try {
      assertSchemaVersion(ydoc, 1);
    } catch (e) {
      expect((e as CollabSchemaVersionError).roomVersion).toBe(2);
      expect((e as CollabSchemaVersionError).supportedVersion).toBe(1);
    }
  });

  it('two clients converge: the marks arrive, cached-only cells keep their saved value on both, an edit clears the mark on both', async () => {
    const a = await importedWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });

    for (const wb of [a, b]) {
      expect(wb.getCell(undefined, 0, 1)).toMatchObject({ formula: `=${LET}`, value: 99, cachedOnly: true });
      expect(wb.getCell(undefined, 1, 1)).toMatchObject({ formula: '=A1*2' });
    }
    // a peer's change of a precedent: the cached-only cell stays as saved on both, the evaluable one follows
    b.setCellValue(undefined, 0, 0, 10);
    for (const wb of [a, b]) {
      expect(wb.getCell(undefined, 0, 1)).toMatchObject({ value: 99, cachedOnly: true });
      expect(wb.getCell(undefined, 1, 1)?.value).toBe(20);
    }
    // an edit of the cell clears the mark for everyone
    a.setFormula(undefined, 0, 1, '=A1+1');
    for (const wb of [a, b]) {
      expect(wb.getCell(undefined, 0, 1)?.cachedOnly).toBeUndefined();
      expect(wb.getCell(undefined, 0, 1)?.value).toBe(11);
    }
    ha.detach();
    hb.detach();
  });

  it('same-cell concurrent edits of a cached-only cell are last writer wins and both clients end identical', async () => {
    const a = await importedWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    a.setCellValue(undefined, 0, 1, 'from A');
    b.setFormula(undefined, 0, 1, '=A1*3');
    const cell = (wb: WorkbookImpl) => {
      const c = wb.getCell(undefined, 0, 1);
      return { formula: c?.formula, value: c?.value, cachedOnly: c?.cachedOnly };
    };
    expect(cell(a)).toEqual(cell(b));
    expect(cell(a).cachedOnly).toBeUndefined();
    ha.detach();
    hb.detach();
  });
});
