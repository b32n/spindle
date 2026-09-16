import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from '../workbook';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

function seededWorkbook(): WorkbookImpl {
  const wb = new WorkbookImpl('room', 'WB');
  wb.setCellValue(undefined, 0, 0, 'hello');
  wb.setCellValue(undefined, 0, 1, 'world');
  wb.setCellValue(undefined, 1, 0, 42);
  return wb;
}

afterEach(() => __resetInMemoryRooms());

describe('workbook collab seeding', () => {
  it('a joiner that also holds the same data does not duplicate sheets or cells', async () => {
    __resetInMemoryRooms();
    const a = seededWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    // B loaded the same document from its own backend before joining — this is
    // exactly what used to push a second entry into the sheetIds/rowOrder
    // arrays and surface a duplicate sheet.
    const b = seededWorkbook();
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });

    expect(b.getData().sheets.length).toBe(1);
    expect(a.getData().sheets.length).toBe(1);
    expect(b.getSheet().getCellValue(0, 0)).toBe('hello');
    expect(b.getSheet().getCellValue(1, 0)).toBe(42);
    // No phantom extra row from a duplicated rowOrder.
    expect(b.getSheet().getCellValue(2, 0)).toBeNull();

    ha.detach();
    hb.detach();
  });

  it('the initial seed is not undoable', async () => {
    __resetInMemoryRooms();
    const a = seededWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });

    // Seeding runs under SEED_ORIGIN, which the UndoManager does not track, so
    // there is nothing to undo and the content must survive.
    expect(a.undo()).toBe(false);
    expect(a.getSheet().getCellValue(0, 0)).toBe('hello');

    ha.detach();
  });
});
