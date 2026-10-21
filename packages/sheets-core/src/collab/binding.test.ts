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

// A provider whose connect() stays pending until released — an attach in flight.
class GatedProvider extends InMemoryProvider {
  disconnected = false;
  private release!: () => void;
  private gate = new Promise<void>((r) => (this.release = r));
  override async connect(roomId: string): Promise<void> {
    await this.gate;
    return super.connect(roomId);
  }
  override disconnect(): void {
    this.disconnected = true;
    super.disconnect();
  }
  open(): void {
    this.release();
  }
}

describe('detachCollab() cancels an in-flight attachCollab()', () => {
  it('the aborted attach never seeds the room, leaves no socket, and rejects with AbortError', async () => {
    __resetInMemoryRooms();
    const first = seededWorkbook();
    const gated = new GatedProvider();
    const pending = first.attachCollab(gated, identity('A'), { roomId: 'room' });
    const settled = pending.then(() => 'resolved', (e: Error) => e.name);
    first.detachCollab();
    // The gate opens only AFTER the detach — the old code seeded the room right here.
    gated.open();
    expect(await settled).toBe('AbortError');
    expect(gated.disconnected).toBe(true);
    expect(first.getCollabHandle()).toBeNull();

    // The room was never seeded, so the next attach is the one and only creator.
    const second = seededWorkbook();
    const h2 = await second.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const joiner = new WorkbookImpl('joiner', 'WB');
    const h3 = await joiner.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(joiner.getData().sheets.length).toBe(1);
    h2.detach();
    h3.detach();
  });

  it('a detach/attach pair (React StrictMode) seeds exactly once', async () => {
    __resetInMemoryRooms();
    const d = seededWorkbook();
    const gated = new GatedProvider();
    const stale = d.attachCollab(gated, identity('A'), { roomId: 'room' }).catch((e: Error) => e.name);
    d.detachCollab();
    const handle = await d.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    gated.open();
    expect(await stale).toBe('AbortError');
    expect(d.getData().sheets.length).toBe(1);
    handle.detach();
  });
});
