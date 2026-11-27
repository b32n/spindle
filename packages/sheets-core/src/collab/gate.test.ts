import * as Y from 'yjs';
import { InMemoryProvider, __resetInMemoryRooms, type CollabIdentity } from '@b32nio/spindle-shared';
import { WorkbookImpl } from '../workbook';
import * as publicApi from '../index';
import * as collabApi from './index';
import * as blockersModule from './blockers';
import * as ySchemaModule from './y-schema';
import {
  COLLAB_SCHEMA_VERSION,
  CollabSchemaVersionError,
  assertSchemaVersion,
  hydrateYDocFromData,
  readSchemaVersion,
} from './y-schema';
import { CollabBlockedError, getCollabBlockers } from './blockers';

const identity = (name: string): CollabIdentity => ({ userId: name, displayName: name, color: '#123456' });

function seededWorkbook(): WorkbookImpl {
  const wb = new WorkbookImpl('room', 'WB');
  wb.setCellValue(undefined, 0, 0, 'hello');
  wb.setCellValue(undefined, 1, 0, 42);
  return wb;
}

describe('getCollabBlockers', () => {
  it('is empty for a workbook, a workbook snapshot and a bare sheet list', () => {
    expect(getCollabBlockers(new WorkbookImpl('a', 'A').getData())).toEqual([]);
    expect(getCollabBlockers(seededWorkbook().getData())).toEqual([]);
    expect(getCollabBlockers({ sheets: [] })).toEqual([]);
  });
});

describe('CollabBlockedError', () => {
  it('carries the blockers and lists them in the message', () => {
    const blockers = [
      { id: 'collab.a', message: 'First thing', count: 2 },
      { id: 'collab.b', message: 'Second thing', count: 1 },
    ];
    const err = new CollabBlockedError(blockers);
    expect(err).toBeInstanceOf(Error);
    expect(err.name).toBe('CollabBlockedError');
    expect(err.blockers).toBe(blockers);
    expect(err.message).toBe('Collaboration is not available for this workbook: First thing (2); Second thing (1)');
  });
});

describe('schema version', () => {
  it('reads a missing key as 1 and ignores values that are not a version', () => {
    const ydoc = new Y.Doc();
    expect(readSchemaVersion(ydoc)).toBe(1);
    const meta = ydoc.getMap<unknown>('meta');
    for (const bad of ['3', null, NaN, Infinity, 0, -2, {}]) {
      meta.set('schemaVersion', bad);
      expect(readSchemaVersion(ydoc)).toBe(1);
    }
    meta.set('schemaVersion', 5);
    expect(readSchemaVersion(ydoc)).toBe(5);
  });

  it('hydration stamps the current version', () => {
    const ydoc = new Y.Doc();
    hydrateYDocFromData(ydoc, seededWorkbook().getData());
    expect(ydoc.getMap('meta').get('schemaVersion')).toBe(COLLAB_SCHEMA_VERSION);
    expect(readSchemaVersion(ydoc)).toBe(COLLAB_SCHEMA_VERSION);
  });

  it('assertSchemaVersion accepts the supported version and older, refuses newer', () => {
    const ydoc = new Y.Doc();
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
    ydoc.getMap('meta').set('schemaVersion', COLLAB_SCHEMA_VERSION);
    expect(() => assertSchemaVersion(ydoc)).not.toThrow();
    ydoc.getMap('meta').set('schemaVersion', COLLAB_SCHEMA_VERSION + 1);
    expect(() => assertSchemaVersion(ydoc)).toThrow(CollabSchemaVersionError);
    expect(() => assertSchemaVersion(ydoc, COLLAB_SCHEMA_VERSION + 1)).not.toThrow();
  });

  it('the error names both versions and tells the user what to do', () => {
    const err = new CollabSchemaVersionError(9);
    expect(err.name).toBe('CollabSchemaVersionError');
    expect(err.roomVersion).toBe(9);
    expect(err.supportedVersion).toBe(COLLAB_SCHEMA_VERSION);
    expect(err.message).toContain('version 9');
    expect(err.message).toContain(`up to version ${COLLAB_SCHEMA_VERSION}`);
    expect(err.message).toContain('Update the editor');
    expect(new CollabSchemaVersionError(9, 4).supportedVersion).toBe(4);
  });
});

// A client that writes a newer layout stamps the version in the same transaction that seeds the room.
async function attachNewerRoom(wb: WorkbookImpl, roomId: string, version: number) {
  const original = ySchemaModule.hydrateYDocFromData;
  const spy = jest.spyOn(ySchemaModule, 'hydrateYDocFromData').mockImplementation((ydoc, data) => {
    original(ydoc, data);
    ydoc.getMap('meta').set('schemaVersion', version);
  });
  try {
    return await wb.attachCollab(new InMemoryProvider(), identity('A'), { roomId });
  } finally {
    spy.mockRestore();
  }
}

describe('attach gate', () => {
  afterEach(() => {
    jest.restoreAllMocks();
    __resetInMemoryRooms();
  });

  it('refuses a room written by a newer client, clearly, and never seeds or attaches', async () => {
    const a = seededWorkbook();
    const ha = await attachNewerRoom(a, 'room', COLLAB_SCHEMA_VERSION + 1);

    const b = new WorkbookImpl('b', 'B');
    const before = b.getData();
    const err = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(CollabSchemaVersionError);
    expect((err as Error).message).toMatch(/schema version 2.*only supports up to version 1/);
    expect(b.getCollabHandle()).toBeNull();
    expect(b.getData().sheets).toEqual(before.sheets);
    expect(b.getSheet().getCellValue(0, 0)).toBeNull();
    ha.detach();
  });

  it('a refused attach leaves the workbook free to attach to another room', async () => {
    const a = seededWorkbook();
    const ha = await attachNewerRoom(a, 'newer', 99);
    const b = new WorkbookImpl('b', 'B');
    await expect(b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'newer' })).rejects.toBeInstanceOf(
      CollabSchemaVersionError,
    );
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'other' });
    expect(b.getCollabHandle()).toBe(hb);
    ha.detach();
    hb.detach();
  });

  it('joins a room written before the key existed (missing = 1) and converges', async () => {
    const a = seededWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    ha.ydoc.getMap('meta').delete('schemaVersion');
    expect(readSchemaVersion(ha.ydoc)).toBe(1);

    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(b.getSheet().getCellValue(0, 0)).toBe('hello');
    expect(b.getSheet().getCellValue(1, 0)).toBe(42);

    a.setCellValue(undefined, 2, 0, 'later');
    expect(b.getSheet().getCellValue(2, 0)).toBe('later');
    ha.detach();
    hb.detach();
  });

  it('two current-version clients still converge, and the room carries the version', async () => {
    const a = seededWorkbook();
    const ha = await a.attachCollab(new InMemoryProvider(), identity('A'), { roomId: 'room' });
    const b = new WorkbookImpl('b', 'B');
    const hb = await b.attachCollab(new InMemoryProvider(), identity('B'), { roomId: 'room' });
    expect(readSchemaVersion(hb.ydoc)).toBe(COLLAB_SCHEMA_VERSION);
    expect(b.getSheet().getCellValue(0, 0)).toBe('hello');
    a.setCellValue(undefined, 3, 3, 7);
    b.setCellValue(undefined, 4, 4, 8);
    expect(a.getSheet().getCellValue(4, 4)).toBe(8);
    expect(b.getSheet().getCellValue(3, 3)).toBe(7);
    ha.detach();
    hb.detach();
  });

  it('rejects with CollabBlockedError carrying the blockers before any room is touched', async () => {
    const blockers = [{ id: 'collab.test', message: 'A test construct', count: 3 }];
    const spy = jest.spyOn(blockersModule, 'getCollabBlockers').mockReturnValue(blockers);
    const provider = new InMemoryProvider();
    const connect = jest.spyOn(provider, 'connect');
    const wb = seededWorkbook();
    const err = await wb.attachCollab(provider, identity('A'), { roomId: 'room' }).catch((e: unknown) => e);
    expect(err).toBeInstanceOf(CollabBlockedError);
    expect((err as CollabBlockedError).blockers).toEqual(blockers);
    expect((err as Error).message).toContain('A test construct (3)');
    expect(spy).toHaveBeenCalledTimes(1);
    expect(connect).not.toHaveBeenCalled();
    expect(wb.getCollabHandle()).toBeNull();
  });
});

describe('public exports', () => {
  it('exposes the gate from the base entry and the collab subpath', () => {
    expect(publicApi.getCollabBlockers).toBe(getCollabBlockers);
    expect(publicApi.CollabBlockedError).toBe(CollabBlockedError);
    expect(collabApi.getCollabBlockers).toBe(getCollabBlockers);
    expect(collabApi.CollabBlockedError).toBe(CollabBlockedError);
    expect(collabApi.CollabSchemaVersionError).toBe(CollabSchemaVersionError);
    expect(collabApi.COLLAB_SCHEMA_VERSION).toBe(COLLAB_SCHEMA_VERSION);
    expect(collabApi.readSchemaVersion).toBe(readSchemaVersion);
    expect(collabApi.assertSchemaVersion).toBe(assertSchemaVersion);
  });
});
