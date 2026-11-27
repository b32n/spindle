import * as fs from 'node:fs';
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, symlinkSync, utimesSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import JSZip from 'jszip';
import { describeIf, endOf, heavyAllowed, importInChild, lastJson, MB, memoryBudget, runChild, slotCount, survived, SURVIVED, sweepStale, withSlot } from './child';

// `renameSync` can be made to fail or to lose a race for one call, and does what it does otherwise
const realRename = jest.requireActual<typeof fs>('node:fs').renameSync;
jest.mock('node:fs', () => {
  const actual = jest.requireActual<typeof fs>('node:fs');
  return { ...actual, renameSync: jest.fn((from: fs.PathLike, to: fs.PathLike) => actual.renameSync(from, to)) };
});
const renameMock = jest.mocked(fs.renameSync);

const BEHAVE = resolve(__dirname, 'behave-probe.ts');
const GB = 1024 * MB;
const run = (mode: string, timeoutMs = 30_000, gc = false) => runChild<{ a?: number }>(BEHAVE, [mode], { heapMb: 128, timeoutMs, gc });

describe('what a machine may run: memory budget, heavy gate, slots', () => {
  it('takes the container limit when there is one below the host memory, else the host memory', () => {
    expect(memoryBudget(4 * GB, 16 * GB)).toBe(4 * GB);
    expect(memoryBudget(undefined, 16 * GB)).toBe(16 * GB);
    expect(memoryBudget(0, 16 * GB)).toBe(16 * GB);
    // no limit is reported as a number above the host's
    expect(memoryBudget(2 ** 60, 16 * GB)).toBe(16 * GB);
  });

  it('defaults to the process own limits', () => {
    expect(memoryBudget()).toBeGreaterThan(0);
  });

  it('allows the heavy shapes from 6 GB, when 2 GB are free (where that is known), unless skipped', () => {
    expect(heavyAllowed(6 * GB, undefined, undefined)).toBe(true);
    expect(heavyAllowed(6 * GB - 1, undefined, undefined)).toBe(false);
    expect(heavyAllowed(16 * GB, 2 * GB, undefined)).toBe(true);
    expect(heavyAllowed(16 * GB, 2 * GB - 1, undefined)).toBe(false);
    expect(heavyAllowed(16 * GB, 8 * GB, '1')).toBe(false);
  });

  it('runs one to three children at once: one per 3 GB, no more than the CPUs', () => {
    expect(slotCount(2 * GB, 8)).toBe(1);
    expect(slotCount(6 * GB, 8)).toBe(2);
    expect(slotCount(64 * GB, 8)).toBe(3);
    expect(slotCount(64 * GB, 2)).toBe(2);
    expect(slotCount(64 * GB, 1)).toBe(1);
    expect(slotCount(8 * GB)).toBeGreaterThanOrEqual(1);
  });
});

describe('withSlot', () => {
  let root: string;
  beforeEach(() => {
    root = mkdtempSync(join(tmpdir(), 'slot-test-'));
  });
  afterEach(() => {
    renameMock.mockReset().mockImplementation(realRename);
    rmSync(root, { recursive: true, force: true });
  });

  it('holds a slot while it runs and frees it after, also when it throws', () => {
    expect(withSlot(() => readdirSync(root), 1, 1_000, root)).toEqual(['spindle-hostile-slot-0']);
    expect(readdirSync(root)).toEqual([]);
    expect(() => withSlot(() => { throw new Error('x'); }, 1, 1_000, root)).toThrow('x');
    expect(readdirSync(root)).toEqual([]);
  });

  it('takes the next free slot, and runs anyway after the wait when all are held by live processes', () => {
    withSlot(() => {
      // slot 0 is ours: a second user takes slot 1
      expect(withSlot(() => readdirSync(root).sort(), 2, 1_000, root)).toEqual(['spindle-hostile-slot-0', 'spindle-hostile-slot-1']);
      // and with one slot only, waits its 150 ms and runs without a slot (and frees none of ours)
      const started = Date.now();
      expect(withSlot(() => 'ran', 1, 150, root)).toBe('ran');
      expect(Date.now() - started).toBeGreaterThanOrEqual(140);
      expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']);
    }, 2, 1_000, root);
  });

  it('takes over the slot of a holder that is gone', () => {
    const dir = join(root, 'spindle-hostile-slot-0');
    mkdirSync(dir);
    // a pid no process has (above the kernel limit)
    writeFileSync(join(dir, 'pid'), '99999999');
    expect(withSlot(() => 'ran', 1, 5_000, root)).toBe('ran');
    expect(readdirSync(root)).toEqual([]);
  });

  it('waits for a holder that has not named itself yet, and one that is somebody else\'s (EPERM) is alive', () => {
    const dir = join(root, 'spindle-hostile-slot-0');
    mkdirSync(dir);
    // no pid file: not taken over, so the wait runs out
    expect(withSlot(() => 'ran', 1, 100, root)).toBe('ran');
    expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']);
    writeFileSync(join(dir, 'pid'), '12345');
    const kill = jest.spyOn(process, 'kill').mockImplementation(() => {
      throw Object.assign(new Error('not yours'), { code: 'EPERM' });
    });
    try {
      expect(withSlot(() => 'ran', 1, 100, root)).toBe('ran');
      expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']);
    } finally {
      kill.mockRestore();
    }
  });

  it('loses nothing when two waiters see the same dead holder: the one that loses the rename leaves the slot to the winner', () => {
    const dir = join(root, 'spindle-hostile-slot-0');
    mkdirSync(dir);
    writeFileSync(join(dir, 'pid'), '99999999');
    // the other waiter has renamed it (and gone on) just before this one does
    renameMock.mockImplementationOnce(() => {
      throw Object.assign(new Error('gone'), { code: 'ENOENT' });
    });
    expect(withSlot(() => 'ran', 1, 5_000, root)).toBe('ran');
    // the dead slot was still there for the next try: taken over, then freed
    expect(readdirSync(root)).toEqual([]);
  });

  it('does not remove a slot that another has taken since the pid was read, and puts it back', () => {
    const dir = join(root, 'spindle-hostile-slot-0');
    mkdirSync(dir);
    writeFileSync(join(dir, 'pid'), '99999999');
    let seen = '';
    // between the read of the dead pid and the claim, another process took the slot (it names itself)
    renameMock.mockImplementationOnce((from, to) => {
      writeFileSync(join(String(from), 'pid'), String(process.pid));
      realRename(from, to);
    });
    withSlot(() => {
      seen = readFileSync(join(dir, 'pid'), 'utf8');
    }, 1, 100, root);
    // it was ours to find alive (a live pid), so the slot was put back and this waiter ran after its wait without it
    expect(seen).toBe(String(process.pid));
    expect(existsSync(dir)).toBe(true);
    expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']);
  });

  it('removes what it claimed when it cannot put it back', () => {
    const dir = join(root, 'spindle-hostile-slot-0');
    mkdirSync(dir);
    writeFileSync(join(dir, 'pid'), '99999999');
    renameMock
      .mockImplementationOnce((from, to) => {
        writeFileSync(join(String(from), 'pid'), String(process.pid));
        realRename(from, to);
      })
      .mockImplementationOnce(() => {
        throw Object.assign(new Error('taken'), { code: 'EEXIST' });
      });
    expect(withSlot(() => 'ran', 1, 5_000, root)).toBe('ran');
    expect(readdirSync(root)).toEqual([]);
  });

  it('frees only a slot that still names it', () => {
    withSlot(() => writeFileSync(join(root, 'spindle-hostile-slot-0', 'pid'), '1'), 1, 1_000, root);
    expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']);
  });

  it('uses the machine default slots in the temp directory', () => {
    expect(withSlot(() => 7)).toBe(7);
  });
});

describe('sweepStale', () => {
  it('removes what its own names left a day ago, and nothing else, and looks once for a directory', () => {
    const root = mkdtempSync(join(tmpdir(), 'sweep-test-'));
    try {
      const old = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000);
      for (const name of ['spindle-hostile-AAAA', 'spindle-hostile-slot-0', 'other-old', 'spindle-hostile-fresh']) mkdirSync(join(root, name));
      writeFileSync(join(root, 'spindle-hostile-AAAA', 'big'), 'x');
      for (const name of ['spindle-hostile-AAAA', 'spindle-hostile-slot-0', 'other-old']) utimesSync(join(root, name), old, old);
      // a link to nothing (it is gone between the listing and the look): left alone
      symlinkSync(join(root, 'nowhere'), join(root, 'spindle-hostile-dangling'));
      sweepStale(root);
      expect(readdirSync(root).sort()).toEqual(['other-old', 'spindle-hostile-dangling', 'spindle-hostile-fresh']);
      // once for each root: a second call does not look again
      mkdirSync(join(root, 'spindle-hostile-BBBB'));
      utimesSync(join(root, 'spindle-hostile-BBBB'), old, old);
      sweepStale(root);
      expect(existsSync(join(root, 'spindle-hostile-BBBB'))).toBe(true);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });

  it('is done by withSlot before it takes a slot', () => {
    const root = mkdtempSync(join(tmpdir(), 'sweep-test-'));
    try {
      const old = new Date(Date.now() - 3 * 24 * 60 * 60 * 1000);
      mkdirSync(join(root, 'spindle-hostile-OLD'));
      utimesSync(join(root, 'spindle-hostile-OLD'), old, old);
      withSlot(() => expect(readdirSync(root)).toEqual(['spindle-hostile-slot-0']), 1, 1_000, root);
    } finally {
      rmSync(root, { recursive: true, force: true });
    }
  });
});

describe('runChild: the endings of a child', () => {
  it('gives the last line of output that is JSON', () => {
    const r = run('json');
    expect(r.out).toEqual({ a: 1 });
    expect(survived(r)).toEqual(SURVIVED);
  });

  it('gives {} for output that is not JSON, or cut short, instead of throwing', () => {
    expect(run('cut').out).toEqual({});
    expect(run('text').out).toEqual({});
    expect(run('none').out).toEqual({});
  });

  it('says why a child failed: its status and stderr', () => {
    const r = run('fail');
    expect(survived(r)).toEqual({ status: 3, signal: null, stderr: expect.stringContaining('boom') });
  });

  it('says why a child that overran its time limit was killed: the signal and the spawn error', () => {
    const r = run('sleep', 1_500);
    expect(r.out).toEqual({});
    expect(r.status).toBeNull();
    expect(r.signal).toBe('SIGTERM');
    expect(r.stderr).toMatch(/ETIMEDOUT/);
  });

  it('exposes the collector to a child that asks for it, and only then', () => {
    expect(run('gc', 30_000, true).out).toEqual({ a: 1 });
    expect(run('gc').out).toEqual({ a: 0 });
  });

  it('reports a child that never started (no stdout, no stderr) by its spawn error', () => {
    expect(endOf({ stdout: null as unknown as string, stderr: null as unknown as string, status: null, signal: null, error: new Error('spawn ENOENT') })).toEqual({ out: {}, status: null, signal: null, stderr: 'spawn ENOENT' });
    expect(endOf({ stdout: '{"a":1}', stderr: '', status: 0, signal: null })).toEqual({ out: { a: 1 }, status: 0, signal: null, stderr: '' });
  });

  it('survived names the error of an import that threw, and the stderr of any child that did not end well', () => {
    expect(survived({ status: 1, signal: null, stderr: 'trace', probe: { error: 'ImportError: x' } as never })).toEqual({ status: 1, signal: null, error: 'ImportError: x', stderr: 'trace' });
    expect(survived({ status: 0, signal: null, stderr: 'warning noise' })).toEqual(SURVIVED);
  });

  it('describeIf is describe where the shapes run and describe.skip where they do not', () => {
    expect(describeIf(true)).toBe(describe);
    expect(describeIf(false)).toBe(describe.skip);
  });

  it('lastJson takes the last line, and no line is {}', () => {
    expect(lastJson<{ a: number }>('x\n{"a":2}\n')).toEqual({ a: 2 });
    expect(lastJson('')).toEqual({});
  });
});

describe('importInChild', () => {
  const tiny = async (): Promise<Uint8Array> => {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    return zip.generateAsync({ type: 'uint8array' });
  };

  it('removes its scratch directory whether the child ends well, fails or is killed', async () => {
    const before = process.env.TMPDIR;
    const scratch = mkdtempSync(join(tmpdir(), 'scratch-root-'));
    process.env.TMPDIR = scratch;
    try {
      const bytes = await tiny();
      // a package with no workbook: the import throws in the child (status 1, the error in `probe`)
      const failed = importInChild(bytes, { heapMb: 256, timeoutMs: 60_000 });
      expect(failed.status).toBe(1);
      expect(failed.probe.error).toMatch(/ImportError/);
      expect(readdirSync(scratch)).toEqual([]);
      // limits are passed through, and a child killed at once leaves nothing either
      const killed = importInChild(bytes, { heapMb: 256, timeoutMs: 1, limits: { maxSheets: 1 } });
      expect(killed.status).toBeNull();
      expect(readdirSync(scratch)).toEqual([]);
    } finally {
      if (before === undefined) delete process.env.TMPDIR;
      else process.env.TMPDIR = before;
      rmSync(scratch, { recursive: true, force: true });
    }
  }, 120_000);
});
