// The formula conformance suite (plan S11, L5): every row of formula-conformance-oracle.json (the 159-formula
// probe plus the batch-A cases, values recalculated by LibreOffice by fixtures/xlsx/generators/build_conformance_oracle.py)
// is run through the engine. Numbers match to a relative 1e-9, text and booleans exactly, errors by code.
//
// Rows the engine cannot pass yet (functions of later batches, text and date functions, syntax S4 handles) are in
// `KNOWN_FAILING` with the reason. The suite asserts that each of them still fails, so the list has to shrink as
// functions land.

import * as fs from 'fs';
import * as path from 'path';
import { WorkbookImpl } from './workbook';

interface OracleRow {
  id: string;
  function: string;
  formula: string;
  sheet: string | null;
  cell: string;
  cells: Record<string, unknown>;
  group: 'probe' | 'batch-a';
  expectation: 'value' | 'error' | 'array';
  expected: unknown;
  oracle: 'libreoffice';
  spill?: [number, number];
  excel?: true;
  unverified?: true;
  libreoffice?: unknown;
  note?: string;
}

interface Oracle {
  libreoffice: string;
  probeData: { sheet: string; cells: Record<string, unknown> };
  rows: OracleRow[];
}

const oracle = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'formula-conformance-oracle.json'), 'utf8')
) as Oracle;

/** Probe rows the engine fails today, with the reason, until the PR that fixes them lands. */
const KNOWN_FAILING: Record<string, string> = {
  'probe-059': 'TEXT() has no date codes yet (number-format engine, PR S3 / #210)',
  'probe-060': 'TEXT() has no date codes yet (number-format engine, PR S3 / #210)',
  'probe-063': 'CLEAN is a batch B text function (S12)',
  'probe-065': 'NUMBERVALUE is a batch B text function (S12)',
  'probe-066': 'TEXTBEFORE is a batch B text function (S12)',
  'probe-067': 'TEXTSPLIT is a batch B text function (S12)',
  'probe-078': 'DAYS is a batch B date function (S12)',
  'probe-079': 'YEARFRAC is a batch B date function (S12)',
  'probe-080': 'ISOWEEKNUM is a batch B date function (S12)',
  'probe-104': 'XNPV is a batch B finance function (S12)',
  'probe-105': 'XIRR is a batch B finance function (S12)',
  'probe-106': 'SLN is a batch B finance function (S12)',
  'probe-126': 'TRANSPOSE is a batch B array function (S12)',
  'probe-132': 'LET is post-launch (LET and the LAMBDA family)',
  'probe-133': 'HSTACK is a post-launch dynamic-array function',
  'probe-134': 'VSTACK is a post-launch dynamic-array function',
  'probe-135': 'TAKE is a post-launch dynamic-array function',
  'probe-152':
    'the file stores a legacy (implicit intersection) formula that Excel shows as #VALUE! in row 152; the engine always evaluates it as a dynamic array (import policy, S4)',
};

function colIndex(label: string): number {
  let col = 0;
  for (const ch of label) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col - 1;
}

function at(addr: string): [number, number] {
  const m = /^([A-Z]+)(\d+)$/.exec(addr) as RegExpExecArray;
  return [Number(m[2]) - 1, colIndex(m[1])];
}

/** The engine's result for a row: the value of the formula cell, or the block it spills for an array row. */
function evaluate(row: OracleRow): unknown {
  const w = new WorkbookImpl('wb', 'WB');
  const sheetName = row.sheet ?? 'Sheet1';
  const sheetId = w.getSheet().id;
  w.renameSheet(sheetId, sheetName);
  const cells = row.group === 'probe' ? oracle.probeData.cells : row.cells;
  for (const [ref, v] of Object.entries(cells)) {
    const [r, c] = at(ref);
    if (typeof v === 'string' && v.startsWith('=')) w.setFormula(sheetId, r, c, v);
    else w.setCellValue(sheetId, r, c, v);
  }
  const [r0, c0] = at(row.cell);
  w.setFormula(sheetId, r0, c0, `=${row.formula}`);
  const read = (r: number, c: number): unknown => w.getCellCalculatedValue(sheetId, r, c);
  if (!row.spill) return read(r0, c0);
  return Array.from({ length: row.spill[0] }, (_, i) =>
    Array.from({ length: row.spill![1] }, (_, j) => read(r0 + i, c0 + j))
  );
}

const ERROR_CODE = /^#[A-Z0-9/?]+[!?]?|^#N\/A/;

/** An engine error carries trailing detail ("#NAME? Function X not found"): keep the code. */
function normalise(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(normalise);
  if (v == null) return '';
  if (typeof v === 'string' && v.startsWith('#'))
    return (ERROR_CODE.exec(v)?.[0] ?? v).split(' ')[0];
  return v;
}

function sameValue(actual: unknown, expected: unknown): boolean {
  if (Array.isArray(expected)) {
    return (
      Array.isArray(actual) &&
      actual.length === expected.length &&
      expected.every((e, i) => sameValue(actual[i], e))
    );
  }
  if (typeof expected === 'number' && typeof actual === 'number') {
    const d = Math.abs(actual - expected);
    return d <= 1e-9 * Math.max(Math.abs(actual), Math.abs(expected)) || d < 1e-12;
  }
  return actual === expected;
}

function passes(row: OracleRow): { ok: boolean; actual: unknown } {
  const actual = normalise(evaluate(row));
  return { ok: sameValue(actual, row.expected), actual };
}

describe('conformance oracle table', () => {
  it('records the LibreOffice build that produced it', () => {
    expect(oracle.libreoffice).toMatch(/^\d+\.\d+/);
  });

  it('has unique ids and a note on every Excel-documented row', () => {
    const ids = oracle.rows.map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const r of oracle.rows) {
      expect(r.oracle).toBe('libreoffice');
      if (r.excel) {
        expect(r.unverified).toBe(true);
        expect(r.note).toBeTruthy();
        expect(r.libreoffice).not.toEqual(r.expected);
      }
    }
  });

  it('covers the 159 formulas of the probe', () => {
    expect(oracle.rows.filter((r) => r.group === 'probe')).toHaveLength(159);
  });

  it('has at least 120 batch-A rows and three per batch-A function', () => {
    const batchA = oracle.rows.filter((r) => r.group === 'batch-a');
    expect(batchA.length).toBeGreaterThanOrEqual(120);
    const perFunction = new Map<string, number>();
    for (const r of batchA) perFunction.set(r.function, (perFunction.get(r.function) ?? 0) + 1);
    const thin = [...perFunction].filter(([, n]) => n < 3).map(([f]) => f);
    expect(thin).toEqual([]);
    expect([...perFunction.keys()].sort()).toEqual(BATCH_A.slice().sort());
  });

  it('knownFailing lists only rows of the table, each with a reason', () => {
    const ids = new Set(oracle.rows.map((r) => r.id));
    for (const [id, reason] of Object.entries(KNOWN_FAILING)) {
      expect(ids.has(id)).toBe(true);
      expect(reason.length).toBeGreaterThan(10);
    }
  });
});

const BATCH_A = [
  'SIN',
  'COS',
  'TAN',
  'ASIN',
  'ACOS',
  'ATAN',
  'ATAN2',
  'SINH',
  'COSH',
  'TANH',
  'ASINH',
  'ACOSH',
  'ATANH',
  'PI',
  'RADIANS',
  'DEGREES',
  'FACT',
  'FACTDOUBLE',
  'COMBIN',
  'PERMUT',
  'SUMSQ',
  'EVEN',
  'ODD',
  'ISEVEN',
  'ISODD',
  'CEILING.MATH',
  'FLOOR.MATH',
  'MAXIFS',
  'MINIFS',
  'AGGREGATE',
  'STDEV.S',
  'STDEV.P',
  'VAR.S',
  'VAR.P',
  'PERCENTILE.INC',
  'PERCENTILE.EXC',
  'QUARTILE.INC',
  'QUARTILE.EXC',
  'RANK.EQ',
  'MODE.SNGL',
  'SLOPE',
  'INTERCEPT',
  'RSQ',
  'FORECAST',
  'FORECAST.LINEAR',
  'GEOMEAN',
  'FREQUENCY',
];

describe('conformance rows the engine passes', () => {
  const rows = oracle.rows.filter((r) => !(r.id in KNOWN_FAILING));
  it.each(rows.map((r) => [r.id, r.formula, r] as const))('%s  =%s', (_id, _formula, row) => {
    const { ok, actual } = passes(row);
    if (!ok) expect(actual).toEqual(row.expected);
    expect(ok).toBe(true);
  });
});

describe('knownFailing rows still fail', () => {
  const rows = oracle.rows.filter((r) => r.id in KNOWN_FAILING);
  it.each(rows.map((r) => [r.id, r.formula, r] as const))('%s  =%s', (id, _formula, row) => {
    expect(KNOWN_FAILING[id]).toBeTruthy();
    expect(passes(row).ok).toBe(false);
  });

  it('the list is sorted by id and has no stale entry', () => {
    expect(Object.keys(KNOWN_FAILING)).toEqual(Object.keys(KNOWN_FAILING).sort());
  });
});

describe('the comparison', () => {
  it('compares numbers to a relative 1e-9, text exactly and arrays element-wise', () => {
    expect(sameValue(1.0000000001, 1)).toBe(true);
    expect(sameValue(1.00001, 1)).toBe(false);
    expect(sameValue(1e-13, 0)).toBe(true);
    expect(sameValue('a', 'A')).toBe(false);
    expect(sameValue(true, 1)).toBe(false);
    expect(sameValue([[1, 'x']], [[1, 'x']])).toBe(true);
    expect(sameValue([[1]], [[1], [2]])).toBe(false);
    expect(sameValue(1, [[1]])).toBe(false);
  });

  it('compares errors by code', () => {
    expect(normalise('#NAME? Function X not found')).toBe('#NAME?');
    expect(normalise('#N/A')).toBe('#N/A');
    expect(normalise('#DIV/0!')).toBe('#DIV/0!');
    expect(normalise(null)).toBe('');
    expect(normalise([['#VALUE! x']])).toEqual([['#VALUE!']]);
  });
});
