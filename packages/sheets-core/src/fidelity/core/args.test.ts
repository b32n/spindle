import { DEFAULT_TIMEOUT_SECONDS, parseArgs } from '../../../../../scripts/fidelity-core/args';

describe('parseArgs', () => {
  it('has quiet defaults', () => {
    expect(parseArgs([])).toEqual({
      update: false,
      matrix: false,
      runtime: false,
      only: [],
      reports: [],
      timeoutSeconds: DEFAULT_TIMEOUT_SECONDS,
    });
  });

  it('reads every flag', () => {
    const args = parseArgs([
      '--update',
      '--matrix',
      '--runtime',
      '--only',
      'a,b',
      '--only',
      'c',
      '--timeout',
      '2.5',
      '--report',
      'md',
      '--report-out',
      'x.md',
      '--report',
      'json',
    ]);
    expect(args).toMatchObject({
      update: true,
      matrix: true,
      runtime: true,
      only: ['a', 'b', 'c'],
      timeoutSeconds: 2.5,
    });
    expect(args.reports).toEqual([{ format: 'md', out: 'x.md' }, { format: 'json' }]);
  });

  it('ignores empty --only entries and a missing --only value', () => {
    expect(parseArgs(['--only', 'a,,b,']).only).toEqual(['a', 'b']);
    expect(parseArgs(['--only']).only).toEqual([]);
  });

  it.each([
    [['--frobnicate'], 'unknown argument: --frobnicate'],
    [['--report', 'html'], '--report expects "md" or "json", got html'],
    [['--report'], '--report expects "md" or "json", got undefined'],
    [['--report-out', 'x'], '--report-out must follow a --report'],
    [
      ['--report', 'md', '--report-out', 'a', '--report-out', 'b'],
      '--report-out must follow a --report',
    ],
    [['--report', 'md', '--report-out'], '--report-out needs a file path'],
    [['--timeout', '0'], '--timeout expects a positive number of seconds'],
    [['--timeout', 'abc'], '--timeout expects a positive number of seconds'],
    [['--timeout'], '--timeout expects a positive number of seconds'],
  ])('rejects %j', (argv, message) => {
    expect(() => parseArgs(argv)).toThrow(message);
  });
});
