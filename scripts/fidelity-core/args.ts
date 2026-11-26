export interface ReportRequest {
  format: 'md' | 'json';
  /** File to write (markdown is appended, as `$GITHUB_STEP_SUMMARY` requires); stdout when absent. */
  out?: string;
}

export interface CliArgs {
  update: boolean;
  matrix: boolean;
  /** Also run the files that are generated at run time (bomb, huge, perf); never compared with goldens. */
  runtime: boolean;
  only: string[];
  reports: ReportRequest[];
  /** Wall-clock limit for one file, in seconds. */
  timeoutSeconds: number;
}

export const DEFAULT_TIMEOUT_SECONDS = 60;

/** Parse the CLI flags; throws on anything unknown or misplaced so a typo in CI fails loudly. */
export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    update: false,
    matrix: false,
    runtime: false,
    only: [],
    reports: [],
    timeoutSeconds: DEFAULT_TIMEOUT_SECONDS,
  };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--update') args.update = true;
    else if (arg === '--matrix') args.matrix = true;
    else if (arg === '--runtime') args.runtime = true;
    else if (arg === '--only') args.only.push(...(argv[++i] ?? '').split(',').filter(Boolean));
    else if (arg === '--timeout') {
      const seconds = Number(argv[++i]);
      if (!Number.isFinite(seconds) || seconds <= 0)
        throw new Error('--timeout expects a positive number of seconds');
      args.timeoutSeconds = seconds;
    } else if (arg === '--report') {
      const format = argv[++i];
      if (format !== 'md' && format !== 'json')
        throw new Error(`--report expects "md" or "json", got ${String(format)}`);
      args.reports.push({ format });
    } else if (arg === '--report-out') {
      const last = args.reports[args.reports.length - 1];
      const out = argv[++i];
      if (!last || last.out) throw new Error('--report-out must follow a --report');
      if (!out) throw new Error('--report-out needs a file path');
      last.out = out;
    } else throw new Error(`unknown argument: ${arg}`);
  }
  return args;
}
