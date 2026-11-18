export interface ReportRequest {
  format: 'md' | 'json';
  /** File to write (markdown is appended, as `$GITHUB_STEP_SUMMARY` requires); stdout when absent. */
  out?: string;
}

export interface CliArgs {
  update: boolean;
  matrix: boolean;
  only: string[];
  reports: ReportRequest[];
}

/** Parse the CLI flags; throws on anything unknown or misplaced so a typo in CI fails loudly. */
export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = { update: false, matrix: false, only: [], reports: [] };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--update') args.update = true;
    else if (arg === '--matrix') args.matrix = true;
    else if (arg === '--only') args.only.push(...(argv[++i] ?? '').split(',').filter(Boolean));
    else if (arg === '--report') {
      const format = argv[++i];
      if (format !== 'md' && format !== 'json') throw new Error(`--report expects "md" or "json", got ${String(format)}`);
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
