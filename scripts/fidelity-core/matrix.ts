/** Generic coverage-matrix helpers (format-neutral): the registry is passed in. */

export interface MatrixFixture {
  id: string;
  constructs: string[];
  title?: string;
  source?: { tool: string };
}

export interface RegistryItem {
  id: string;
  label: string;
}

/** Keep a table cell on one line and stop a `|` from splitting it. */
const cell = (text: string) => text.replace(/\|/g, '\\|').replace(/[\r\n]+/g, ' ');

/**
 * Markdown for the generated README block: a numbered file list (title and producer, from the
 * manifest) followed by the construct x file matrix. A cell is "yes" only for constructs the manifest claims.
 */
export function buildMatrix(registry: RegistryItem[], fixtures: MatrixFixture[]): string {
  const files = [
    '| # | Fixture | Title | Produced by |',
    '|--:|---|---|---|',
    ...fixtures.map(
      (f, i) =>
        `| ${i + 1} | \`${f.id}\` | ${cell(f.title ?? '')} | ${cell(f.source?.tool ?? '')} |`
    ),
  ];
  const header = `| Construct | ${fixtures.map((_, i) => i + 1).join(' | ')} |`;
  const sep = `|---|${fixtures.map(() => ':-:').join('|')}|`;
  const rows = registry.map(
    (c) =>
      `| ${c.label} (\`${c.id}\`) | ${fixtures.map((f) => (f.constructs.includes(c.id) ? 'yes' : '')).join(' | ')} |`
  );
  return [...files, '', header, sep, ...rows].join('\n');
}

/** Registry ids that no fixture claims. */
export function uncoveredConstructs(
  registry: RegistryItem[],
  fixtures: Array<{ constructs: string[] }>
): string[] {
  return registry.map((c) => c.id).filter((id) => !fixtures.some((f) => f.constructs.includes(id)));
}

/** Replace the generated block between `start` and `end` markers in a README. */
export function injectMatrix(readme: string, matrix: string, start: string, end: string): string {
  const a = readme.indexOf(start);
  const b = readme.indexOf(end);
  if (a < 0 || b < a) throw new Error('README is missing the coverage-matrix markers');
  return `${readme.slice(0, a)}${start}\n\n${matrix}\n\n${readme.slice(b)}`;
}
