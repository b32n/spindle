import { createLossReport, type ImportLossReport } from '../loss';

/**
 * CSV carries no formatting, so nothing is lost by importing one. The one
 * thing worth warning about is a delimiter the importer does not split on:
 * semicolon- or tab-separated files import as a single column per line.
 */
export function detectCsvLoss(text: string): ImportLossReport {
  const lines = text.split(/\r?\n/).filter((l) => l.trim() !== '').slice(0, 20);
  const countOutsideQuotes = (line: string, ch: string) => {
    let n = 0;
    let quoted = false;
    for (const c of line) {
      if (c === '"') quoted = !quoted;
      else if (c === ch && !quoted) n++;
    }
    return n;
  };
  // A delimiter "is" the file's delimiter when every line has the same number of it (≥ 1).
  const consistent = (ch: string): number => {
    const counts = lines.map((l) => countOutsideQuotes(l, ch));
    return counts.length > 0 && counts[0] >= 1 && counts.every((c) => c === counts[0]) ? counts[0] : 0;
  };
  const commas = consistent(',');
  const semis = consistent(';');
  const tabs = consistent('\t');
  // Decimal commas ("2,5") make the comma count uneven, so the test is "no consistent comma, but
  // a consistent ; or tab", not "no comma at all".
  if (lines.length > 0 && commas === 0 && (semis > 0 || tabs > 0)) {
    return createLossReport('csv', [
      {
        id: 'csv.delimiter',
        label: semis >= tabs ? 'semicolon-separated line(s)' : 'tab-separated line(s)',
        count: lines.length,
        effect: 'only commas split columns, so each line imports into a single column',
      },
    ]);
  }
  return createLossReport('csv', []);
}
