import { columnIndexToLabel } from '@b32nio/spindle-sheets-core';
import type { Range } from '@b32nio/spindle-sheets-core';

/**
 * Human-readable A1-style label for a selected range, used by the name box and
 * exposed on the canvas for assistive tech / tests.
 *
 *   single cell        -> "B2"
 *   block              -> "B2:D5"
 *   whole column(s)    -> "B:B" / "B:D"
 *   whole row(s)       -> "2:2" / "2:5"
 *   entire sheet       -> "A1:CV1000" (whole-row and whole-column at once)
 */
export function formatSelectionLabel(range: Range | undefined, rowCount: number, colCount: number): string {
  if (!range) return '';
  const r1 = Math.min(range.startRow, range.endRow);
  const r2 = Math.max(range.startRow, range.endRow);
  const c1 = Math.min(range.startCol, range.endCol);
  const c2 = Math.max(range.startCol, range.endCol);
  const wholeCols = r1 === 0 && r2 >= rowCount - 1;
  const wholeRows = c1 === 0 && c2 >= colCount - 1;
  if (wholeCols && !wholeRows) return `${columnIndexToLabel(c1)}:${columnIndexToLabel(c2)}`;
  if (wholeRows && !wholeCols) return `${r1 + 1}:${r2 + 1}`;
  const start = `${columnIndexToLabel(c1)}${r1 + 1}`;
  if (r1 === r2 && c1 === c2) return start;
  return `${start}:${columnIndexToLabel(c2)}${r2 + 1}`;
}
