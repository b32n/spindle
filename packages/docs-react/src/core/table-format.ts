/**
 * Table formatting shared by the page painter and the measurer: where a cell sits in the table's grid, the
 * border lines it shows, its vertical alignment and its width. Both build the same CSS so a table measures at
 * the height it is drawn at.
 */

import type { TableBlock, TableBorder, TableBorders, TableCell } from './flow-blocks';
import { safeCssColor } from './sanitize';

/** The line editor-made tables (no `borders` anywhere) show. */
const DEFAULT_BORDER = '1px solid #dadce0';

export interface CellPlace {
  row: number;
  /** First grid column the cell occupies. */
  col: number;
  colspan: number;
  rowspan: number;
  rows: number;
  cols: number;
}

/** The grid position of every cell, in row order (cells under a row-spanning cell are skipped over). */
export function tableCellPlaces(block: TableBlock): CellPlace[][] {
  const rows = block.rows.length;
  const occupied: boolean[][] = Array.from({ length: rows }, () => []);
  const places: CellPlace[][] = [];
  let cols = 0;
  block.rows.forEach((row, r) => {
    let col = 0;
    const rowPlaces: CellPlace[] = [];
    for (const cell of row.cells) {
      while (occupied[r][col]) col++;
      const colspan = Math.max(1, cell.colspan ?? 1);
      const rowspan = Math.max(1, cell.rowspan ?? 1);
      for (let rr = r; rr < Math.min(rows, r + rowspan); rr++) for (let cc = col; cc < col + colspan; cc++) occupied[rr][cc] = true;
      rowPlaces.push({ row: r, col, colspan, rowspan, rows, cols: 0 });
      col += colspan;
    }
    cols = Math.max(cols, col, occupied[r].length);
    places.push(rowPlaces);
  });
  for (const rowPlaces of places) for (const p of rowPlaces) p.cols = cols;
  return places;
}

function borderCss(side: 'top' | 'right' | 'bottom' | 'left', border: TableBorder | undefined, scale: number): string {
  if (!border || border.style === 'none') return `border-${side}: none;`;
  const width = Math.max(1, (border.widthPx ?? 1) * scale);
  // A double line needs room for its two strokes to show.
  const px = border.style === 'double' ? Math.max(3, width) : width;
  const color = safeCssColor(border.color) ?? '#000000';
  return `border-${side}: ${px}px ${border.style} ${color};`;
}

/** The CSS borders of a cell: its own sides, else the table's line at that side, else (no `borders` anywhere) the default grid. */
export function cellBorderCss(block: TableBlock, cell: TableCell, place: CellPlace, scale: number): string {
  const table: TableBorders | undefined = block.borders;
  if (!table && !cell.borders) return `border: ${DEFAULT_BORDER};`;
  const sides: Array<['top' | 'right' | 'bottom' | 'left', TableBorder | undefined]> = [
    ['top', place.row === 0 ? table?.top : table?.insideH],
    ['right', place.col + place.colspan >= place.cols ? table?.right : table?.insideV],
    ['bottom', place.row + place.rowspan >= place.rows ? table?.bottom : table?.insideH],
    ['left', place.col === 0 ? table?.left : table?.insideV],
  ];
  return sides
    .map(([side, tableLine]) => {
      const own = cell.borders?.[side];
      if (own) return borderCss(side, own, scale);
      if (!table) return `border-${side}: ${DEFAULT_BORDER};`;
      return borderCss(side, tableLine, scale);
    })
    .join('');
}

/** Width and vertical alignment of a cell (width is the sum of the grid columns it spans). */
export function cellLayoutCss(cell: TableCell, scale: number): string {
  let css = `vertical-align: ${cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign === 'bottom' ? 'bottom' : 'top'};`;
  const widths = cell.colwidth?.filter((w) => Number.isFinite(w) && w > 0);
  if (widths && widths.length > 0) css += `width: ${widths.reduce((sum, w) => sum + w, 0) * scale}px;`;
  return css;
}
