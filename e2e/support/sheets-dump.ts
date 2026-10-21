import type { Page } from '@playwright/test';

export interface DumpedCell {
  value: unknown;
  formula?: string;
  style?: Record<string, unknown>;
  /** Number-format descriptor from the workbook format pool. */
  format?: Record<string, unknown>;
}
export interface DumpedSheet {
  name: string;
  /** A1-keyed cells that have a value, formula, style or format. */
  cells: Record<string, DumpedCell>;
  /** 0-based [startRow, startCol, endRow, endCol]. */
  merges: Array<[number, number, number, number]>;
  frozen: { rows: number; cols: number };
  colWidths: number[];
}

/** Read every sheet of the demo's workbook (model, not pixels) for the first `rows` x `cols` cells. */
export async function dumpWorkbook(page: Page, rows = 12, cols = 8): Promise<DumpedSheet[]> {
  return page.evaluate(
    ([R, C]) => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const wb = (window as any).__spindle.workbook;
      const label = (c: number) => String.fromCharCode(65 + c);
      return wb.getData().sheets.map((sd: { id: string; name: string }) => {
        const sheet = wb.getSheet(sd.id);
        const cells: Record<string, DumpedCell> = {};
        for (let r = 0; r < R; r++) {
          for (let c = 0; c < C; c++) {
            const cell = wb.getCell(sd.id, r, c);
            if (!cell) continue;
            const out: DumpedCell = {
              value: wb.getCellCalculatedValue(sd.id, r, c) ?? cell.value ?? null,
            };
            if (cell.formula) out.formula = cell.formula;
            if (cell.styleId) out.style = wb.getStylePool().get(cell.styleId);
            if (cell.formatId) out.format = wb.getFormatPool().get(cell.formatId);
            if (out.value !== null || out.formula || out.style || out.format)
              cells[`${label(c)}${r + 1}`] = out;
          }
        }
        return {
          name: sd.name,
          cells,
          merges: sheet
            .getMergedRegions()
            .map((m: { startRow: number; startCol: number; endRow: number; endCol: number }) => [
              m.startRow,
              m.startCol,
              m.endRow,
              m.endCol,
            ]),
          frozen: { rows: sheet.getFrozenRows(), cols: sheet.getFrozenCols() },
          colWidths: Array.from({ length: C }, (_, c) => sheet.getColWidth(c)),
        };
      });
    },
    [rows, cols] as const
  );
}
