import { expect, type Locator, type Page } from '@playwright/test';

export const SHEETS_URL = 'http://127.0.0.1:5175';

// Grid geometry of the sheets demo (CanvasGrid defaults).
export const HEADER_W = 100;
export const HEADER_H = 20;
export const COL_W = 100;
export const ROW_H = 20;

export const isMac = process.platform === 'darwin';
export const MOD = isMac ? 'Meta' : 'Control';

/** Page object for the canvas-based sheets editor (single-editor demo). */
export class SheetsPage {
  readonly canvas: Locator;
  readonly nameBox: Locator;
  readonly formulaInput: Locator;

  constructor(readonly page: Page) {
    this.canvas = page.getByTestId('sheet-canvas');
    this.nameBox = page.getByTestId('name-box');
    this.formulaInput = page.getByPlaceholder('Enter formula or value');
  }

  async goto(query = ''): Promise<void> {
    await this.page.goto(`${SHEETS_URL}/${query}`);
    await expect(this.canvas).toBeVisible();
    // The model handle is installed by the demo once React has mounted.
    await this.page.waitForFunction(() => !!(window as unknown as { __spindle?: unknown }).__spindle);
  }

  // ---- geometry -----------------------------------------------------------
  async point(x: number, y: number): Promise<{ x: number; y: number }> {
    const box = await this.canvas.boundingBox();
    if (!box) throw new Error('canvas not laid out');
    return { x: box.x + x, y: box.y + y };
  }
  async cellCenter(row: number, col: number) {
    return this.point(HEADER_W + col * COL_W + COL_W / 2, HEADER_H + row * ROW_H + ROW_H / 2);
  }
  async colHeaderCenter(col: number) {
    return this.point(HEADER_W + col * COL_W + COL_W / 2, HEADER_H / 2);
  }
  async rowHeaderCenter(row: number) {
    return this.point(HEADER_W / 2, HEADER_H + row * ROW_H + ROW_H / 2);
  }
  async corner() {
    return this.point(HEADER_W / 2 - 20, HEADER_H / 2);
  }

  // ---- interactions -------------------------------------------------------
  async clickCell(row: number, col: number, opts: { modifiers?: Array<'Shift' | 'Control' | 'Meta' | 'Alt'> } = {}) {
    await this.click(await this.cellCenter(row, col), opts.modifiers);
  }
  async dblClickCell(row: number, col: number) {
    const p = await this.cellCenter(row, col);
    await this.page.mouse.dblclick(p.x, p.y);
  }
  async click(p: { x: number; y: number }, modifiers: readonly string[] = []) {
    for (const m of modifiers) await this.page.keyboard.down(m);
    await this.page.mouse.click(p.x, p.y);
    for (const m of [...modifiers].reverse()) await this.page.keyboard.up(m);
  }
  async drag(from: { x: number; y: number }, to: { x: number; y: number }) {
    await this.page.mouse.move(from.x, from.y);
    await this.page.mouse.down();
    await this.page.mouse.move((from.x + to.x) / 2, (from.y + to.y) / 2, { steps: 4 });
    await this.page.mouse.move(to.x, to.y, { steps: 4 });
    await this.page.mouse.up();
  }
  async select(row: number, col: number) {
    await this.clickCell(row, col);
    await expect(this.canvas).toHaveAttribute('data-active-cell', `${colLabel(col)}${row + 1}`);
  }
  /** Type into the active cell (starts edit mode) and commit with Enter. */
  async type(text: string, commit: 'Enter' | 'Tab' | 'Escape' | null = 'Enter') {
    await this.page.keyboard.type(text);
    if (commit) await this.page.keyboard.press(commit);
  }
  async enter(row: number, col: number, text: string) {
    await this.select(row, col);
    await this.type(text);
  }
  async focusGrid() {
    await this.canvas.focus();
  }

  // ---- model inspection ---------------------------------------------------
  async selectionLabel(): Promise<string> {
    return (await this.canvas.getAttribute('data-selection')) ?? '';
  }
  async modelSelection() {
    return this.page.evaluate(() => (window as any).__spindle.workbook.getSelection());
  }
  async cellValue(row: number, col: number) {
    return this.page.evaluate(
      ([r, c]) => {
        const wb = (window as any).__spindle.workbook;
        const spilled = wb.getSpilledValue(undefined, r, c);
        return spilled !== undefined ? spilled : (wb.getCell(undefined, r, c)?.value ?? null);
      },
      [row, col],
    );
  }
  async cellFormula(row: number, col: number) {
    return this.page.evaluate(
      ([r, c]) => (window as any).__spindle.workbook.getCell(undefined, r, c)?.formula ?? null,
      [row, col],
    );
  }
  async cellStyle(row: number, col: number): Promise<Record<string, unknown>> {
    return this.page.evaluate(([r, c]) => {
      const wb = (window as any).__spindle.workbook;
      const cell = wb.getCell(undefined, r, c);
      if (!cell?.styleId) return {};
      return wb.getStylePool().get(cell.styleId) ?? {};
    }, [row, col]);
  }
  async expectValue(row: number, col: number, value: unknown) {
    await expect.poll(() => this.cellValue(row, col)).toEqual(value);
  }
  async expectBold(row: number, col: number, bold = true) {
    await expect.poll(async () => !!(await this.cellStyle(row, col)).bold).toBe(bold);
  }

  toolbar(name: string | RegExp) {
    return this.page.getByRole('button', { name });
  }
}

export function colLabel(col: number): string {
  let n = col;
  let s = '';
  do {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return s;
}
