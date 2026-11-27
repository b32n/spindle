// ---------------------------------------------------------------------------
// Warnings about single items of a part (a hyperlink without its relationship, a sheet that repeats a part) must not grow with
// the file: a hostile part names two million of them. The first few are reported as they are, the rest are counted in one closing
// message, as the sheet reader does for cells (`MAX_CELL_WARNINGS`).
// ---------------------------------------------------------------------------

/** Most warnings of one kind that are given in full; the rest are counted. */
export const MAX_ITEM_WARNINGS = 20;

export class CappedWarnings {
  private count = 0;

  /** `warn` receives what is reported; `what` completes the closing message ("N more ${what} were left out."). */
  constructor(
    private readonly warn: (message: string) => void,
    private readonly what: string,
    private readonly cap: number = MAX_ITEM_WARNINGS
  ) {}

  add(message: string): void {
    if (this.count++ < this.cap) this.warn(message);
  }

  /** The closing message, when anything was left out. Safe to call once, at the end. */
  close(): void {
    if (this.count > this.cap) this.warn(`${this.count - this.cap} more ${this.what} were left out.`);
  }
}
