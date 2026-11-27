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

/** Longest piece of the file's own text (a sheet name, a cell value, a reference, a part name) that a warning quotes. */
export const MAX_QUOTED_CHARS = 64;

/**
 * The first `max` characters of `text` (fewer when the cut would leave half a surrogate pair), as a copy: a slice of a long text
 * keeps the whole text alive, and this is for what is kept (a sheet name) or quoted (a warning) of a file's text of any length.
 */
export function truncate(text: string, max: number): string {
  if (text.length <= max) return text;
  const last = text.charCodeAt(max - 1);
  return text.slice(0, last >= 0xd800 && last <= 0xdbff ? max - 1 : max).split('').join('');
}

/**
 * `text` as a warning may quote it: at most `max` characters, the last of them an ellipsis when it was cut. The text is the file's, so
 * it can be of any length, and a warning that repeats it must not grow with it.
 */
export function abbreviate(text: string, max: number = MAX_QUOTED_CHARS): string {
  return text.length <= max ? text : `${truncate(text, max - 1)}…`;
}
