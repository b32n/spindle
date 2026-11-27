import type { XmlElement } from '../xml';
import { scanXml, type SaxAttrs } from './sax';

// ---------------------------------------------------------------------------
// Element trees for the workbook-level parts, built from the scanner's events under an element budget. The budget is what keeps
// memory bounded by what is kept: a part of five million `<border>` elements costs `maxPartElements` elements, not five million,
// and is then dropped. (The general `parseXml` of `../xml.ts` builds the whole raw tree of fast-xml-parser first, several times
// the part's size, and is for the loss detector.)
// ---------------------------------------------------------------------------

/** Elements that may still be built; shared by every builder that draws on it. */
export interface ElementBudget {
  left: number;
}

/** Builds the element tree of the events it is given. Elements past the budget are not built (nor are the ones below them); `dropped` says so. */
export class TreeBuilder {
  root: XmlElement | undefined;
  /** An element was left out because the budget ran out. */
  dropped = false;
  private readonly stack: XmlElement[] = [];
  private skipped = 0;

  constructor(private readonly budget: ElementBudget) {}

  /** Whether an element is open (built or skipped): the events belong to this builder until it is not. */
  get open(): boolean {
    return this.stack.length > 0 || this.skipped > 0;
  }

  start(name: string, attrs: SaxAttrs): void {
    // a second root (damage the scanner does not report) is skipped like an element over the budget
    if (this.skipped > 0 || (this.stack.length === 0 && this.root !== undefined)) {
      this.skipped++;
      return;
    }
    if (this.budget.left <= 0) {
      this.dropped = true;
      this.skipped++;
      return;
    }
    this.budget.left--;
    // the scanner's attribute object is the element's own (a null-prototype object, so `__proto__` is a name like any other)
    const el: XmlElement = { name, ns: '', attrs: attrs as Record<string, string>, text: '', children: [] };
    const parent = this.stack[this.stack.length - 1];
    if (parent) parent.children.push(el);
    else this.root = el;
    this.stack.push(el);
  }

  end(): void {
    if (this.skipped > 0) this.skipped--;
    else this.stack.pop();
  }

  text(chunk: string): void {
    if (this.skipped === 0 && this.stack.length > 0) this.stack[this.stack.length - 1].text += chunk;
  }
}

export interface ParsedTree {
  /** The root element; `undefined` when the text is not XML, is damaged, or is over the budget (`overLimit`). */
  root: XmlElement | undefined;
  /** More elements than the budget allowed: the part is not read. */
  overLimit: boolean;
}

/**
 * The tree of a part of at most `maxElements` elements. Damaged markup (see `scanXml`), no root element and more elements than
 * allowed all give no root. The scan stops at the first element over the budget, so a hostile part costs the budget, not its size.
 */
export function parseTree(xml: string, maxElements: number): ParsedTree {
  const builder = new TreeBuilder({ left: maxElements });
  const wellFormed = scanXml(xml, {
    open: (name, attrs) => builder.start(name, attrs),
    close: () => builder.end(),
    text: (chunk) => builder.text(chunk),
    halted: () => builder.dropped,
  });
  if (builder.dropped) return { root: undefined, overLimit: true };
  return { root: wellFormed ? builder.root : undefined, overLimit: false };
}
