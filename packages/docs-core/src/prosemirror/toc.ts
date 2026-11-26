import type { Node as PmNode, Schema } from 'prosemirror-model';
import { NodeSelection, TextSelection, type Command } from 'prosemirror-state';

// ---------------------------------------------------------------------------
// Table of contents (`toc`): a block holding one entry per heading, with a link to it and its page number. The editor
// writes the entries when the table is inserted or updated; between those it does not touch them, so a table imported
// from a file keeps the numbers it came with until it is updated.
// ---------------------------------------------------------------------------

/** What Word writes for a table of contents of the first three heading levels, with links. */
export const DEFAULT_TOC_INSTRUCTION = 'TOC \\o "1-3" \\h \\z \\u';

/** Shown when a table of contents has nothing to list. */
export const EMPTY_TOC_TEXT = 'No table of contents entries found.';

/** Indent of an entry per heading level below the first, in pixels. */
const ENTRY_INDENT_PX = 24;

/** What a table of contents instruction asks for (the parts of it that are acted on). */
export interface TocSettings {
  /** The heading levels listed, from the first to the last (headings go from 1 to 6). */
  fromLevel: number;
  toLevel: number;
  /** Entries link to their headings (`\h`). */
  hyperlinks: boolean;
  /** Entries show page numbers (no `\n`). */
  pageNumbers: boolean;
  /** Switches the editor does not act on (they stay in the instruction): `\t` styles, `\b` bookmarks, `\f` entries, ... */
  ignored: string[];
}

/** Read the switches of a table of contents instruction. Heading levels default to 1 to 3. */
export function parseTocInstruction(instruction: string): TocSettings {
  const levels = /\\o\s+"\s*(\d)\s*-\s*(\d)\s*"/i.exec(instruction);
  const clamp = (n: number) => Math.min(6, Math.max(1, n));
  const from = levels ? clamp(Number(levels[1])) : 1;
  const to = levels ? clamp(Number(levels[2])) : 3;
  const ignored = [...instruction.matchAll(/\\([a-z])\b/gi)].map((m) => m[1].toLowerCase()).filter((c) => !['o', 'h', 'z', 'u', 'n'].includes(c));
  return {
    fromLevel: Math.min(from, to),
    toLevel: Math.max(from, to),
    hyperlinks: /\\h\b/i.test(instruction),
    pageNumbers: !/\\n\b/i.test(instruction),
    ignored: [...new Set(ignored)],
  };
}

/** The `_Toc` bookmark a heading already has, if any. */
function tocBookmarkOf(heading: PmNode): string | null {
  let found: string | null = null;
  heading.forEach((child) => {
    if (!found && child.type.name === 'bookmark' && typeof child.attrs.name === 'string' && child.attrs.name.startsWith('_Toc')) found = child.attrs.name;
  });
  return found;
}

interface Operation {
  pos: number;
  /** What to do at `pos`: insert a node, or replace the content of the table of contents there. */
  apply: (tr: import('prosemirror-state').Transaction) => void;
}

/** Everything the headings, bookmarks and tables of contents of a document need for the tables to be written. */
function planEntries(
  doc: PmNode,
  schema: Schema,
  pageLabelAt: (pos: number) => string | undefined,
  tables: Array<{ pos: number; instruction: string; node?: PmNode }>
): { operations: Operation[]; entries: PmNode[][] } {
  const headings: Array<{ node: PmNode; pos: number }> = [];
  const names = new Set<string>();
  doc.descendants((node, pos) => {
    if (node.type.name === 'toc') return false;
    if (node.type.name === 'heading') headings.push({ node, pos });
    if (node.type.name === 'bookmark') names.add(String(node.attrs.name));
    return true;
  });

  const operations: Operation[] = [];
  const bookmarkOf = new Map<number, string>();
  let counter = 0;
  const nameFor = (heading: { node: PmNode; pos: number }): string => {
    const known = bookmarkOf.get(heading.pos) ?? tocBookmarkOf(heading.node);
    if (known) {
      bookmarkOf.set(heading.pos, known);
      return known;
    }
    let name: string;
    do name = `_Toc${100000000 + ++counter}`;
    while (names.has(name));
    names.add(name);
    bookmarkOf.set(heading.pos, name);
    const at = heading.pos + 1;
    operations.push({ pos: at, apply: (tr) => tr.insert(at, schema.nodes.bookmark.create({ name })) });
    return name;
  };

  const entries = tables.map((table) => {
    const settings = parseTocInstruction(table.instruction);
    const rows: PmNode[] = [];
    for (const heading of headings) {
      const level = Number(heading.node.attrs.level) || 1;
      const text = finalText(heading.node).trim();
      if (level < settings.fromLevel || level > settings.toLevel || !text) continue;
      const content: PmNode[] = [];
      const label = settings.pageNumbers ? pageLabelAt(heading.pos) : undefined;
      const link = settings.hyperlinks ? [schema.marks.link.create({ href: `#${nameFor(heading)}` })] : [];
      content.push(schema.text(text, link));
      if (label) content.push(schema.text('\t'), schema.text(label));
      rows.push(
        schema.nodes.paragraph.create(
          { namedStyleId: `TOC${level}`, ...(level > 1 && { leftIndent: (level - 1) * ENTRY_INDENT_PX }) },
          content
        )
      );
    }
    if (rows.length === 0) rows.push(schema.nodes.paragraph.create(null, schema.text(EMPTY_TOC_TEXT)));
    return rows;
  });
  return { operations, entries };
}

/** The text of a node as it reads once every tracked deletion is accepted (a table of contents lists the final text). */
function finalText(node: PmNode): string {
  let text = '';
  node.descendants((child) => {
    if (child.isText && !child.marks.some((mark) => mark.type.name === 'deletion')) text += child.text;
    return true;
  });
  return text;
}

/** The tables of contents of a document, with their positions (the ones inside another are not looked for: they cannot nest). */
function tablesOf(doc: PmNode): Array<{ pos: number; node: PmNode }> {
  const tables: Array<{ pos: number; node: PmNode }> = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'toc') {
      tables.push({ pos, node });
      return false;
    }
    return true;
  });
  return tables;
}

/**
 * Write the entries of every table of contents of the document from its headings (the ones in `toc.fromLevel` to
 * `toLevel` of its instruction), giving each heading a `_Toc` bookmark to link to when it has none. `pageLabelAt` says the
 * page number to show for the heading at a position of the document (nothing: no number). One transaction, so one undo.
 */
export function updateTableOfContents(pageLabelAt: (pos: number) => string | undefined = () => undefined): Command {
  return (state, dispatch) => {
    const tables = tablesOf(state.doc);
    if (tables.length === 0) return false;
    if (dispatch) {
      const { operations, entries } = planEntries(
        state.doc,
        state.schema,
        pageLabelAt,
        tables.map((t) => ({ pos: t.pos, instruction: String(t.node.attrs.instruction ?? DEFAULT_TOC_INSTRUCTION) }))
      );
      tables.forEach((table, i) => {
        operations.push({ pos: table.pos + 1, apply: (tr) => tr.replaceWith(table.pos + 1, table.pos + 1 + table.node.content.size, entries[i]) });
      });
      const tr = state.tr;
      // From the end of the document to its start, so the positions of the steps still to apply are not moved.
      for (const op of operations.sort((a, b) => b.pos - a.pos)) op.apply(tr);
      dispatch(tr);
    }
    return true;
  };
}

/** Where a table of contents goes: after the top-level block the selection ends in. */
function insertionPoint(doc: PmNode, to: number): number {
  const $to = doc.resolve(to);
  return $to.depth >= 1 ? $to.after(1) : to;
}

/**
 * Insert a table of contents after the block the selection ends in, with its entries written (see `updateTableOfContents`).
 * Not applicable where a block of that kind cannot go.
 */
export function insertTableOfContents(pageLabelAt: (pos: number) => string | undefined = () => undefined, instruction: string = DEFAULT_TOC_INSTRUCTION): Command {
  return (state, dispatch) => {
    const type = state.schema.nodes.toc;
    const at = insertionPoint(state.doc, state.selection.to);
    if (!type || !state.doc.resolve(at).parent.canReplaceWith(state.doc.resolve(at).index(), state.doc.resolve(at).index(), type)) return false;
    if (dispatch) {
      const { operations, entries } = planEntries(state.doc, state.schema, pageLabelAt, [{ pos: at, instruction }]);
      const node = type.create({ instruction }, entries[0]);
      // The table is itself an operation, at its own position (it has no heading to be listed among).
      operations.push({ pos: at, apply: (tr) => tr.insert(at, node) });
      const tr = state.tr;
      for (const op of operations.sort((a, b) => b.pos - a.pos)) op.apply(tr);
      const after = tr.mapping.map(at, 1);
      tr.setSelection(TextSelection.near(tr.doc.resolve(Math.min(after, tr.doc.content.size)), 1));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** The table of contents that is selected, that holds the selection, or else the nearest one (before or after it). */
function tableAt(state: { doc: PmNode; selection: import('prosemirror-state').Selection }): { pos: number; node: PmNode } | null {
  const { selection, doc } = state;
  if (selection instanceof NodeSelection && selection.node.type.name === 'toc') return { pos: selection.from, node: selection.node };
  const tables = tablesOf(doc);
  const holding = tables.find((t) => selection.from >= t.pos && selection.from <= t.pos + t.node.nodeSize);
  if (holding) return holding;
  const before = [...tables].reverse().find((t) => t.pos + t.node.nodeSize < selection.from);
  const after = tables.find((t) => t.pos > selection.from);
  if (before && after) return selection.from - (before.pos + before.node.nodeSize) <= after.pos - selection.from ? before : after;
  return before ?? after ?? null;
}

/** Delete the table of contents at the selection, or the nearest one (see `tableAt`). Not applicable when the document has none. */
export const removeTableOfContents: Command = (state, dispatch) => {
  const table = tableAt(state);
  if (!table) return false;
  if (dispatch) dispatch(state.tr.delete(table.pos, table.pos + table.node.nodeSize));
  return true;
};
