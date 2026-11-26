// Suggesting mode: what the user does is recorded as tracked changes instead of being done.
//
// A plugin looks at each transaction after it was applied and puts the document into the state a suggestion would have
// left it in, in a transaction of its own that joins the user's in the undo history:
//   - text that was inserted gets an insertion mark (a paragraph break gets an inserted paragraph mark);
//   - text that was deleted is put back with a deletion mark (a paragraph break that was joined away is put back and marked
//     deleted), except text the same author inserted themselves, which is simply gone, as in Word;
//   - a change of character formatting records what the formatting was (`formatChange`), and a change of a paragraph's
//     formatting records the paragraph's old formatting.
// What cannot be recorded (a list, a table, an image or other structure inserted or deleted, wrapping blocks in a list,
// turning text into a list…) is left as the user made it, and the host is told with a notice. Changes from collaborators and
// undo and redo are never touched.

import { Fragment, Slice, type Mark, type Node as PmNode } from 'prosemirror-model';
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { AddMarkStep, AttrStep, Mapping, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, type Step } from 'prosemirror-transform';
import type { RevisionMeta } from '../types';
import { docsSchema } from './schema';
import { newRevisionId } from './revision-edit';
import { readBlockRevision, withBlockRevision, withBlockRevisionKind } from './block-revision';
import { paragraphFormatOf, runFormatOf } from './sync';

export const suggestingKey = new PluginKey('suggesting');

export interface SuggestingOptions {
  /** Whether the editor is in Suggesting mode right now. */
  isActive: () => boolean;
  /** Who is suggesting (the author written on the changes). */
  author: () => string;
  /** The time written on the changes (ISO 8601); default: now. */
  now?: () => string;
  /** The id of a new change; default: random. */
  newId?: () => string;
  /** Called with what had to be done directly instead of as a suggestion. */
  onNotice?: (notices: string[]) => void;
}

/** What a transaction must not be touched for: changes from collaborators, undo and redo, and the plugin's own. */
const SKIPPED_META = ['y-sync$', 'history$', 'y-undo$', 'suggesting$'];

const TEXTBLOCKS = new Set(['paragraph', 'heading']);
/** The attributes that hold the changes of a paragraph mark or row: setting them is not an edit to suggest. */
const BLOCK_REVISION_ATTR_NAMES = new Set(['revisionInsertion', 'revisionDeletion', 'revisionFormat']);
const NOT_FORMATTING = new Set(['insertion', 'deletion', 'formatChange', 'comment', 'link']);

const { insertion: insertionType, deletion: deletionType, formatChange: formatChangeType } = docsSchema.marks;

interface Context {
  out: Transaction;
  author: string;
  now: () => string;
  newId: () => string;
  notices: Set<string>;
  /** The `previous` of the formatting changes written in this call, so an earlier step of the same transaction can replace it. */
  recorded: WeakSet<object>;
}

type StepFields = { from: number; to: number; slice: Slice };

/** The meta of a change by this author that touches `pos` (the node before or after it), so a run of typing or backspacing is one change. */
function neighbour(ctx: Context, pos: number, type: typeof insertionType | typeof deletionType | typeof formatChangeType): RevisionMeta | undefined {
  const $pos = ctx.out.doc.resolve(pos);
  for (const node of [$pos.nodeBefore, $pos.nodeAfter]) {
    const mark = node?.marks.find((m) => m.type === type && m.attrs.author === ctx.author);
    if (mark) return { id: String(mark.attrs.revisionId), author: ctx.author, ...(mark.attrs.date && { date: String(mark.attrs.date) }) };
  }
  return undefined;
}

const freshMeta = (ctx: Context): RevisionMeta => ({ id: ctx.newId(), author: ctx.author, date: ctx.now() });

const markOf = (type: typeof insertionType | typeof deletionType | typeof formatChangeType, meta: RevisionMeta, extra: Record<string, unknown> = {}): Mark =>
  type.create({ revisionId: meta.id, author: meta.author, date: meta.date ?? null, ...extra });

/**
 * True when a slice holds only text (a line break counts, it is kept as text) and paragraphs and headings: what a suggestion
 * can record. An inline image, field, bookmark, note reference or content control is not: a mark on such a node is not kept
 * by the collaboration binding (marks sync on text only), so a collaborator would not see the change.
 */
function isRecordable(content: Fragment): boolean {
  let ok = true;
  content.descendants((node) => {
    if (!(node.isText || node.type === docsSchema.nodes.hard_break || TEXTBLOCKS.has(node.type.name))) ok = false;
    return ok;
  });
  return ok;
}

/** The paragraphs and headings whose closing token (the paragraph mark) is inside [from, to). */
function paragraphMarksIn(doc: PmNode, from: number, to: number): Array<{ pos: number; node: PmNode }> {
  const found: Array<{ pos: number; node: PmNode }> = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (TEXTBLOCKS.has(node.type.name)) {
      const closing = pos + node.nodeSize - 1;
      if (closing >= from && closing < to) found.push({ pos, node });
      return false;
    }
    return true;
  });
  return found;
}

/** Sets one kind of change on the paragraph mark of the node at `pos` (its own attribute, so other people's records stay). */
function setMarkRecord(ctx: Context, pos: number, kind: 'insertion' | 'deletion' | 'formatChange', value: unknown): void {
  const node = ctx.out.doc.nodeAt(pos);
  if (!node) return;
  ctx.out.setNodeMarkup(pos, undefined, withBlockRevisionKind(node.attrs, kind, value));
}

/** What was inserted at [from, to) of the document `ctx.out` holds is marked as inserted. */
function markInserted(ctx: Context, from: number, to: number, pasted: boolean): void {
  const { out } = ctx;
  // Marks the inserted content carried with it (pasted from elsewhere in a document with changes) are not this author's.
  for (const type of [insertionType, deletionType, formatChangeType]) out.removeMark(from, to, type);
  const meta = neighbour(ctx, from, insertionType) ?? neighbour(ctx, to, insertionType) ?? freshMeta(ctx);
  out.addMark(from, to, markOf(insertionType, meta));
  for (const { pos } of paragraphMarksIn(out.doc, from, to)) {
    // The paragraph mark that was just made did not exist before, so no one can have deleted it: a copy of the deletion record
    // of the mark the paragraph was split at (Enter) is not carried over. The old mark stays with the second half.
    setMarkRecord(ctx, pos, 'deletion', null);
    setMarkRecord(ctx, pos, 'insertion', meta);
    // A paragraph pasted from elsewhere brings the record of a formatting change someone else made there; it is not a change of
    // this document. (An Enter copies the record of the paragraph it splits to both halves, which stays: it is the same change.)
    if (pasted) setMarkRecord(ctx, pos, 'formatChange', null);
  }
  // A line break is kept in the text of a tracked run: the Y binding keeps marks on text only, not on a break node.
  const breaks: number[] = [];
  out.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type === docsSchema.nodes.hard_break) breaks.push(pos);
    return true;
  });
  for (const pos of breaks.reverse()) out.replaceWith(pos, pos + 1, docsSchema.text('\n', out.doc.nodeAt(pos)!.marks));
}

/** The fragment with everything in it marked deleted; text this author inserted is left out. */
function markDeleted(ctx: Context, content: Fragment, meta: RevisionMeta): Fragment {
  const nodes: PmNode[] = [];
  content.forEach((node) => {
    if (node.isInline) {
      if (node.marks.some((m) => m.type === insertionType && m.attrs.author === ctx.author)) return;
      // A deleted line break is kept as text too (see `markInserted`).
      const kept = node.type === docsSchema.nodes.hard_break ? docsSchema.text('\n', node.marks) : node;
      nodes.push(kept.marks.some((m) => m.type === deletionType) ? kept : kept.mark(markOf(deletionType, meta).addToSet(kept.marks)));
    } else {
      nodes.push(node.copy(markDeleted(ctx, node.content, meta)));
    }
  });
  return Fragment.from(nodes);
}

/** Handle one replace step: what it inserted is marked, what it deleted is put back and marked. Returns where the caret goes for a deletion. */
function replaceStep(ctx: Context, step: ReplaceStep, docBefore: PmNode, after: Mapping, caret: { forward: boolean }): number | null {
  const { out } = ctx;
  const { from, to, slice } = step as unknown as StepFields;
  const start = out.mapping.map(after.map(from, 1), 1);
  const end = Math.max(start, out.mapping.map(Math.max(after.map(from, 1), after.map(from + slice.size, -1)), -1));
  let caretAt: number | null = null;

  if (slice.size > 0) {
    // Structure that cannot be recorded stays as it was inserted, unmarked, as the notice says.
    const split = slice.content.childCount === 2 && slice.size === 2 && slice.openStart === 1 && slice.openEnd === 1;
    if (isRecordable(slice.content)) markInserted(ctx, start, end, !split);
    else ctx.notices.add('A list, table, image or other structure that was inserted could not be suggested and was inserted directly.');
  }

  if (to > from) {
    const deleted = docBefore.slice(from, to);
    if (!isRecordable(deleted.content)) {
      ctx.notices.add('A list, table, image or other structure that was deleted could not be suggested and was deleted directly.');
      return caretAt;
    }
    // A paragraph mark this author inserted and now deletes is simply gone (Enter, then Backspace).
    const lone = deleted.content.childCount === 2 && deleted.openStart === 1 && deleted.openEnd === 1 && deleted.size === 2 && deleted.content.child(0).childCount === 0 && deleted.content.child(1).childCount === 0;
    const firstParent = docBefore.resolve(from).parent;
    const own = readBlockRevision(firstParent.attrs)?.insertion;
    if (lone && own && own.author === ctx.author && TEXTBLOCKS.has(firstParent.type.name)) {
      // The join keeps the first paragraph's attributes, but the paragraph that results ends with the second one's mark: it takes
      // that mark's records (inserted, deleted) and keeps the first one's formatting change, as accepting a deleted mark does.
      const secondParent = docBefore.resolve(to).parent;
      const first = readBlockRevision(firstParent.attrs) ?? {};
      const second = readBlockRevision(secondParent.attrs) ?? {};
      const merged = out.doc.resolve(start);
      ctx.out.setNodeMarkup(
        merged.before(),
        undefined,
        withBlockRevision(merged.parent.attrs, {
          ...(second.insertion && { insertion: second.insertion }),
          ...(second.deletion && { deletion: second.deletion }),
          ...(first.formatChange && { formatChange: first.formatChange }),
        })
      );
      return start;
    }
    const meta = neighbour(ctx, start, deletionType) ?? freshMeta(ctx);
    const content = markDeleted(ctx, deleted.content, meta);
    if (content.size === 0) return start;
    const before = out.doc.content.size;
    try {
      out.replace(start, start, new Slice(content, deleted.openStart, deleted.openEnd));
    } catch {
      ctx.notices.add('Deleted text could not be put back as a suggestion and was deleted directly.');
      return start;
    }
    const size = out.doc.content.size - before;
    if (size === 0) {
      ctx.notices.add('Deleted text could not be put back as a suggestion and was deleted directly.');
      return start;
    }
    for (const { pos, node } of paragraphMarksIn(out.doc, start, start + size)) {
      // A mark already deleted (by anyone) keeps its record: deleting it again is not a new change.
      if (!node.attrs.revisionDeletion) setMarkRecord(ctx, pos, 'deletion', meta);
    }
    // The caret stays after what was typed over the deleted text; with nothing typed it is after text deleted forward, before text backspaced.
    if (slice.size === 0) caretAt = caret.forward ? start + size : start;
  }
  return caretAt;
}

/** A change of character formatting records what the formatting was, on the text it changed. */
function markStep(ctx: Context, step: AddMarkStep | RemoveMarkStep, docBefore: PmNode, after: Mapping): void {
  const { out } = ctx;
  const { from, to, mark } = step as unknown as { from: number; to: number; mark: Mark };
  if (NOT_FORMATTING.has(mark.type.name)) return;
  const start = out.mapping.map(after.map(from, 1), 1);
  const end = out.mapping.map(after.map(to, -1), -1);
  const meta = neighbour(ctx, start, formatChangeType) ?? neighbour(ctx, end, formatChangeType) ?? freshMeta(ctx);
  docBefore.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const a = out.mapping.map(after.map(Math.max(pos, from), 1), 1);
    const b = out.mapping.map(after.map(Math.min(pos + node.nodeSize, to), -1), -1);
    if (a >= b) return false;
    const previous = runFormatOf(node.marks);
    out.doc.nodesBetween(a, b, (now, nowPos) => {
      if (!now.isInline) return true;
      const na = Math.max(nowPos, a);
      const nb = Math.min(nowPos + now.nodeSize, b);
      // A mark on an image, field, bookmark or other inline node would not sync (marks travel on text only): the formatting stays, unrecorded.
      if (!now.isText) {
        ctx.notices.add('The formatting of an image, field or other object in the selection cannot be suggested and was applied directly.');
        return false;
      }
      if (now.marks.some((m) => m.type === insertionType && m.attrs.author === ctx.author)) return false;
      const existing = now.marks.find((m) => m.type === formatChangeType);
      if (existing && !ctx.recorded.has(existing.attrs.previous as object)) {
        // Formatting that is back to what it was is no change.
        if (existing.attrs.author === ctx.author && sameFormat(existing.attrs.previous, runFormatOf(now.marks))) out.removeMark(na, nb, formatChangeType);
        return false;
      }
      // Steps are handled from the last to the first: when a later step of this transaction recorded the text already, the
      // formatting before this earlier step is the older, so it replaces that record (the first record is what Reject restores).
      if (existing && sameFormat(previous, runFormatOf(now.marks))) {
        out.removeMark(na, nb, formatChangeType);
        return false;
      }
      const attrs = { previous };
      ctx.recorded.add(previous);
      out.addMark(na, nb, markOf(formatChangeType, existing ? { id: String(existing.attrs.revisionId), author: ctx.author, ...(existing.attrs.date && { date: String(existing.attrs.date) }) } : meta, attrs));
      return false;
    });
    return false;
  });
}

const FORMAT_KEYS = ['bold', 'italic', 'underline', 'strikethrough', 'superscript', 'subscript', 'smallCaps', 'allCaps', 'hidden', 'color', 'backgroundColor', 'fontSize', 'fontFamily', 'fontFamilyEastAsia', 'fontFamilyComplex', 'namedStyleId'] as const;

function sameFormat(a: unknown, b: unknown): boolean {
  const x = (a ?? {}) as Record<string, unknown>;
  const y = (b ?? {}) as Record<string, unknown>;
  return FORMAT_KEYS.every((key) => (x[key] || undefined) === (y[key] || undefined));
}

/** What a change to the type or attributes of a node that is not a paragraph or heading says (made directly). */
const ATTRIBUTE_NOTICE = 'A change to the properties of a list, table, image, form control or other structure cannot be suggested and was made directly.';

/**
 * A paragraph whose formatting or type was changed records the formatting it had. Nothing is recorded, on purpose and
 * without a notice, when the change is already recorded (the first record keeps the formatting from before), when the
 * paragraph is the author's own insertion (it has no earlier formatting), or when the formatting did not change.
 */
function paragraphChanged(ctx: Context, before: PmNode, pos: number): void {
  const node = ctx.out.doc.nodeAt(pos);
  if (!node || !TEXTBLOCKS.has(node.type.name) || !TEXTBLOCKS.has(before.type.name)) return;
  const revision = readBlockRevision(node.attrs) ?? {};
  if (revision.insertion?.author === ctx.author) return;
  const previous = paragraphFormatOf(before);
  const current = paragraphFormatOf(node);
  const same = JSON.stringify(previous) === JSON.stringify(current);
  // An attribute a record cannot carry (the indent level of Increase / Decrease indent) was changed: say it was made directly.
  if (changedAttributes(before, node)) ctx.notices.add(UNRECORDED_NOTICE);
  if (revision.formatChange || same) return;
  setMarkRecord(ctx, pos, 'formatChange', { ...freshMeta(ctx), previous });
}

/** True when an attribute of a paragraph that neither the change records nor the formatting a record carries differs between two versions of it. */
function changedAttributes(before: PmNode, after: PmNode): boolean {
  const covered = new Set(Object.keys(paragraphFormatOf(before)).concat(Object.keys(paragraphFormatOf(after)), ['level']));
  for (const key of new Set([...Object.keys(before.attrs), ...Object.keys(after.attrs)])) {
    if (key.startsWith('revision') || covered.has(key)) continue;
    if (JSON.stringify(before.attrs[key] ?? null) !== JSON.stringify(after.attrs[key] ?? null)) return true;
  }
  return false;
}

/** A change of a paragraph attribute that a formatting change does not record. */
const UNRECORDED_NOTICE = "A change of a paragraph's indent cannot be suggested and was made directly.";

/**
 * Tell the host what was done directly. The plugin runs inside `state.apply`, where a host must not set its own state, and
 * a transaction applied only to look at it would tell twice, so the notice goes out after the current task step.
 */
function notify(options: SuggestingOptions, notices: Set<string>): void {
  if (notices.size === 0 || !options.onNotice) return;
  const list = [...notices];
  queueMicrotask(() => options.onNotice?.(list));
}

/**
 * The changes a suggestion records, from the transactions the plugin was given; null when there is nothing to do. A
 * transaction another plugin of the editor appends is recorded like a user's (only the meta keys above are skipped), so a host
 * plugin that rewrites the document in the background should be given one of them.
 */
export function suggestChanges(trs: readonly Transaction[], oldState: EditorState, newState: EditorState, options: SuggestingOptions): Transaction | null {
  if (!options.isActive()) return null;
  if (trs.some((tr) => SKIPPED_META.some((key) => tr.getMeta(key) !== undefined) || tr.getMeta('addToHistory') === false)) return null;
  const steps: Array<{ step: Step; before: PmNode }> = [];
  for (const tr of trs) tr.steps.forEach((step, i) => steps.push({ step, before: tr.docs[i] }));
  if (steps.length === 0) return null;

  const ctx: Context = {
    out: newState.tr,
    author: options.author(),
    now: options.now ?? (() => new Date().toISOString()),
    newId: options.newId ?? newRevisionId,
    notices: new Set(),
    recorded: new WeakSet(),
  };
  const all = new Mapping(steps.map(({ step }) => step.getMap()));
  const old = oldState.selection;
  // Backspace deletes the text before the caret, or the paragraph mark before it (a join: from one before the caret to one after it).
  const first = steps[0].step as unknown as StepFields;
  const backspace = old.empty && steps.length === 1 && first.from < first.to && (first.to === old.head || (first.from === old.head - 1 && first.to === old.head + 1));
  const caret = { forward: !backspace };
  let caretAt: number | null = null;
  const onlyOne = steps.length === 1;

  // From the last step to the first, so the positions of the earlier ones are not moved by what is done for the later ones.
  for (let i = steps.length - 1; i >= 0; i--) {
    const { step, before } = steps[i];
    const after = all.slice(i + 1);
    if (step instanceof ReplaceStep) {
      const at = replaceStep(ctx, step, before, after, caret);
      if (onlyOne) caretAt = at;
    } else if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      markStep(ctx, step, before, after);
    } else if (step instanceof AttrStep) {
      const { pos, attr } = step as unknown as { pos: number; attr: string };
      const node = before.nodeAt(pos);
      if (node && !BLOCK_REVISION_ATTR_NAMES.has(attr)) {
        if (TEXTBLOCKS.has(node.type.name)) paragraphChanged(ctx, node, ctx.out.mapping.map(after.map(pos, -1), -1));
        else ctx.notices.add(ATTRIBUTE_NOTICE);
      }
    } else if (step instanceof ReplaceAroundStep) {
      const { from, to, gapFrom, gapTo, slice } = step as unknown as StepFields & { gapFrom: number; gapTo: number };
      const node = before.nodeAt(from);
      // Setting a node's type or attributes keeps its content, and is how a paragraph becomes a heading or is aligned.
      if (node && gapFrom === from + 1 && gapTo === to - 1 && slice.size === 2) {
        if (TEXTBLOCKS.has(node.type.name)) paragraphChanged(ctx, node, ctx.out.mapping.map(after.map(from, -1), -1));
        else ctx.notices.add(ATTRIBUTE_NOTICE);
      } else ctx.notices.add('A change to the structure of the document (a list, a quote, a table) cannot be suggested and was made directly.');
    } else {
      ctx.notices.add('A change that cannot be suggested was made directly.');
    }
  }

  const out = ctx.out;
  if (!out.docChanged) {
    notify(options, ctx.notices);
    return null;
  }
  // The selection stays as it was mapped through what was done (a range formatted stays selected); only a deletion, which
  // puts the text back, says where the caret goes.
  if (caretAt !== null) out.setSelection(TextSelection.near(out.doc.resolve(Math.min(caretAt, out.doc.content.size))));
  out.setMeta('suggesting$', true);
  notify(options, ctx.notices);
  return out;
}

/** The plugin: in Suggesting mode, what the user does becomes tracked changes. */
export function suggestingPlugin(options: SuggestingOptions): Plugin {
  return new Plugin({
    key: suggestingKey,
    appendTransaction: (trs, oldState, newState) => suggestChanges(trs, oldState, newState, options),
  });
}
