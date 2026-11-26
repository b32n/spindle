/**
 * How a run with a tracked change is drawn, as plain data the painter applies to the run's element (so it can be tested
 * without a DOM): inserted text underlined, deleted text struck through, both in the colour of the author who made the
 * change (the last one to change it when it was inserted and deleted).
 */

import { DEFAULT_REVISION_COLOR, type Run } from './flow-blocks';
import { safeCssColor } from './sanitize';

export interface RevisionDecoration {
  color: string;
  textDecoration: string;
  /** `data-*` attributes, by their dataset names. */
  data: Record<string, string>;
  title: string;
}

export function revisionDecoration(run: Run): RevisionDecoration | null {
  if ((run.kind !== 'text' && run.kind !== 'link') || !run.revision) return null;
  const { insertion, deletion } = run.revision;
  const author = (deletion ?? insertion)?.author;
  if (author === undefined) return null;
  const lines = new Set<string>();
  if (run.kind === 'link' ? !run.plain : run.underline) lines.add('underline');
  if (run.kind === 'text' && run.strikethrough) lines.add('line-through');
  if (insertion) lines.add('underline');
  if (deletion) lines.add('line-through');
  const data: Record<string, string> = { revisionAuthor: author };
  if (insertion) data.revisionInsertion = insertion.id;
  if (deletion) data.revisionDeletion = deletion.id;
  return {
    color: safeCssColor((deletion ?? insertion)?.color) || DEFAULT_REVISION_COLOR,
    textDecoration: [...lines].join(' '),
    data,
    title: [insertion && `Inserted by ${insertion.author}`, deletion && `Deleted by ${deletion.author}`].filter(Boolean).join('; '),
  };
}

/** Adds to an element's tooltip, keeping what it has (a link's title, a content control's title). */
export function addTitle(el: { title: string }, text: string): void {
  el.title = el.title ? `${el.title}\n${text}` : text;
}
