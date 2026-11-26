/**
 * How a run with a tracked change is drawn, as plain data the painter applies to the run's element (so it can be tested
 * without a DOM), the way Word's "All Markup" does: inserted text underlined, deleted text struck through, both in the
 * colour of the author who made the change (the last one to change it when it was inserted and deleted); the two ends of
 * a move with a double line; a change of formatting alone with a dotted underline and the text's own colour (a run that
 * is underlined already shows that underline dotted).
 */

import { DEFAULT_REVISION_COLOR, type Run, type TableRow } from './flow-blocks';
import { safeCssColor } from './sanitize';

export interface RevisionDecoration {
  /** Absent for a change of formatting alone, which keeps the text's colour. */
  color?: string;
  textDecoration: string;
  textDecorationColor: string;
  textDecorationStyle?: 'double' | 'dotted';
  /** `data-*` attributes, by their dataset names. */
  data: Record<string, string>;
  title: string;
}

export function revisionDecoration(run: Run): RevisionDecoration | null {
  if ((run.kind !== 'text' && run.kind !== 'link') || !run.revision) return null;
  const { insertion, deletion, formatChange } = run.revision;
  const shown = deletion ?? insertion ?? formatChange;
  if (!shown) return null;
  const lines = new Set<string>();
  if (run.kind === 'link' ? !run.plain : run.underline) lines.add('underline');
  if (run.kind === 'text' && run.strikethrough) lines.add('line-through');
  if (insertion) lines.add('underline');
  if (deletion) lines.add('line-through');
  if (!insertion && !deletion) lines.add('underline');
  const color = safeCssColor(shown.color) || DEFAULT_REVISION_COLOR;
  const move = deletion?.moveId ?? insertion?.moveId;
  const data: Record<string, string> = { revisionAuthor: shown.author };
  if (insertion) data.revisionInsertion = insertion.id;
  if (deletion) data.revisionDeletion = deletion.id;
  if (formatChange) {
    data.revisionFormat = formatChange.id;
    // The author of the formatting change, when another change is what is shown.
    if (insertion || deletion) data.revisionFormatAuthor = formatChange.author;
  }
  if (move) data.revisionMove = move;
  return {
    ...((insertion || deletion) && { color }),
    textDecoration: [...lines].join(' '),
    textDecorationColor: color,
    ...(move ? { textDecorationStyle: 'double' as const } : !insertion && !deletion ? { textDecorationStyle: 'dotted' as const } : {}),
    data,
    title: [
      insertion && `${insertion.moveId ? 'Moved to here' : 'Inserted'} by ${insertion.author}`,
      deletion && `${deletion.moveId ? 'Moved from here' : 'Deleted'} by ${deletion.author}`,
      formatChange && `Formatted by ${formatChange.author}`,
    ]
      .filter(Boolean)
      .join('; '),
  };
}

/** The tint of the cells of a row inserted or deleted as a tracked change, in the author's colour (hex colours only: the tint adds an alpha to it). */
export function rowRevisionTint(revision: TableRow['revision']): { backgroundImage: string; data: Record<string, string> } | null {
  const shown = revision?.deletion ?? revision?.insertion;
  const color = shown ? safeCssColor(shown.color) : undefined;
  if (!shown || !color || !/^#[0-9a-f]{6}$/i.test(color)) return null;
  return {
    backgroundImage: `linear-gradient(${color}26, ${color}26)`,
    data: { revisionRow: revision?.deletion ? 'deletion' : 'insertion', revisionAuthor: shown.author },
  };
}

/** Adds to an element's tooltip, keeping what it has (a link's title, a content control's title). */
export function addTitle(el: { title: string }, text: string): void {
  el.title = el.title ? `${el.title}\n${text}` : text;
}
