import type { PptxCommentThread } from '@b32nio/spindle-shared/export/pptx';
import { attrs, esc, XML_DECLARATION } from './writer-xml';

// Comments (plan K13), in the legacy format: `ppt/commentAuthors.xml` and one `ppt/comments/commentN.xml` per slide with threads.
// A reply is a comment with the `p15:parentCm` threading extension, as PowerPoint 2013 writes them. Modern comments are written in
// this format too (the export warnings say so). The inverse of `readThreads` for legacy comments.

export const NS_P15 = 'http://schemas.microsoft.com/office/powerpoint/2012/main';
const THREADING_EXT = '{C676402C-5697-4E1C-873F-D02D1690AC5C}';
/** Legacy positions are in 1/576 inch; a px is 1/96 inch. */
const LEGACY_PER_PX = 6;

/** The authors of a deck's comments, numbered in the order they first comment. */
export class CommentAuthors {
  private readonly list: { name: string; initials: string; last: number }[] = [];
  private readonly byKey = new Map<string, number>();
  /** The id of the author, adding them on first use, and the next comment index they have. */
  next(name: string, initials: string | undefined): { author: number; idx: number } {
    const key = `${name}\u0000${initials ?? ''}`;
    let author = this.byKey.get(key);
    if (author === undefined) {
      author = this.list.length;
      this.list.push({ name, initials: initials ?? '', last: 0 });
      this.byKey.set(key, author);
    }
    return { author, idx: ++this.list[author].last };
  }
  xml(): string {
    const items = this.list.map((a, i) => `<p:cmAuthor${attrs({ id: i, name: a.name, initials: a.initials, lastIdx: a.last, clrIdx: i })}/>`).join('');
    return `${XML_DECLARATION}<p:cmAuthorLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${items}</p:cmAuthorLst>`;
  }
}

/** `dt` of a comment: the UTC time without a zone, the way the files write it (the reader takes a zoneless time as UTC). */
const stamp = (iso: string | undefined): string | undefined => (iso !== undefined ? iso.replace(/Z$/, '') : undefined);

/** `ppt/comments/commentN.xml` for the threads of one slide. */
export function commentsXml(threads: PptxCommentThread[], authors: CommentAuthors): string {
  const items: string[] = [];
  for (const thread of threads) {
    const pos = `<p:pos x="${Math.round((thread.position?.x ?? 0) * LEGACY_PER_PX)}" y="${Math.round((thread.position?.y ?? 0) * LEGACY_PER_PX)}"/>`;
    let root: { author: number; idx: number } | undefined;
    for (const c of thread.comments) {
      const at = authors.next(c.author, c.initials);
      root ??= at;
      const threading = at === root ? '' : `<p:extLst><p:ext uri="${THREADING_EXT}"><p15:threadingInfo xmlns:p15="${NS_P15}" timeZoneBias="0"><p15:parentCm authorId="${root.author}" idx="${root.idx}"/></p15:threadingInfo></p:ext></p:extLst>`;
      items.push(`<p:cm${attrs({ authorId: at.author, dt: stamp(c.createdAt), idx: at.idx })}>${pos}<p:text>${esc(c.body)}</p:text>${threading}</p:cm>`);
    }
  }
  return `${XML_DECLARATION}<p:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${items.join('')}</p:cmLst>`;
}
