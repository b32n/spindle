import JSZip from 'jszip';
import type { DocBlock, DocCommentThread, DocInline, DocSection } from '@b32nio/spindle-shared/export/docx';

// ---------------------------------------------------------------------------
// Comments on export.
//
// The `docx` library writes the range markers (commentRangeStart/End and the reference run) when handed
// them as paragraph children, but it does not write threading (commentsExtended) or authors (people), and its
// comments part rewrites dates. So the comment parts are generated here and added to the finished package
// (zip post-processing; no network, no new dependency).
// ---------------------------------------------------------------------------

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const COMMENTS_EXTENDED_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const PEOPLE_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml';
const COMMENTS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const COMMENTS_EXTENDED_REL = 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const PEOPLE_REL = 'http://schemas.microsoft.com/office/2011/relationships/people';

/** A thread with the numeric `w:id` of each of its comments (what the range markers refer to). */
export interface CommentPlanThread {
  thread: DocCommentThread;
  ids: number[];
  /** Tagged runs in the body, in total, and how many the writer has emitted so far. */
  total: number;
  seen: number;
}

export interface CommentPlan {
  threads: Map<string, CommentPlanThread>;
  /** Threads with no tagged run: anchored (collapsed) at the first body paragraph. */
  orphans: CommentPlanThread[];
}

function runsOf(content: DocInline[], visit: (threadId: string) => void): void {
  for (const inline of content) {
    if ((inline.type === 'text' || inline.type === 'link') && inline.commentThreadId) visit(inline.commentThreadId);
  }
}

function blockRuns(block: DocBlock, visit: (threadId: string) => void): void {
  if (block.type === 'table') {
    for (const row of block.rows) for (const cell of row.cells) runsOf(cell.content, visit);
  } else if ('content' in block) {
    runsOf(block.content, visit);
  }
}

/** Number every comment (document-wide, from 0) and count how many body runs each thread is anchored to. */
export function planComments(threads: DocCommentThread[] | undefined, sections: DocSection[]): CommentPlan {
  const plan: CommentPlan = { threads: new Map(), orphans: [] };
  let next = 0;
  for (const thread of threads ?? []) {
    if (thread.comments.length === 0 || plan.threads.has(thread.id)) continue;
    plan.threads.set(thread.id, { thread, ids: thread.comments.map(() => next++), total: 0, seen: 0 });
  }
  if (plan.threads.size === 0) return plan;
  for (const section of sections) {
    for (const block of section.blocks) {
      blockRuns(block, (id) => {
        const planned = plan.threads.get(id);
        if (planned) planned.total++;
      });
    }
  }
  plan.orphans = [...plan.threads.values()].filter((t) => t.total === 0);
  return plan;
}

// ---------------------------------------------------------------------------
// Parts
// ---------------------------------------------------------------------------

/** Control characters other than tab, newline and carriage return cannot appear in XML 1.0. */
function stripControlChars(text: string): string {
  return [...text].filter((ch) => {
    const code = ch.charCodeAt(0);
    return code >= 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
  }).join('');
}

function xml(text: string): string {
  return stripControlChars(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** `w:date` is an xsd:dateTime; Word writes whole seconds, so fractions are dropped and anything unparseable omitted. */
function wordDate(date: string | undefined): string | undefined {
  const m = date?.match(/^(\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d)(?:\.\d+)?(Z|[+-]\d\d:\d\d)?$/);
  return m ? m[1] + (m[2] ?? '') : undefined;
}

function paragraphXml(line: string, paraId: string, first: boolean): string {
  const runs = line
    .split('\t')
    .map((part) => (part ? `<w:r><w:t xml:space="preserve">${xml(part)}</w:t></w:r>` : ''))
    .join('<w:r><w:tab/></w:r>');
  return `<w:p w14:paraId="${paraId}">${first ? '<w:r><w:annotationRef/></w:r>' : ''}${runs}</w:p>`;
}

const paraIdFor = (n: number) => (0x10000000 + n).toString(16).toUpperCase();

export interface CommentParts {
  comments: string;
  commentsExtended: string;
  people: string;
}

export function buildCommentParts(plan: CommentPlan): CommentParts {
  const comments: string[] = [];
  const extended: string[] = [];
  const authors = new Map<string, string>();
  let paraCounter = 0;
  for (const { thread, ids } of plan.threads.values()) {
    let rootPara = '';
    thread.comments.forEach((c, i) => {
      const lines = c.text.split('\n');
      const paras = lines.map((line, n) => {
        const paraId = paraIdFor(++paraCounter);
        return { paraId, xml: paragraphXml(line, paraId, n === 0) };
      });
      const lastPara = paras[paras.length - 1].paraId;
      if (i === 0) rootPara = lastPara;
      const date = wordDate(c.date);
      comments.push(
        `<w:comment w:id="${ids[i]}" w:author="${xml(c.author)}"${date ? ` w:date="${date}"` : ''}` +
          `${c.initials ? ` w:initials="${xml(c.initials)}"` : ''}>${paras.map((p) => p.xml).join('')}</w:comment>`
      );
      extended.push(
        `<w15:commentEx w15:paraId="${lastPara}"${i > 0 ? ` w15:paraIdParent="${rootPara}"` : ''} w15:done="${thread.resolved ? 1 : 0}"/>`
      );
      if (!authors.has(c.author)) authors.set(c.author, c.authorId ?? c.author);
    });
  }
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const ns = `xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14 w15"`;
  return {
    comments: `${head}<w:comments ${ns}>${comments.join('')}</w:comments>`,
    commentsExtended: `${head}<w15:commentsEx ${ns}>${extended.join('')}</w15:commentsEx>`,
    people:
      `${head}<w15:people ${ns}>` +
      [...authors]
        .map(
          ([author, id]) =>
            `<w15:person w15:author="${xml(author)}"><w15:presenceInfo w15:providerId="None" w15:userId="${xml(id)}"/></w15:person>`
        )
        .join('') +
      '</w15:people>',
  };
}

/** Add the comment parts, their content types and relationships to a package the `docx` library produced. */
export async function addCommentParts(bytes: Uint8Array, parts: CommentParts): Promise<Uint8Array<ArrayBuffer>> {
  const zip = await JSZip.loadAsync(bytes);
  zip.file('word/comments.xml', parts.comments);
  zip.file('word/commentsExtended.xml', parts.commentsExtended);
  zip.file('word/people.xml', parts.people);

  const types = (await zip.file('[Content_Types].xml')?.async('string')) ?? '';
  const overrides = [
    ['/word/comments.xml', COMMENTS_TYPE],
    ['/word/commentsExtended.xml', COMMENTS_EXTENDED_TYPE],
    ['/word/people.xml', PEOPLE_TYPE],
  ]
    .map(([name, type]) => `<Override ContentType="${type}" PartName="${name}"/>`)
    .join('');
  zip.file('[Content_Types].xml', types.replace('</Types>', `${overrides}</Types>`));

  const relsPath = 'word/_rels/document.xml.rels';
  const rels = (await zip.file(relsPath)?.async('string')) ?? '';
  let nextId = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map((m) => Number(m[1]))) + 1;
  const added = [
    [COMMENTS_REL, 'comments.xml'],
    [COMMENTS_EXTENDED_REL, 'commentsExtended.xml'],
    [PEOPLE_REL, 'people.xml'],
  ]
    .map(([type, target]) => `<Relationship Id="rId${nextId++}" Type="${type}" Target="${target}"/>`)
    .join('');
  zip.file(relsPath, rels.replace('</Relationships>', `${added}</Relationships>`));

  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
