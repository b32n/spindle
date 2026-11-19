import { Schema, Node as PmNode } from 'prosemirror-model';
import { proseMirrorToFlowBlocks } from './pm-to-blocks';
import type { ParagraphBlock, LinkRun } from './flow-blocks';

// Minimal schema: enough to build a paragraph containing a linked text run.
const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*', toDOM: () => ['p', 0] },
    text: { group: 'inline' },
  },
  marks: {
    link: {
      attrs: { href: {} },
      toDOM: (m) => ['a', { href: m.attrs.href }, 0],
    },
  },
});

function docWithLink(href: string): PmNode {
  const mark = schema.marks.link.create({ href });
  const textNode = schema.text('click me', [mark]);
  const para = schema.nodes.paragraph.create(null, textNode);
  return schema.nodes.doc.create(null, para);
}

describe('proseMirrorToFlowBlocks — link marks', () => {
  it('maps a link mark to a LinkRun carrying the href (not a plain text run)', () => {
    const blocks = proseMirrorToFlowBlocks(docWithLink('https://example.com'));
    const para = blocks[0] as ParagraphBlock;
    expect(para.kind).toBe('paragraph');
    const run = para.runs[0] as LinkRun;
    expect(run.kind).toBe('link');
    expect(run.href).toBe('https://example.com');
    expect(run.text).toBe('click me');
    expect(run.underline).toBe(true);
  });

  it('neutralizes a javascript: href when building the model', () => {
    const blocks = proseMirrorToFlowBlocks(docWithLink('javascript:alert(1)'));
    const run = (blocks[0] as ParagraphBlock).runs[0] as LinkRun;
    expect(run.kind).toBe('link');
    expect(run.href).toBe('');
  });
});

// The real docs schema: paragraph formatting and run flags travel through ProseMirror attrs and marks.
import { docsSchema } from '@b32nio/spindle-docs-core';
import type { ListItemBlock, TextRun } from './flow-blocks';

const j = (json: unknown) => docsSchema.nodeFromJSON({ type: 'doc', content: json });
const text = (s: string, marks?: Array<{ type: string; attrs?: object }>) => (marks ? { type: 'text', text: s, marks } : { type: 'text', text: s });

describe('proseMirrorToFlowBlocks: DOCX paragraph and run properties', () => {
  it('maps indents, spacing and a custom line spacing of a paragraph', () => {
    const doc = j([
      { type: 'paragraph', attrs: { leftIndent: 36, rightIndent: 12, firstLineIndent: -18, spaceAfter: 0, lineSpacing: { type: 'multiple', value: 1.2 } }, content: [text('x')] },
    ]);
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    expect(para.attrs).toMatchObject({ leftIndent: 36, rightIndent: 12, firstLineIndent: -18, spaceAfter: 0 });
    expect(para.attrs?.lineHeight).toBeCloseTo(1.8); // 1.2 x the layout's single-spacing default of 1.5
  });

  it('scales the keyword spacings and ignores heights it cannot express', () => {
    const lineHeight = (lineSpacing: unknown) =>
      (proseMirrorToFlowBlocks(j([{ type: 'paragraph', attrs: { lineSpacing }, content: [text('x')] }]))[0] as ParagraphBlock).attrs?.lineHeight;
    expect(lineHeight('single')).toBeUndefined();
    expect(lineHeight({ type: 'onePointFive' })).toBe(2.25);
    expect(lineHeight({ type: 'double' })).toBe(3);
    expect(lineHeight({ type: 'exactly', value: 20 })).toBeUndefined();
    expect(lineHeight({ type: 'atLeast', value: 20 })).toBeUndefined();
    expect(lineHeight({ type: 'multiple', value: 0 })).toBeUndefined();
    expect(lineHeight({ type: 'multiple', value: 99 })).toBeUndefined();
  });

  it('carries heading spacing and indents', () => {
    const doc = j([{ type: 'heading', attrs: { level: 2, spaceBefore: 12, spaceAfter: 0, leftIndent: 8 }, content: [text('H')] }]);
    expect(proseMirrorToFlowBlocks(doc)[0]).toMatchObject({ kind: 'heading', attrs: { spaceBefore: 12, spaceAfter: 0, leftIndent: 8 } });
  });

  it('maps small caps, all caps and hidden marks onto the run', () => {
    const doc = j([
      { type: 'paragraph', content: [text('a', [{ type: 'smallCaps' }]), text('b', [{ type: 'allCaps' }]), text('c', [{ type: 'hidden' }]), text('d')] },
    ]);
    const runs = (proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock).runs as TextRun[];
    expect(runs.map((r) => [r.smallCaps, r.allCaps, r.hidden])).toEqual([
      [true, undefined, undefined],
      [undefined, true, undefined],
      [undefined, undefined, true],
      [undefined, undefined, undefined],
    ]);
  });

  it('ignores the named style mark and attr (the formatting is already on the run and block)', () => {
    const doc = j([
      { type: 'paragraph', attrs: { namedStyleId: 'Callout' }, content: [text('a', [{ type: 'namedStyle', attrs: { id: 'AccentWord' } }, { type: 'bold' }])] },
    ]);
    const para = proseMirrorToFlowBlocks(doc)[0] as ParagraphBlock;
    const run = para.runs[0] as TextRun;
    expect(run.text).toBe('a');
    expect(run.bold).toBe(true);
    expect(JSON.stringify(para)).not.toContain('AccentWord');
  });

  it('numbers an ordered list from its start attribute', () => {
    const item = (s: string) => ({ type: 'list_item', attrs: { level: 0 }, content: [{ type: 'paragraph', content: [text(s)] }] });
    const doc = j([{ type: 'ordered_list', attrs: { start: 7 }, content: [item('a'), item('b')] }, { type: 'ordered_list', content: [item('c')] }]);
    const items = proseMirrorToFlowBlocks(doc) as ListItemBlock[];
    expect(items.map((i) => i.listIndex)).toEqual([7, 8, 1]);
  });
});
