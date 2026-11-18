import { createHash } from 'node:crypto';
import { digestDataUrl, normalizeModel, resolveStyles } from '../../../../scripts/docx-fidelity/normalize';
import type { Json } from '../../../../scripts/docx-fidelity/types';
import { get } from './support';

const model = (extra: Record<string, unknown> = {}) => ({
  id: 'doc-abc',
  title: 'T',
  createdAt: '2026-01-01',
  updatedAt: '2026-01-02',
  sections: [
    {
      id: 'section_9',
      pageConfig: { orientation: 'portrait' },
      blocks: [
        { id: 'block_1700_1_zz', type: 'paragraph', styleId: 'style_b', content: [{ type: 'text', text: 'hi', styleId: 'style_t2' }] },
        {
          id: 'block_1700_2_yy',
          type: 'table',
          rows: [{ id: 'row_1_a', cells: [{ id: 'cell_1_a', content: [{ type: 'text', text: 'a', styleId: 'style_t1' }] }, { id: 'cell_2_b', content: [] }] }],
        },
        { id: 'block_3', type: 'image', src: 'data:image/png;base64,AAAA', width: 1, height: 1 },
        { id: 'block_4', type: 'image', src: 'https://example.com/x.png', width: 1, height: 1 },
      ],
    },
    { id: 'section_10', blocks: [{ id: 'block_5', type: 'paragraph', content: [{ type: 'link', text: 'l', href: 'x', styleId: 'style_t1' }] }] },
  ],
  textStylePool: { style_t1: { bold: true }, style_t2: { italic: true }, style_unused2: { color: '#fff' }, style_unused1: { color: '#000' } },
  paragraphStylePool: { style_b: { alignment: 'center' } },
  ...extra,
});

describe('normalizeModel', () => {
  it('replaces generated ids with sequential placeholders in document order', () => {
    const out = normalizeModel(model());
    expect([get(out, 'sections.0.id'), get(out, 'sections.1.id')]).toEqual(['section-1', 'section-2']);
    expect([0, 1, 2, 3].map((i) => get(out, `sections.0.blocks.${i}.id`))).toEqual(['block-1', 'block-2', 'block-3', 'block-4']);
    expect(get(out, 'sections.1.blocks.0.id')).toBe('block-5');
    expect(get(out, 'sections.0.blocks.1.rows.0.id')).toBe('row-1');
    expect([0, 1].map((i) => get(out, `sections.0.blocks.1.rows.0.cells.${i}.id`))).toEqual(['cell-1', 'cell-2']);
  });

  it('drops the document id and timestamps', () => {
    const out = normalizeModel(model()) as Record<string, unknown>;
    expect(out).not.toHaveProperty('id');
    expect(out).not.toHaveProperty('createdAt');
    expect(out).not.toHaveProperty('updatedAt');
    expect(out.title).toBe('T');
  });

  it('renames style pool keys by first reference, then unreferenced entries by content', () => {
    const out = normalizeModel(model());
    expect(get(out, 'paragraphStylePool')).toEqual({ 'paragraph-style-1': { alignment: 'center' } });
    // referenced in order: style_t2 (block 1), style_t1 (table), then unused sorted by content
    expect(get(out, 'textStylePool')).toEqual({
      'text-style-1': { italic: true },
      'text-style-2': { bold: true },
      'text-style-3': { color: '#000' },
      'text-style-4': { color: '#fff' },
    });
    expect(get(out, 'sections.0.blocks.0.styleId')).toBe('paragraph-style-1');
    expect(get(out, 'sections.0.blocks.0.content.0.styleId')).toBe('text-style-1');
    expect(get(out, 'sections.1.blocks.0.content.0.styleId')).toBe('text-style-2');
  });

  it('is independent of the generated id values', () => {
    const renamed = JSON.parse(JSON.stringify(model()).replace(/block_|style_|section_|row_|cell_/g, (m) => `${m}x`));
    expect(normalizeModel(renamed)).toEqual(normalizeModel(model()));
  });

  it('replaces data URLs with a digest and leaves other sources alone', () => {
    const out = normalizeModel(model());
    expect(get(out, 'sections.0.blocks.2.src')).toBe(digestDataUrl('data:image/png;base64,AAAA'));
    expect(get(out, 'sections.0.blocks.3.src')).toBe('https://example.com/x.png');
  });

  it('numbers ids it has no kind for and tolerates missing pools', () => {
    const out = normalizeModel({ id: 'x', sections: [], threads: [{ id: 't', comments: [{ id: 'c1' }, { id: 'c2' }] }], extra: { id: 'q' } });
    expect(get(out, 'threads.0.id')).toBe('thread-1');
    expect([get(out, 'threads.0.comments.0.id'), get(out, 'threads.0.comments.1.id')]).toEqual(['comment-1', 'comment-2']);
    expect(get(out, 'extra.id')).toBe('id-1');
    expect(out).not.toHaveProperty('textStylePool');
  });

  it('keeps references to styles that are not in the pool consistent', () => {
    const out = normalizeModel({ sections: [{ id: 's', blocks: [{ id: 'b', type: 'paragraph', styleId: 'ghost', content: [] }] }], paragraphStylePool: {} });
    expect(get(out, 'sections.0.blocks.0.styleId')).toBe('paragraph-style-1');
    expect(get(out, 'paragraphStylePool')).toEqual({});
  });

  it('passes non-object input through', () => {
    expect(normalizeModel(null)).toBeNull();
    expect(normalizeModel(undefined)).toBeNull();
    expect(normalizeModel('x')).toBe('x');
  });
});

describe('digestDataUrl', () => {
  it('is deterministic and records mime type and size', () => {
    const url = 'data:image/png;base64,iVBORw0KGgo=';
    const sha = createHash('sha256').update(url).digest('hex').slice(0, 16);
    expect(digestDataUrl(url)).toBe(`data:image/png;sha256=${sha};chars=12`);
    expect(digestDataUrl('data:,')).toMatch(/^data:unknown;sha256=[0-9a-f]{16};chars=0$/);
    expect(digestDataUrl('data:text/plain')).toMatch(/^data:text\/plain;sha256=[0-9a-f]{16};chars=0$/);
  });
});

describe('resolveStyles', () => {
  it('inlines pool entries and drops the pools', () => {
    const normalized = normalizeModel(model());
    const out = resolveStyles(normalized);
    expect(out).not.toHaveProperty('textStylePool');
    expect(out).not.toHaveProperty('paragraphStylePool');
    expect(get(out, 'sections.0.blocks.0.style')).toEqual({ alignment: 'center' });
    expect(get(out, 'sections.0.blocks.0.content.0.style')).toEqual({ italic: true });
    expect(get(out, 'sections.0.blocks.0')).not.toHaveProperty('styleId');
  });

  it('uses null for unknown style ids and passes scalars through', () => {
    const out = resolveStyles({ sections: [{ type: 'text', styleId: 'nope' }] } as Json);
    expect(get(out, 'sections.0.style')).toBeNull();
    expect(resolveStyles('s')).toBe('s');
  });
});
