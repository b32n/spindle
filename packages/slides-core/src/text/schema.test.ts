import { slidesSchema } from './schema';
import type { RichTextDoc } from './model';

const rich: RichTextDoc = {
  type: 'doc',
  content: [
    {
      type: 'paragraph',
      attrs: { align: 'center', listType: 'bullet', indent: 1, lineHeight: 1.4, spaceBefore: 0, spaceAfter: 8 },
      content: [
        { type: 'text', text: 'Bold ', marks: [{ type: 'bold' }] },
        { type: 'text', text: 'blue', marks: [{ type: 'textColor', attrs: { color: { kind: 'rgb', hex: '#2D7FF9' } } }] },
        { type: 'text', text: ' big', marks: [{ type: 'fontSize', attrs: { size: 32 } }] },
      ],
    },
  ],
};

describe('slidesSchema round-trip', () => {
  it('parses RichTextDoc JSON and re-serializes it losslessly', () => {
    const node = slidesSchema.nodeFromJSON(rich);
    const json = node.toJSON();
    // Marks + text + attrs survive.
    expect(json.content[0].content[0].marks[0].type).toBe('bold');
    expect(json.content[0].content[1].marks[0].attrs.color.hex).toBe('#2D7FF9');
    expect(json.content[0].content[2].marks[0].attrs.size).toBe(32);
    expect(json.content[0].attrs.align).toBe('center');
    expect(json.content[0].attrs.listType).toBe('bullet');
  });

  it('is idempotent (re-parsing the serialized form yields the same JSON)', () => {
    const json1 = slidesSchema.nodeFromJSON(rich).toJSON();
    const json2 = slidesSchema.nodeFromJSON(json1).toJSON();
    expect(json2).toEqual(json1);
  });

  it('fills default paragraph attrs for a bare doc', () => {
    const json = slidesSchema.nodeFromJSON({ type: 'doc', content: [{ type: 'paragraph' }] }).toJSON();
    expect(json.content[0].attrs).toMatchObject({ align: 'left', listType: 'none', indent: 0 });
  });
});

describe('slidesSchema toDOM sanitizes untrusted color/font attrs', () => {
  // Extract the inline `style` string from a DOMOutputSpec array like
  // ['span', { style: '…' }, 0].
  const styleOf = (spec: unknown): string => {
    const attrs = Array.isArray(spec) && spec[1] && typeof spec[1] === 'object' ? (spec[1] as Record<string, string>) : {};
    return attrs.style || '';
  };
  const renderMark = (name: string, attrs: Record<string, unknown>): string => {
    const mark = slidesSchema.marks[name].create(attrs);
    const toDOM = slidesSchema.marks[name].spec.toDOM as (m: typeof mark, inline: boolean) => unknown;
    return styleOf(toDOM(mark, false));
  };

  it('drops a CSS-injection payload smuggled through a literal hex', () => {
    expect(renderMark('textColor', { color: { kind: 'rgb', hex: 'red;background:url(//evil)' } })).toBe('');
  });

  it('keeps a valid literal color', () => {
    expect(renderMark('textColor', { color: { kind: 'rgb', hex: '#2D7FF9' } })).toBe('color:#2D7FF9');
  });

  it('emits var(--slot-…) only for a recognized theme slot', () => {
    expect(renderMark('textColor', { color: { kind: 'theme', slot: 'accent1' } })).toBe('color:var(--slot-accent1)');
    expect(renderMark('textColor', { color: { kind: 'theme', slot: 'x)}{color:red' } })).toBe('');
  });

  it('sanitizes highlight (background-color) the same way', () => {
    expect(renderMark('highlight', { color: { kind: 'rgb', hex: '#abc' } })).toBe('background-color:#abc');
    expect(renderMark('highlight', { color: { kind: 'rgb', hex: 'x;background:url(//evil)' } })).toBe('');
  });

  it('resolves symbolic fonts to CSS vars and rejects injection in literal fonts', () => {
    expect(renderMark('fontFamily', { family: 'major' })).toBe('font-family:var(--font-major)');
    expect(renderMark('fontFamily', { family: 'Georgia, serif' })).toBe('font-family:Georgia, serif');
    expect(renderMark('fontFamily', { family: 'Arial;background:url(//evil)' })).toBe('');
  });

  it('coerces a crafted font-size to a number', () => {
    expect(renderMark('fontSize', { size: '10px;background:url(//evil)' })).toBe('font-size:calc(18px * var(--spindle-fs, 1))');
    expect(renderMark('fontSize', { size: 24 })).toBe('font-size:calc(24px * var(--spindle-fs, 1))');
  });
});

describe('slidesSchema paragraph toDOM sanitizes untrusted align/line-height/spacing', () => {
  const styleOfNode = (attrs: Record<string, unknown>): string => {
    const node = slidesSchema.nodes.paragraph.create(attrs);
    const toDOM = slidesSchema.nodes.paragraph.spec.toDOM as (n: typeof node) => unknown;
    const spec = toDOM(node);
    const a = Array.isArray(spec) && spec[1] && typeof spec[1] === 'object' ? (spec[1] as Record<string, string>) : {};
    return a.style || '';
  };

  it('whitelists text-align and drops an injection payload', () => {
    expect(styleOfNode({ align: 'center' })).toContain('text-align:center;');
    // A crafted align falls back to the default keyword, not the raw payload.
    const injected = styleOfNode({ align: 'left;background:url(//evil)' });
    expect(injected).toContain('text-align:left;');
    expect(injected).not.toContain('url(');
  });

  it('coerces line-height and spacing to numbers', () => {
    expect(styleOfNode({ lineHeight: 1.5 })).toContain('line-height:calc(1.5 * var(--spindle-ls, 1));');
    const injected = styleOfNode({ lineHeight: '1;background:url(//evil)', spaceBefore: '4;x', spaceAfter: '2}x' });
    expect(injected).toContain('line-height:calc(1 * var(--spindle-ls, 1));'); // parseFloat keeps the leading number
    expect(injected).not.toContain('url(');
    expect(injected).not.toContain('background');
    // Non-numeric spacing is dropped entirely rather than interpolated raw.
    expect(injected).not.toContain(';x');
    expect(injected).not.toContain('}x');
  });
});
