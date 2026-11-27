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

describe('K3c run marks and fields', () => {
  type Spec = { toDOM: (m: unknown) => unknown[]; parseDOM?: Array<{ getAttrs: (d: unknown) => unknown }> };
  const markDom = (type: string, attrs: Record<string, unknown>) => (slidesSchema.marks[type].spec as unknown as Spec).toDOM(slidesSchema.marks[type].create(attrs));

  it('round-trips the run marks and a field through JSON', () => {
    const doc: RichTextDoc = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'text', text: 'x2', marks: [{ type: 'baseline', attrs: { pct: 30 } }, { type: 'caps', attrs: { kind: 'small' } }, { type: 'charSpacing', attrs: { px: 2 } }, { type: 'lang', attrs: { lang: 'fr' } }] },
            { type: 'field', attrs: { kind: 'slidenum', format: null }, content: [{ type: 'text', text: '‹#›', marks: [{ type: 'bold' }] }] },
          ],
        },
      ],
    };
    const json = slidesSchema.nodeFromJSON(doc).toJSON();
    expect(json.content[0].content[0].marks.map((m: { type: string }) => m.type).sort()).toEqual(['baseline', 'caps', 'charSpacing', 'lang']);
    expect(json.content[0].content[1]).toMatchObject({ type: 'field', attrs: { kind: 'slidenum' }, content: [{ text: '‹#›', marks: [{ type: 'bold' }] }] });
  });

  it('renders each mark as a sanitized style or attribute, and nothing for a bad value', () => {
    expect(markDom('baseline', { pct: 30 })).toEqual(['span', { style: 'position:relative;top:-0.462em;font-size:0.65em;line-height:0' }, 0]);
    expect(markDom('baseline', { pct: 0 })).toEqual(['span', {}, 0]);
    expect(markDom('caps', { kind: 'all' })).toEqual(['span', { style: 'text-transform:uppercase' }, 0]);
    expect(markDom('caps', { kind: 'small' })).toEqual(['span', { style: 'font-variant:small-caps' }, 0]);
    expect(markDom('caps', { kind: 'x;color:red' })).toEqual(['span', {}, 0]);
    expect(markDom('charSpacing', { px: 1.5 })).toEqual(['span', { style: 'letter-spacing:1.5px' }, 0]);
    expect(markDom('charSpacing', { px: 'wide' })).toEqual(['span', {}, 0]);
    expect(markDom('lang', { lang: 'de-DE' })).toEqual(['span', { lang: 'de-DE' }, 0]);
    expect(markDom('lang', { lang: '"><script>' })).toEqual(['span', {}, 0]);
  });

  it('draws a field as a non-editable span around its text, naming the kind and format', () => {
    const toDOM = (attrs: Record<string, unknown>) => (slidesSchema.nodes.field.spec as unknown as Spec).toDOM(slidesSchema.nodes.field.create(attrs));
    expect(toDOM({ kind: 'slidenum' })).toEqual(['span', { 'data-field': 'slidenum', contenteditable: 'false' }, 0]);
    expect(toDOM({ kind: 'datetime', format: 'time' })).toEqual(['span', { 'data-field': 'datetime', contenteditable: 'false', 'data-format': 'time' }, 0]);
    expect(toDOM({ kind: 'bogus' })).toEqual(['span', { 'data-field': 'footer', contenteditable: 'false' }, 0]);
  });

  it('parses a field back from pasted HTML, degrading a crafted kind', () => {
    const parse = (slidesSchema.nodes.field.spec as unknown as Spec).parseDOM![0].getAttrs;
    const dom = (kind: string, format: string | null) => ({ getAttribute: (n: string) => (n === 'data-field' ? kind : format) });
    expect(parse(dom('slidenum', null))).toEqual({ kind: 'slidenum', format: null });
    expect(parse(dom('nope', 'date'))).toEqual({ kind: 'footer', format: 'date' });
  });
});
