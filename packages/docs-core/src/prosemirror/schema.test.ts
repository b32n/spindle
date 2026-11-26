import { docsSchema } from './schema';

// Extract the inline `style` string from a DOMOutputSpec array like
// ['p', { style: '…' }, 0].
function styleOf(spec: unknown): string {
  const attrs =
    Array.isArray(spec) && spec[1] && typeof spec[1] === 'object'
      ? (spec[1] as Record<string, string>)
      : {};
  return attrs.style || '';
}

function nodeStyle(name: string, attrs: Record<string, unknown>): string {
  const node = docsSchema.nodes[name].create(attrs);
  const toDOM = docsSchema.nodes[name].spec.toDOM as (n: typeof node) => unknown;
  return styleOf(toDOM(node));
}

function markStyle(name: string, attrs: Record<string, unknown>): string {
  const mark = docsSchema.marks[name].create(attrs);
  const toDOM = docsSchema.marks[name].spec.toDOM as (m: typeof mark, inline: boolean) => unknown;
  return styleOf(toDOM(mark, false));
}

describe('docsSchema toDOM sanitizes untrusted block attrs', () => {
  it('paragraph: whitelists alignment and coerces spacing', () => {
    expect(nodeStyle('paragraph', { alignment: 'center', spaceBefore: 4, spaceAfter: 6 })).toBe(
      'text-align: center; margin-top: 4px; margin-bottom: 6px;'
    );
    // A CSS-injection payload in alignment falls back to the default keyword.
    const injected = nodeStyle('paragraph', {
      alignment: 'left; background: url(//evil)',
      spaceBefore: '0; background: url(//evil)',
      spaceAfter: 8,
    });
    expect(injected).toContain('text-align: left;');
    expect(injected).not.toContain('url(');
    expect(injected).not.toContain('background');
  });

  it('heading: whitelists alignment', () => {
    expect(nodeStyle('heading', { level: 2, alignment: 'right' })).toBe('text-align: right');
    const injected = nodeStyle('heading', { level: 1, alignment: 'left; background: url(//evil)' });
    expect(injected).toBe('text-align: left');
  });

  it('heading: clamps level to a valid h1-h6 tag (no InvalidCharacterError)', () => {
    const tagOf = (attrs: Record<string, unknown>): string => {
      const node = docsSchema.nodes.heading.create(attrs);
      const spec = (docsSchema.nodes.heading.spec.toDOM as (n: typeof node) => unknown)(node);
      return Array.isArray(spec) ? String(spec[0]) : '';
    };
    expect(tagOf({ level: 2 })).toBe('h2');
    expect(tagOf({ level: 99 })).toBe('h6'); // clamped high
    expect(tagOf({ level: 0 })).toBe('h1'); // clamped low
    expect(tagOf({ level: -3 })).toBe('h1');
    expect(tagOf({ level: 2.9 })).toBe('h2'); // truncated
    expect(tagOf({ level: 'evil' })).toBe('h1'); // non-numeric → h1, valid tag
    expect(tagOf({ level: NaN })).toBe('h1');
  });

  it('textStyle: coerces fontSize to a number', () => {
    expect(markStyle('textStyle', { fontSize: 12 })).toContain('font-size: 12pt');
    const injected = markStyle('textStyle', { fontSize: '0pt; background: url(//evil)' });
    expect(injected).not.toContain('url(');
    expect(injected).not.toContain('background');
  });
});

describe('section_break', () => {
  it('is an atom block that renders its section id and has no parseDOM, so pasted HTML cannot create one', () => {
    const type = docsSchema.nodes.section_break;
    expect(type.isAtom).toBe(true);
    expect(type.isBlock).toBe(true);
    // Allowed between the blocks of the body only, not in table cells or list items.
    expect(docsSchema.nodes.doc.contentMatch.matchType(type)).not.toBeNull();
    expect(docsSchema.nodes.table_cell.contentMatch.matchType(type)).toBeNull();
    expect(docsSchema.nodes.list_item.contentMatch.matchType(type)).toBeNull();
    expect(docsSchema.nodes.table_header.contentMatch.matchType(type)).toBeNull();
    expect(type.spec.parseDOM).toBeUndefined();
    const toDOM = type.spec.toDOM as (n: ReturnType<typeof type.create>) => unknown;
    expect(toDOM(type.create({ sectionId: 'sec_1' }))).toEqual(['div', { class: 'section-break', 'data-section-id': 'sec_1' }]);
    // A section id with markup in it is an attribute value, not markup.
    expect(toDOM(type.create({ sectionId: '"><img src=x>' }))).toEqual(['div', { class: 'section-break', 'data-section-id': '"><img src=x>' }]);
  });
});
