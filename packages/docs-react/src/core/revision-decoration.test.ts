import { addTitle, revisionDecoration, rowRevisionTint } from './revision-decoration';
import { DEFAULT_REVISION_COLOR, type Run } from './flow-blocks';

const insertion = { id: '1', author: 'Ada', color: '#1a73e8' };
const deletion = { id: '2', author: 'Bo', color: '#188038' };

describe('revisionDecoration', () => {
  it('is nothing for a run without a change, for a break, and for a change without an author', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x' })).toBeNull();
    expect(revisionDecoration({ kind: 'lineBreak' })).toBeNull();
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: {} })).toBeNull();
  });

  it('underlines an insertion in its author\'s colour and names it', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion } })).toEqual({
      color: '#1a73e8',
      textDecoration: 'underline',
      textDecorationColor: '#1a73e8',
      data: { revisionAuthor: 'Ada', revisionInsertion: '1' },
      title: 'Inserted by Ada',
    });
  });

  it('strikes through a deletion', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { deletion } })).toEqual({
      color: '#188038',
      textDecoration: 'line-through',
      textDecorationColor: '#188038',
      data: { revisionAuthor: 'Bo', revisionDeletion: '2' },
      title: 'Deleted by Bo',
    });
  });

  it('draws text that is both in the colour of the deletion, with both lines and both ids', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion, deletion } })).toEqual({
      color: '#188038',
      textDecoration: 'underline line-through',
      textDecorationColor: '#188038',
      data: { revisionAuthor: 'Bo', revisionInsertion: '1', revisionDeletion: '2' },
      title: 'Inserted by Ada; Deleted by Bo',
    });
  });

  it('keeps the underline and strike-through the text has', () => {
    const run: Run = { kind: 'text', text: 'x', underline: true, strikethrough: true, revision: { insertion } };
    expect(revisionDecoration(run)?.textDecoration).toBe('underline line-through');
  });

  it('keeps the underline of a link, but not a link without one', () => {
    const link = { kind: 'link' as const, text: 'x', href: 'https://example.com/', revision: { deletion } };
    expect(revisionDecoration(link)?.textDecoration).toBe('underline line-through');
    expect(revisionDecoration({ ...link, plain: true })?.textDecoration).toBe('line-through');
    expect(revisionDecoration({ ...link, plain: true, revision: { insertion } })?.textDecoration).toBe('underline');
  });

  it('uses the default colour when the author colour is missing or unsafe', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion: { id: '1', author: 'Ada' } as never } })?.color).toBe(DEFAULT_REVISION_COLOR);
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion: { id: '1', author: 'Ada', color: 'red; background: url(x)' } } })?.color).toBe(DEFAULT_REVISION_COLOR);
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion: { id: '1', author: 'Ada', color: '' } } })?.color).toBe(DEFAULT_REVISION_COLOR);
  });
});

describe('revisionDecoration of moves and formatting changes', () => {
  const format = { id: '3', author: 'Cy', color: '#a142f4' };

  it('draws the two ends of a move with a double line and says which end it is', () => {
    const from = revisionDecoration({ kind: 'text', text: 'x', revision: { deletion: { ...deletion, moveId: 'm1' } } });
    expect(from).toMatchObject({ textDecoration: 'line-through', textDecorationStyle: 'double', data: { revisionMove: 'm1' }, title: 'Moved from here by Bo' });
    const to = revisionDecoration({ kind: 'text', text: 'x', revision: { insertion: { ...insertion, moveId: 'm1' } } });
    expect(to).toMatchObject({ textDecoration: 'underline', textDecorationStyle: 'double', data: { revisionMove: 'm1' }, title: 'Moved to here by Ada' });
  });

  it('draws a change of formatting alone dotted, in the author\'s colour, and keeps the text colour', () => {
    const d = revisionDecoration({ kind: 'text', text: 'x', revision: { formatChange: format } })!;
    expect(d).toMatchObject({ textDecoration: 'underline', textDecorationColor: '#a142f4', textDecorationStyle: 'dotted', data: { revisionFormat: '3', revisionAuthor: 'Cy' }, title: 'Formatted by Cy' });
    expect(d).not.toHaveProperty('color');
  });

  it('shows an underlined run with a formatting change alone as dotted, and keeps its strike-through', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', underline: true, strikethrough: true, revision: { formatChange: format } })).toMatchObject({ textDecoration: 'underline line-through', textDecorationStyle: 'dotted' });
  });

  it('shows an insertion over a formatting change, and names the author of both', () => {
    const d = revisionDecoration({ kind: 'text', text: 'x', revision: { insertion, formatChange: format } })!;
    expect(d).toMatchObject({ color: '#1a73e8', data: { revisionAuthor: 'Ada', revisionInsertion: '1', revisionFormat: '3', revisionFormatAuthor: 'Cy' }, title: 'Inserted by Ada; Formatted by Cy' });
    expect(d).not.toHaveProperty('textDecorationStyle');
  });
});

describe('rowRevisionTint', () => {
  it('tints an inserted or a deleted row in the author\'s colour, the deletion first', () => {
    expect(rowRevisionTint({ insertion: { id: '1', author: 'Ada', color: '#1a73e8' } })).toEqual({
      backgroundImage: 'linear-gradient(#1a73e826, #1a73e826)',
      data: { revisionRow: 'insertion', revisionAuthor: 'Ada' },
    });
    expect(rowRevisionTint({ insertion: insertion, deletion })).toMatchObject({ data: { revisionRow: 'deletion', revisionAuthor: 'Bo' } });
  });

  it('is nothing for a row without a change, or with a colour that is not a six digit hex colour', () => {
    expect(rowRevisionTint(undefined)).toBeNull();
    expect(rowRevisionTint({})).toBeNull();
    expect(rowRevisionTint({ insertion: { id: '1', author: 'Ada', color: 'red' } })).toBeNull();
    expect(rowRevisionTint({ insertion: { id: '1', author: 'Ada', color: '#fff' } })).toBeNull();
    expect(rowRevisionTint({ insertion: { id: '1', author: 'Ada', color: 'url(x)' } })).toBeNull();
  });
});

describe('addTitle', () => {
  it('sets a tooltip and adds to one that is there', () => {
    const el = { title: '' };
    addTitle(el, 'Inserted by Ada');
    expect(el.title).toBe('Inserted by Ada');
    addTitle(el, 'Name');
    expect(el.title).toBe('Inserted by Ada\nName');
  });
});
