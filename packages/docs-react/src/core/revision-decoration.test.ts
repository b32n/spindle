import { addTitle, revisionDecoration } from './revision-decoration';
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
      data: { revisionAuthor: 'Ada', revisionInsertion: '1' },
      title: 'Inserted by Ada',
    });
  });

  it('strikes through a deletion', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { deletion } })).toEqual({
      color: '#188038',
      textDecoration: 'line-through',
      data: { revisionAuthor: 'Bo', revisionDeletion: '2' },
      title: 'Deleted by Bo',
    });
  });

  it('draws text that is both in the colour of the deletion, with both lines and both ids', () => {
    expect(revisionDecoration({ kind: 'text', text: 'x', revision: { insertion, deletion } })).toEqual({
      color: '#188038',
      textDecoration: 'underline line-through',
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

describe('addTitle', () => {
  it('sets a tooltip and adds to one that is there', () => {
    const el = { title: '' };
    addTitle(el, 'Inserted by Ada');
    expect(el.title).toBe('Inserted by Ada');
    addTitle(el, 'Name');
    expect(el.title).toBe('Inserted by Ada\nName');
  });
});
