import { createTextRun, getPlainText, hasTextContent, isBlockEmpty, cloneBlock } from './utils';
import { createParagraphFromText } from './paragraph';
import {
  createHeadingFromText,
  changeHeadingLevel,
  headingToParagraph,
  paragraphToHeading,
  getHeadingFontSize,
} from './heading';

describe('inline content helpers', () => {
  it('createTextRun builds a text run', () => {
    expect(createTextRun('hi', 'style_1')).toEqual({ type: 'text', text: 'hi', styleId: 'style_1' });
  });

  it('getPlainText concatenates text and link runs', () => {
    const content = [createTextRun('Hello '), createTextRun('world')];
    expect(getPlainText(content)).toBe('Hello world');
  });

  it('getPlainText leaves out text deleted by a tracked change, unless asked for it', () => {
    const gone = { ...createTextRun('2024'), revision: { deletion: { id: '1', author: 'Bo' } } };
    const content = [createTextRun('in '), gone, { ...createTextRun('2025'), revision: { insertion: { id: '2', author: 'Ada' } } }];
    expect(getPlainText(content)).toBe('in 2025');
    expect(getPlainText(content, true)).toBe('in 20242025');
    const link = { type: 'link' as const, text: 'old', href: 'https://a.example', revision: { deletion: { id: '3', author: 'Bo' } } };
    expect(getPlainText([createTextRun('x'), link])).toBe('x');
    // A block that holds only deleted text still has content.
    const para = createParagraphFromText('');
    para.content = [gone];
    expect(hasTextContent(para)).toBe(true);
    expect(isBlockEmpty(para)).toBe(false);
  });
});

describe('paragraph blocks', () => {
  it('createParagraphFromText carries the text and an id', () => {
    const p = createParagraphFromText('some text');
    expect(p.type).toBe('paragraph');
    expect(getPlainText(p.content)).toBe('some text');
    expect(p.id).toBeTruthy();
  });

  it('reports empty vs non-empty blocks', () => {
    expect(isBlockEmpty(createParagraphFromText(''))).toBe(true);
    expect(isBlockEmpty(createParagraphFromText('x'))).toBe(false);
    expect(hasTextContent(createParagraphFromText('x'))).toBe(true);
    expect(hasTextContent(createParagraphFromText(''))).toBe(false);
  });

  it('cloneBlock copies content but assigns a new id', () => {
    const p = createParagraphFromText('dup');
    const clone = cloneBlock(p);
    expect(clone.id).not.toBe(p.id);
    expect(getPlainText((clone as typeof p).content)).toBe('dup');
  });
});

describe('heading blocks', () => {
  it('creates a heading at a level and maps level → font size', () => {
    const h = createHeadingFromText(2, 'Title');
    expect(h.type).toBe('heading');
    expect(h.level).toBe(2);
    expect(getPlainText(h.content)).toBe('Title');
    expect(getHeadingFontSize(1)).toBeGreaterThan(getHeadingFontSize(3));
  });

  it('changes level without touching content or id', () => {
    const h = createHeadingFromText(1, 'X');
    const h3 = changeHeadingLevel(h, 3);
    expect(h3.level).toBe(3);
    expect(h3.id).toBe(h.id);
    expect(getPlainText(h3.content)).toBe('X');
  });

  it('converts heading ↔ paragraph preserving id and content', () => {
    const h = createHeadingFromText(2, 'Convert me');
    const p = headingToParagraph(h);
    expect(p.type).toBe('paragraph');
    expect(p.id).toBe(h.id);
    expect(getPlainText(p.content)).toBe('Convert me');

    const back = paragraphToHeading(p, 4);
    expect(back.type).toBe('heading');
    expect(back.level).toBe(4);
    expect(back.id).toBe(h.id);
  });
});
