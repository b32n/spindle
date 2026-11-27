import { contentTypeOf, EMPTY_CONTENT_TYPES, parseContentTypes, partsWithContentType } from './content-types';
import { el } from './test-support';

const NS = 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"';

describe('parseContentTypes', () => {
  const types = parseContentTypes(
    el(`<Types ${NS}>
      <Default Extension="XML" ContentType="application/xml"/>
      <Default Extension="xml" ContentType="ignored: the first default of an extension wins"/>
      <Default Extension="png" ContentType="image/png"/>
      <Default ContentType="no-extension"/>
      <Default Extension="bin"/>
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
      <Override PartName="/xl/workbook.xml" ContentType="ignored"/>
      <Override PartName="xl/Styles.xml" ContentType="styles"/>
      <Override PartName="/x.xml"/>
      <Override ContentType="no-part"/>
    </Types>`),
  );
  it('keys defaults by lower-cased extension and overrides by lower-cased part name without the slash', () => {
    expect([...types.defaults]).toEqual([
      ['xml', 'application/xml'],
      ['png', 'image/png'],
    ]);
    expect([...types.overrides.keys()]).toEqual(['xl/workbook.xml', 'xl/styles.xml']);
  });
  it('answers an override first, then the extension default, case-insensitively', () => {
    expect(contentTypeOf(types, 'xl/workbook.xml')).toMatch(/sheet\.main\+xml$/);
    expect(contentTypeOf(types, 'XL/STYLES.XML')).toBe('styles');
    expect(contentTypeOf(types, 'xl/media/a.PNG')).toBe('image/png');
    expect(contentTypeOf(types, 'xl/media/a.gif')).toBeUndefined();
    expect(contentTypeOf(types, 'noextension')).toBeUndefined();
  });
  it('lists the parts whose type passes a test', () => {
    expect(partsWithContentType(types, (t) => t === 'styles')).toEqual(['xl/styles.xml']);
  });
  it('is empty for a missing root', () => {
    const none = parseContentTypes(undefined);
    expect(none.defaults.size + none.overrides.size).toBe(0);
    expect(EMPTY_CONTENT_TYPES.defaults.size).toBe(0);
  });
});
