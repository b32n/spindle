import { sanitizeHref, sanitizeImageSrc, safeCssColor, safeFontFamily, safeFontName } from './sanitize';

describe('sanitizeHref', () => {
  it('neutralizes javascript: hrefs', () => {
    expect(sanitizeHref('javascript:alert(1)')).toBe('');
    expect(sanitizeHref('JavaScript:alert(1)')).toBe('');
    expect(sanitizeHref('  javascript:alert(1)  ')).toBe('');
    // Browsers ignore whitespace/control chars inside the scheme.
    expect(sanitizeHref('java\tscript:alert(1)')).toBe('');
    expect(sanitizeHref('java\nscript:alert(1)')).toBe('');
  });

  it('rejects vbscript: and data: hrefs', () => {
    expect(sanitizeHref('vbscript:msgbox(1)')).toBe('');
    expect(sanitizeHref('data:text/html,<script>alert(1)</script>')).toBe('');
  });

  it('allows safe schemes and relative/anchor URLs', () => {
    expect(sanitizeHref('https://example.com')).toBe('https://example.com');
    expect(sanitizeHref('http://example.com')).toBe('http://example.com');
    expect(sanitizeHref('mailto:a@b.com')).toBe('mailto:a@b.com');
    expect(sanitizeHref('tel:+15551234')).toBe('tel:+15551234');
    expect(sanitizeHref('/relative/path')).toBe('/relative/path');
    expect(sanitizeHref('#anchor')).toBe('#anchor');
    expect(sanitizeHref('page.html?x=1:2')).toBe('page.html?x=1:2');
  });

  it('returns empty for empty/nullish input', () => {
    expect(sanitizeHref('')).toBe('');
    expect(sanitizeHref(null)).toBe('');
    expect(sanitizeHref(undefined)).toBe('');
  });
});

describe('sanitizeImageSrc', () => {
  it('allows http/https/data:image and relative, rejects javascript/non-image data', () => {
    expect(sanitizeImageSrc('https://x/y.png')).toBe('https://x/y.png');
    expect(sanitizeImageSrc('data:image/png;base64,AAAA')).toBe('data:image/png;base64,AAAA');
    expect(sanitizeImageSrc('/a.png')).toBe('/a.png');
    expect(sanitizeImageSrc('javascript:alert(1)')).toBe('');
    expect(sanitizeImageSrc('data:text/html,<script>')).toBe('');
  });
});

describe('safeCssColor', () => {
  it('accepts hex/functional/keyword/modern colors', () => {
    expect(safeCssColor('#1a73e8')).toBe('#1a73e8');
    expect(safeCssColor('rgb(1,2,3)')).toBe('rgb(1,2,3)');
    expect(safeCssColor('red')).toBe('red');
    expect(safeCssColor('var(--brand)')).toBe('var(--brand)');
  });

  it('rejects injection payloads', () => {
    expect(safeCssColor('red; background: url(//evil)')).toBeUndefined();
    expect(safeCssColor('url(//evil)')).toBeUndefined();
    expect(safeCssColor('')).toBeUndefined();
  });
});

describe('safeFontFamily', () => {
  it('accepts font lists, rejects injection', () => {
    expect(safeFontFamily('Arial, sans-serif')).toBe('Arial, sans-serif');
    expect(safeFontFamily('Arial; background: url(//evil)')).toBeUndefined();
    expect(safeFontFamily('url(//evil)')).toBeUndefined();
  });
});

describe('safeFontName', () => {
  it('keeps non-ASCII and punctuated font names', () => {
    for (const name of ['宋体', '游ゴシック', 'ＭＳ 明朝', 'Arial Narrow', 'Cambria Math', 'AT&T Garamond', "Book's Roman", 'Sylfaen.2']) expect(safeFontName(name)).toBe(name);
    expect(safeFontName('  Arial  ')).toBe('Arial');
  });

  it('rejects empty, long, control and injecting names', () => {
    for (const bad of ['', '   ', null, undefined, 'a;b', 'a{b}', 'a(b)', 'a<b>', 'a\\b', 'a:b', 'url(//evil)', 'a\u0000b', 'a\nb'.replace('\n', '\u0007'), 'x'.repeat(101)]) expect(safeFontName(bad as never)).toBeUndefined();
  });
});
