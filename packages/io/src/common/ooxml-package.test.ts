import JSZip from 'jszip';
import { OoxmlPackage, PACKAGE_LIMITS, relsPart, resolveTarget } from './ooxml-package';

const RELS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';

async function open(files: Record<string, string>): Promise<OoxmlPackage> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  return new OoxmlPackage(await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' })));
}

describe('resolveTarget', () => {
  it.each([
    ['ppt/slides/slide1.xml', '../media/a.png', 'ppt/media/a.png'],
    ['ppt/slides/slide1.xml', 'sub/./b.xml', 'ppt/slides/sub/b.xml'],
    ['ppt/slides/slide1.xml', '/ppt/theme/theme1.xml', 'ppt/theme/theme1.xml'],
    ['ppt/presentation.xml', 'slides//slide1.xml', 'ppt/slides/slide1.xml'],
    ['ppt/slides/slide1.xml', '../../../../x.xml', 'x.xml'],
    ['presentation.xml', 'a.xml', 'a.xml'],
  ])('%s + %s -> %s', (base, target, expected) => {
    expect(resolveTarget(base, target)).toBe(expected);
  });
});

describe('relsPart', () => {
  it('puts the relationships of a part in the _rels folder next to it', () => {
    expect(relsPart('ppt/slides/slide1.xml')).toBe('ppt/slides/_rels/slide1.xml.rels');
    expect(relsPart('document.xml')).toBe('_rels/document.xml.rels');
  });
});

describe('OoxmlPackage', () => {
  it('lists file entries but not directories, and says which exist', async () => {
    const zip = new JSZip();
    zip.file('a/b.xml', '<x/>');
    zip.folder('empty');
    const pkg = new OoxmlPackage(await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' })));
    expect(pkg.names).toEqual(['a/b.xml']);
    expect(pkg.has('a/b.xml')).toBe(true);
    expect(pkg.has('empty/')).toBe(false);
    expect(pkg.has('nope.xml')).toBe(false);
  });

  it('reads text and parsed XML once, counting the inflated bytes', async () => {
    const pkg = await open({ 'a.xml': '<root><c/></root>', 'bad.xml': '<root' });
    expect(await pkg.text('a.xml')).toBe('<root><c/></root>');
    expect(pkg.read).toBe(17);
    expect((await pkg.xml('a.xml'))?.children).toHaveLength(1);
    expect((await pkg.xml('a.xml'))?.children).toHaveLength(1);
    expect(pkg.read).toBe(34); // text() above plus one parse: the second xml() call is cached
    expect(await pkg.xml('bad.xml')).toBeUndefined();
    expect(await pkg.xml('missing.xml')).toBeUndefined();
    expect(await pkg.text('missing.xml')).toBeUndefined();
  });

  it('reads relationships, resolving internal targets and keeping external ones, and caches them', async () => {
    const pkg = await open({
      'ppt/slides/_rels/slide1.xml.rels': `<Relationships ${RELS}><Relationship Id="rId1" Type="http://x/relationships/image" Target="../media/a.png"/><Relationship Id="rId2" Type="http://x/relationships/hyperlink" Target="https://e.com/" TargetMode="External"/><Relationship Type="plain"/><Relationship Id="rId4" Target="x.xml"/></Relationships>`,
    });
    const rels = await pkg.rels('ppt/slides/slide1.xml');
    expect(rels).toEqual([
      { id: 'rId1', type: 'image', target: 'ppt/media/a.png', external: false, raw: '../media/a.png' },
      { id: 'rId2', type: 'hyperlink', target: 'https://e.com/', external: true, raw: 'https://e.com/' },
      { id: '', type: 'plain', target: 'ppt/slides', external: false, raw: '' },
      { id: 'rId4', type: '', target: 'ppt/slides/x.xml', external: false, raw: 'x.xml' },
    ]);
    const before = pkg.read;
    expect(await pkg.rels('ppt/slides/slide1.xml')).toBe(rels);
    expect(pkg.read).toBe(before);
    expect(await pkg.rels('ppt/slides/slide9.xml')).toEqual([]);
  });

  it('records a part over the limit as skipped, and still reads the others', async () => {
    const saved = PACKAGE_LIMITS.partBytes;
    PACKAGE_LIMITS.partBytes = 50;
    try {
      const pkg = await open({ 'big.xml': `<a>${'x'.repeat(200)}</a>`, 'small.xml': '<a/>' });
      expect(await pkg.text('big.xml')).toBeUndefined();
      expect(await pkg.text('small.xml')).toBe('<a/>');
      expect(pkg.skipped).toEqual(['big.xml']);
    } finally {
      PACKAGE_LIMITS.partBytes = saved;
    }
  });
});
