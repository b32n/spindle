import { decodeEntities, ENTITY_WINDOW, descendants, isMain, isOn, isTrue, kid, kids, parseXml, walk, type XmlElement } from './xml';

const MAIN = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

describe('decodeEntities', () => {
  it('decodes the predefined and numeric references in one pass', () => {
    expect(decodeEntities('a &lt;b&gt; &amp; &quot;c&quot; &apos;d&apos; &#65;&#x42;')).toBe('a <b> & "c" \'d\' AB');
  });

  it('does not decode twice', () => {
    expect(decodeEntities('&amp;lt;')).toBe('&lt;');
  });

  it('leaves text without an ampersand, unknown entities and invalid code points alone', () => {
    expect(decodeEntities('plain')).toBe('plain');
    expect(decodeEntities('&lol;')).toBe('&lol;');
    expect(decodeEntities('&#0;')).toBe('&#0;');
    expect(decodeEntities('&#xD800;')).toBe('&#xD800;');
    expect(decodeEntities('&#x110000;')).toBe('&#x110000;');
    expect(decodeEntities('&#128512;')).toBe('\u{1F600}');
  });
});

/** The decoding as one `replace` over the whole text: what the windowed one must equal. */
const WHOLE = (s: string): string =>
  s.replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|(lt|gt|amp|quot|apos));/g, (whole, dec?: string, hex?: string, named?: string) => {
    if (named) return { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }[named] as string;
    const code = dec !== undefined ? Number(dec) : parseInt(hex as string, 16);
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : whole;
  });

describe('decodeEntities in windows', () => {
  const table: Array<[text: string, decoded: string]> = [
    ['&lt;&gt;&amp;&quot;&apos;', '<>&"\''],
    ['&#65;&#x41;&#X41;', 'AA&#X41;'],
    ['&#0065;&#x000041;', 'AA'],
    ['&#128512;&#x1F600;', '\u{1F600}\u{1F600}'],
    ['&#xD83D;&#xDE00;', '&#xD83D;&#xDE00;'],
    ['&#55357;&#56832;', '&#55357;&#56832;'],
    ['&#0;&#x110000;&#1114112;', '&#0;&#x110000;&#1114112;'],
    ['&amp;lt; &amp;#65;', '&lt; &#65;'],
    ['&lol; &amp &#; &#x; & &;', '&lol; &amp &#; &#x; & &;'],
    ['\u{1F600}&amp;\u{1F600}', '\u{1F600}&\u{1F600}'],
    ['&&amp;&amp;&', '&&&&'],
  ];

  it.each(table)('%j decodes to the expected text at every window size', (text, decoded) => {
    expect(WHOLE(text)).toBe(decoded);
    for (let size = 1; size <= text.length + 2; size++) expect(decodeEntities(text, size)).toBe(decoded);
    expect(decodeEntities(text)).toBe(decoded);
  });

  it('equals the whole-text decoding for random entity-dense text, whatever the window (references split at every position)', () => {
    const pieces = ['&', ';', '#', 'x', 'X', '6', '5', 'A', 'a', 'lt', 'amp', '&#65;', '&#x1F600;', '&#xD83D;', '&amp;', '&lt;x', '\u{1F600}', ' ', '&#00', '&;'];
    let seed = 12345;
    const next = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % n;
    };
    for (let round = 0; round < 300; round++) {
      let text = '';
      for (let i = next(40); i >= 0; i--) text += pieces[next(pieces.length)];
      const expected = WHOLE(text);
      for (const size of [1, 2, 3, 5, 8, 13]) expect(decodeEntities(text, size)).toBe(expected);
    }
  });

  it('extends a window over a stretch with no semicolon, to the next one or to the end', () => {
    const noEnd = `${'&'.repeat(100)}x`;
    expect(decodeEntities(noEnd, 10)).toBe(noEnd);
    const late = `${'&'.repeat(50)}#65;tail`;
    expect(decodeEntities(late, 10)).toBe(`${'&'.repeat(49)}Atail`);
    expect(decodeEntities(`${'a'.repeat(30)}&amp;${'b'.repeat(30)}`, 7)).toBe(`${'a'.repeat(30)}&${'b'.repeat(30)}`);
  });

  it('decodes a part far larger than a window, and a window-sized text, as the whole-text decoding does', () => {
    const unit = '&amp;&#65;&lt;x';
    const text = unit.repeat(ENTITY_WINDOW);
    expect(text.length).toBeGreaterThan(10 * ENTITY_WINDOW);
    const out = decodeEntities(text);
    expect(out.length).toBe(ENTITY_WINDOW * 4);
    expect(out.startsWith('&A<x&A<x')).toBe(true);
    expect(out.endsWith('&A<x')).toBe(true);
    const exactly = 'a'.repeat(ENTITY_WINDOW - 5) + '&amp;';
    expect(decodeEntities(exactly)).toBe(`${'a'.repeat(ENTITY_WINDOW - 5)}&`);
    expect(decodeEntities(`${exactly}&lt;`)).toBe(`${'a'.repeat(ENTITY_WINDOW - 5)}&<`);
  });
});

describe('parseXml', () => {
  it('builds an element tree with local names, decoded attributes and text', () => {
    const root = parseXml(`<?xml version="1.0"?><a ${MAIN} x="1&lt;2"><b>t&amp;u</b><c/></a>`) as XmlElement;
    expect(root.name).toBe('a');
    expect(root.ns).toBe('main');
    expect(root.attrs).toEqual({ x: '1<2' });
    expect(root.children.map((c) => c.name)).toEqual(['b', 'c']);
    expect(root.children[0].text).toBe('t&u');
    expect(root.text).toBe('');
  });

  it('resolves namespaces whatever the prefix, including the strict URI and inherited declarations', () => {
    const root = parseXml(
      '<x:a xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:y="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:z="urn:other">' +
        '<y:b><x:c/><z:d/></y:b><e/></x:a>'
    ) as XmlElement;
    expect(root.ns).toBe('main');
    const b = root.children[0];
    expect(b.ns).toBe('x14');
    expect(b.children.map((c) => c.ns)).toEqual(['main', 'urn:other']);
    expect(root.children[1].ns).toBe('');
    expect(isMain(root)).toBe(true);
    expect(isMain(root.children[1])).toBe(true);
    expect(isMain(b)).toBe(false);
  });

  it('maps the known namespaces to short codes', () => {
    const root = parseXml(
      '<a xmlns="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><b xmlns="http://schemas.microsoft.com/office/excel/2006/main"/>' +
        '<c xmlns="http://schemas.openxmlformats.org/markup-compatibility/2006"/><d xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>' +
        '<e xmlns="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"/><f xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/></a>'
    ) as XmlElement;
    expect([root.ns, ...root.children.map((c) => c.ns)]).toEqual(['x15', 'xm', 'mc', 'xdr', 'xdr', 'rel']);
  });

  it('keeps an unprefixed attribute over a prefixed one of the same local name, in either order', () => {
    const a = parseXml('<a xmlns:r="urn:r" id="1" r:id="2"/>') as XmlElement;
    const b = parseXml('<a xmlns:r="urn:r" r:id="2" id="1"/>') as XmlElement;
    const c = parseXml('<a xmlns:r="urn:r" r:id="2" r:id2="3"/>') as XmlElement;
    expect(a.attrs).toEqual({ id: '1' });
    expect(b.attrs).toEqual({ id: '1' });
    expect(c.attrs).toEqual({ id: '2', id2: '3' });
  });

  it('never expands entities, and ignores a DOCTYPE with an internal subset', () => {
    const bomb =
      '<?xml version="1.0"?><!DOCTYPE a [<!ENTITY l "lol"><!ENTITY l2 "&l;&l;&l;"><!ENTITY x SYSTEM "file:///etc/passwd">]>' +
      '<a v="&l2;">&l2;&x;</a>';
    const root = parseXml(bomb) as XmlElement;
    expect(root.attrs.v).toBe('&l2;');
    expect(root.text).toBe('&l2;&x;');
  });

  it('handles a DOCTYPE without a subset and refuses one it cannot delimit', () => {
    expect((parseXml('<!DOCTYPE a><a/>') as XmlElement).name).toBe('a');
    expect((parseXml('<!DOCTYPE a [<!ENTITY l "x">]  ><a/>') as XmlElement).name).toBe('a');
    expect(parseXml('<!DOCTYPE a [<!ENTITY l "x"><a/>')).toBeNull();
    expect(parseXml('<!DOCTYPE a')).toBeNull();
  });

  it('returns null for empty, garbage, too deeply nested and unparseable input, and never throws', () => {
    expect(parseXml('')).toBeNull();
    expect(parseXml('just text')).toBeNull();
    expect(parseXml('<?xml version="1.0"?>')).toBeNull();
    expect(parseXml(`${'<a>'.repeat(500)}x`)).toBeNull();
  });

  it('reads CDATA as text and skips comments and processing instructions', () => {
    const root = parseXml('<?xml version="1.0"?><!-- c --><a><![CDATA[x<y]]><?pi data?></a>') as XmlElement;
    expect(root.text).toBe('x<y');
    expect(root.children).toEqual([]);
  });
});

describe('tree helpers', () => {
  const root = parseXml(
    '<a xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      '<b><c n="1"/></b><b/><mc:AlternateContent><mc:Choice><c n="2"/></mc:Choice><mc:Fallback><c n="3"/></mc:Fallback></mc:AlternateContent>' +
      '<Fallback><c n="4"/></Fallback></a>'
  ) as XmlElement;

  it('kids and kid select direct children by local name and tolerate undefined', () => {
    expect(kids(root, 'b')).toHaveLength(2);
    expect(kid(root, 'b')).toBe(root.children[0]);
    expect(kids(undefined, 'b')).toEqual([]);
    expect(kid(undefined, 'b')).toBeUndefined();
  });

  it('walk and descendants visit in document order and skip mc:Fallback branches only', () => {
    expect(descendants(root, 'c').map((c) => c.attrs.n)).toEqual(['1', '2', '4']);
    expect(descendants(root, 'c', 'mc')).toEqual([]);
    expect(descendants(root, 'AlternateContent', 'mc')).toHaveLength(1);
    const seen: string[] = [];
    walk(root, (e) => seen.push(e.name));
    expect(seen).toEqual(['b', 'c', 'b', 'AlternateContent', 'Choice', 'c', 'Fallback', 'c']);
    expect(() => walk(undefined, () => undefined)).not.toThrow();
  });

  it('reads OOXML booleans', () => {
    expect([isTrue('1'), isTrue('true'), isTrue('0'), isTrue('false'), isTrue(undefined)]).toEqual([true, true, false, false, false]);
    expect([isOn('1'), isOn('true'), isOn('0'), isOn('false'), isOn(undefined)]).toEqual([true, true, false, false, true]);
  });
});

describe('relationship attributes', () => {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  it('keeps r:id apart from an unprefixed id, whatever the prefix is called', () => {
    const root = parseXml(`<a xmlns:rel="${R}" xmlns:o="urn:other"><b id="256" rel:id="rId4" o:id="x"/><c id="1"/><d xmlns:q="${R}" q:embed="rId9"/></a>`) as XmlElement;
    const [b, c, d] = root.children;
    expect(b.attrs.id).toBe('256');
    expect(b.rattrs).toEqual({ id: 'rId4' });
    expect(c.rattrs).toBeUndefined();
    expect(d.rattrs).toEqual({ embed: 'rId9' });
  });

  it('knows the strict namespace and an undeclared prefix is not one', () => {
    // network-allow: XML namespace identifier of the ISO strict variant, never requested
    const root = parseXml('<a xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><b r:id="rId1" z:id="no"/></a>') as XmlElement;
    expect(root.children[0].rattrs).toEqual({ id: 'rId1' });
  });
});
