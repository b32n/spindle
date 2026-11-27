import { MAX_ATTRIBUTES, scanXml, type SaxAttrs } from './sax';

type Event = ['open', string, SaxAttrs, boolean] | ['close', string] | ['text', string];

function scan(xml: string, maxDepth?: number): { ok: boolean; events: Event[] } {
  const events: Event[] = [];
  const ok = scanXml(
    xml,
    {
      open: (name, attrs, selfClosing) => events.push(['open', name, attrs, selfClosing]),
      close: (name) => events.push(['close', name]),
      text: (text) => events.push(['text', text]),
    },
    maxDepth
  );
  return { ok, events };
}

const names = (events: Event[]) => events.map((e) => (e[0] === 'text' ? `"${e[1]}"` : `${e[0]} ${e[1]}`));

describe('scanXml', () => {
  it('reports start tags, text and end tags in order', () => {
    const { ok, events } = scan('<a><b x="1">hi</b><c/></a>');
    expect(ok).toBe(true);
    expect(names(events)).toEqual(['open a', 'open b', '"hi"', 'close b', 'open c', 'close c', 'close a']);
    expect(events[1][2]).toEqual({ x: '1' });
    expect(events[4][3]).toBe(true);
  });

  it('drops prefixes from element and attribute names, and skips xmlns attributes', () => {
    const { events } = scan('<x:a xmlns:x="urn:x" xmlns="urn:y" r:id="rId1" x:id="other" id="plain"><x:b/></x:a>');
    expect(events[0][1]).toBe('a');
    expect(events[0][2]).toEqual({ id: 'plain' });
    expect(scan('<a r:id="rId1" id="plain"/>').events[0][2]).toEqual({ id: 'plain' });
    expect(scan('<a id="plain" r:id="rId1"/>').events[0][2]).toEqual({ id: 'plain' });
    expect(scan('<a r:id="rId1"/>').events[0][2]).toEqual({ id: 'rId1' });
    expect(names(events)).toEqual(['open a', 'open b', 'close b', 'close a']);
  });

  it('reads single-quoted attributes and spaces around the equals sign, and decodes entities', () => {
    const { events } = scan("<a p='x &amp; y' q = \"&lt;1&gt;&#65;&#x42;\" r=\"&bogus;\"/>");
    expect(events[0][2]).toEqual({ p: 'x & y', q: '<1>AB', r: '&bogus;' });
  });

  it('decodes entities in text and keeps CDATA as written', () => {
    const { events } = scan('<a>1 &lt; 2<![CDATA[ <raw> & ]]></a>');
    expect(events.filter((e) => e[0] === 'text').map((e) => e[1])).toEqual(['1 < 2', ' <raw> & ']);
  });

  it('skips the declaration, comments and a DOCTYPE with an internal subset, without expanding entities', () => {
    const xml = '<?xml version="1.0"?><!-- c --><!DOCTYPE a [<!ENTITY e "boom"><!ENTITY f "&e;&e;">]><a>&f;</a>';
    const { ok, events } = scan(xml);
    expect(ok).toBe(true);
    expect(events.filter((e) => e[0] === 'text').map((e) => e[1])).toEqual(['&f;']);
    expect(scan('<!DOCTYPE a><a/>').ok).toBe(true);
  });

  it('ignores text outside the root element', () => {
    const { ok, events } = scan(' \n<a/>\n trailing ');
    expect(ok).toBe(true);
    expect(names(events)).toEqual(['open a', 'close a']);
  });

  it('keeps attribute names like Object.prototype members as ordinary attributes', () => {
    const { events } = scan('<a __proto__="p" constructor="c" toString="t" hasOwnProperty="h"/>');
    const attrs = events[0][2] as SaxAttrs;
    expect(Object.getPrototypeOf(attrs)).toBeNull();
    expect(Object.keys(attrs).sort()).toEqual(['__proto__', 'constructor', 'hasOwnProperty', 'toString']);
    expect(attrs.__proto__).toBe('p');
    expect(({} as Record<string, unknown>).constructor).toBe(Object);
    expect((scan('<a/>').events[0][2] as SaxAttrs).constructor).toBeUndefined();
  });

  it.each([
    ['an unclosed element', '<a><b>'],
    ['a stray end tag', '<a></b></a>'],
    ['an end tag with nothing open', '</a>'],
    ['a mismatched end tag', '<a><b></a></b>'],
    ['an unterminated declaration', '<?xml version="1.0"'],
    ['an unterminated comment', '<a><!-- oops </a>'],
    ['an unterminated CDATA', '<a><![CDATA[ oops </a>'],
    ['an unterminated DOCTYPE', '<!DOCTYPE a [ <!ENTITY e "x"'],
    ['an unterminated DOCTYPE subset', '<!DOCTYPE a [ <!ENTITY e "x">'],
    ['an unterminated DOCTYPE without a subset', '<!DOCTYPE a'],
    ['a start tag cut short', '<a><b x="1"'],
    ['a start tag with a bad attribute', '<a x=1/>'],
    ['an end tag cut short', '<a></a'],
    ['a bare less-than sign', '<a>< </a>'],
    ['a less-than sign at the very end', '<a><'],
  ])('answers false for %s, after reporting what precedes it', (_name, xml) => {
    expect(scan(xml).ok).toBe(false);
  });

  it('reports what precedes the damage', () => {
    const { ok, events } = scan('<a><b>one</b><b>two</c></a>');
    expect(ok).toBe(false);
    expect(names(events)).toEqual(['open a', 'open b', '"one"', 'close b', 'open b', '"two"']);
  });

  it('caps the nesting depth', () => {
    const deep = (n: number) => '<a>'.repeat(n) + '</a>'.repeat(n);
    expect(scan(deep(10), 10).ok).toBe(true);
    expect(scan(deep(11), 10).ok).toBe(false);
    expect(scan(deep(300)).ok).toBe(false);
    expect(scan(deep(256)).ok).toBe(true);
  });

  it('ends the scan where the handler says it is halted, and answers true', () => {
    const events: string[] = [];
    let stop = false;
    const ok = scanXml('<a><b/><b/><b/></a>', {
      open: (name) => {
        events.push(`open ${name}`);
        if (events.length === 3) stop = true;
      },
      close: () => undefined,
      text: () => undefined,
      halted: () => stop,
    });
    expect(ok).toBe(true);
    expect(events).toEqual(['open a', 'open b', 'open b']);
  });
});

describe('scanXml: declarations and their brackets', () => {
  it('looks for the internal subset of a declaration only before its own `>`', () => {
    // the `[` after the first `>` is text, not the start of a subset: the declaration ends at its own `>`
    const { ok, events } = scan('<a><!x>[<b/></a>');
    expect(ok).toBe(true);
    expect(names(events)).toEqual(['open a', '"["', 'open b', 'close b', 'close a']);
    expect(scan('<a><!x><b/>]</a>').ok).toBe(true);
  });

  it('skips a subset that holds `>` and `[`, up to the `]>` that ends it', () => {
    const { ok, events } = scan('<!DOCTYPE a [<!ENTITY x "a>[b">]><a>t</a>');
    expect(ok).toBe(true);
    expect(names(events)).toEqual(['open a', '"t"', 'close a']);
  });

  it('is damage when the subset or the declaration is never closed', () => {
    expect(scan('<!DOCTYPE a [<!ENTITY x "1"><a/>').ok).toBe(false);
    expect(scan('<!DOCTYPE a [<!ENTITY x "1">]').ok).toBe(false);
    expect(scan('<a><!DOCTYPE a').ok).toBe(false);
  });
});

// Hostile shapes of about a million items each (a file of a few KB compresses to this): every one must be read in time that grows
// with the size of the text. The quadratic form of the first of them (an unbounded `indexOf('[')` per declaration) took 122 seconds
// for 1.6 million of them; the limit here is a hundredth of that. `ok` is what the scanner says of the text.
describe('scanXml: hostile shapes of a million items', () => {
  const N = 1_000_000;
  const open = '<worksheet xmlns="x">';
  const shapes: Array<[name: string, xml: () => string, ok: boolean]> = [
    ['1.6 million `<!a>` declarations with no `[` after them', () => `${open}${'<!a>'.repeat(1_600_000)}</worksheet>`, true],
    ['a million `<!DOCTYPE a>` declarations', () => `${open}${'<!DOCTYPE a>'.repeat(N)}</worksheet>`, true],
    ['a million `<!a>` declarations, a `[` at the very end', () => `${open}${'<!a>'.repeat(N)}</worksheet>[`, true],
    ['a million declarations with a subset each', () => `${open}${'<!a [x]>'.repeat(N)}</worksheet>`, true],
    ['a million processing instructions', () => `${open}${'<?x?>'.repeat(N)}</worksheet>`, true],
    ['a million empty comments', () => `${open}${'<!---->'.repeat(N)}</worksheet>`, true],
    ['a million CDATA sections', () => `${open}${'<![CDATA[]]>'.repeat(N)}</worksheet>`, true],
    ['a million unterminated comment openers', () => `${open}${'<!--'.repeat(N)}`, false],
    ['a million unterminated processing instructions', () => `${open}${'<?x'.repeat(N)}`, false],
    ['a million unterminated CDATA openers', () => `${open}${'<![CDATA['.repeat(N)}`, false],
    ['a million unterminated declarations', () => `${open}${'<!a'.repeat(N)}`, false],
    ['a million unterminated subsets', () => `${open}${'<!a [>'.repeat(N)}`, false],
    ['a million `]` after a subset opener', () => `${open}<!a [${']'.repeat(N)}`, false],
    ['a million `<`', () => `${open}${'<'.repeat(N)}`, false],
    ['a million stray end tags', () => `${open}${'</a>'.repeat(N)}`, false],
    ['a million attributes on one element (past the cap: damage)', () => `${open}<a ${'b="1" '.repeat(N)}/></worksheet>`, false],
    ['a million attributes without a value', () => `${open}<a ${'b '.repeat(N)}/></worksheet>`, false],
    ['a million unclosed attribute values', () => `${open}<a ${'b="1'.repeat(N)}/></worksheet>`, false],
    ['a million `xmlns` declarations on one element (past the cap: damage)', () => `${open}<a ${Array.from({ length: N }, (_, i) => `xmlns:p${i}="u"`).join(' ')}/></worksheet>`, false],
    ['an attribute value of 20 million characters', () => `${open}<a b="${'x'.repeat(20_000_000)}"/></worksheet>`, true],
    ['an attribute value of 2 million entities', () => `${open}<a b="${'&#65;'.repeat(2_000_000)}"/></worksheet>`, true],
    ['2 million entities in text', () => `${open}<a>${'&amp;'.repeat(2_000_000)}</a></worksheet>`, true],
    ['a million elements nested (the depth cap ends it)', () => `${open}${'<a>'.repeat(N)}`, false],
    ['4 million characters of whitespace in a tag', () => `${open}<a${' '.repeat(4_000_000)}/></worksheet>`, true],
    ['a million `<` in text', () => `${open}<a>${'x<'.repeat(N)}</a>`, false],
    ['a million `/` in a tag', () => `${open}<a ${'/'.repeat(N)}>`, false],
    ['a million `=` in a tag', () => `${open}<a ${'='.repeat(N)}>`, false],
  ];

  it.each(shapes)('%s', (_name, xml, ok) => {
    const text = xml();
    // (a quadratic scanner never ends on these: the test's own timeout is the failure, not a reading of the clock)
    const result = scanXml(text, { open: () => undefined, close: () => undefined, text: () => undefined });
    expect(result).toBe(ok);
  }, 120_000);

  it('reads 800,000 declarations (it was quadratic in them)', () => {
    expect(scanXml(`${open}${'<!a>'.repeat(800_000)}</worksheet>`, { open: () => undefined, close: () => undefined, text: () => undefined })).toBe(true);
  }, 120_000);
});

describe('scanXml: the attribute cap', () => {
  const attrs = (n: number, prefix = 'a'): string => Array.from({ length: n }, (_, i) => ` ${prefix}${i}="v"`).join('');

  it('reads an element of exactly MAX_ATTRIBUTES attributes', () => {
    const { ok, events } = scan(`<r><c${attrs(MAX_ATTRIBUTES)}/></r>`);
    expect(ok).toBe(true);
    expect(Object.keys(events[1][2] as SaxAttrs)).toHaveLength(MAX_ATTRIBUTES);
  });

  it('answers false for one attribute more, and the element is not reported (what came before stands)', () => {
    const { ok, events } = scan(`<r><b/><c${attrs(MAX_ATTRIBUTES + 1)}/><d/></r>`);
    expect(ok).toBe(false);
    expect(names(events)).toEqual(['open r', 'open b', 'close b']);
  });

  it('counts namespace declarations, which are not reported, and prefixed names too', () => {
    expect(scan(`<c${attrs(MAX_ATTRIBUTES, 'xmlns:p')}/>`).ok).toBe(true);
    expect(scan(`<c${attrs(MAX_ATTRIBUTES + 1, 'xmlns:p')}/>`).ok).toBe(false);
    expect(scan(`<c${attrs(MAX_ATTRIBUTES / 2, 'xmlns:p')}${attrs(MAX_ATTRIBUTES / 2 + 1, 'x:a')}/>`).ok).toBe(false);
  });

  it('counts per element, not per part: many elements of 1,024 each are read', () => {
    expect(scan(`<r>${`<c${attrs(MAX_ATTRIBUTES)}/>`.repeat(20)}</r>`).ok).toBe(true);
  });

  it('is the cap for repeats of one name as well (they cost a scan each, and the cap bounds that)', () => {
    expect(scan(`<c${' a="1"'.repeat(MAX_ATTRIBUTES + 1)}/>`).ok).toBe(false);
  });
});

describe('scanXml without a text handler', () => {
  it('reads the same structure and skips text and CDATA, whatever they hold', () => {
    const opened: string[] = [];
    const ok = scanXml('<a> &amp; <b><![CDATA[x<y]]></b>tail<c/></a>', { open: (name) => void opened.push(name), close: () => undefined });
    expect(ok).toBe(true);
    expect(opened).toEqual(['a', 'b', 'c']);
  });
});
