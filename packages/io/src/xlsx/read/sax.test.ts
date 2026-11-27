import { scanXml, type SaxAttrs } from './sax';

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
