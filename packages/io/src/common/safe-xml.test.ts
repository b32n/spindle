import { createSafeXmlParser, parseXmlTree, stripDoctype, XML_LIMITS } from './safe-xml';

const nested = (depth: number) => '<a>'.repeat(depth) + '</a>'.repeat(depth);

describe('stripDoctype', () => {
  it('leaves XML without a DOCTYPE alone', () => {
    expect(stripDoctype('<a/>')).toBe('<a/>');
  });
  it('removes a simple DOCTYPE', () => {
    expect(stripDoctype('<!DOCTYPE a><a/>')).toBe('<a/>');
  });
  it('removes a DOCTYPE with an internal subset', () => {
    expect(stripDoctype('<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>')).toBe('<a>&x;</a>');
  });
  it('removes a DOCTYPE whose "[" belongs to a later part of the document', () => {
    expect(stripDoctype('<!DOCTYPE a><a>[</a>')).toBe('<a>[</a>');
  });
  it.each([['<!DOCTYPE a [<!ENTITY x "y">'], ['<!DOCTYPE a']])('cannot delimit %p', (xml) => {
    expect(stripDoctype(xml)).toBeNull();
  });
});

describe('parseXmlTree', () => {
  it('parses to an order-preserving tree', () => {
    const tree = parseXmlTree('<a><b/><c/></a>') as Record<string, unknown>[];
    expect(Object.keys(tree[0])).toEqual(['a']);
    expect((tree[0].a as Record<string, unknown>[]).map((n) => Object.keys(n)[0])).toEqual(['b', 'c']);
  });

  it('never expands entities: a billion-laughs document parses to its small, unexpanded form', () => {
    const laughs = '<!DOCTYPE lolz [<!ENTITY lol "lol">' + [1, 2, 3, 4, 5, 6, 7, 8, 9]
      .map((i) => `<!ENTITY lol${i} "${`&lol${i - 1 || ''};`.repeat(10)}">`).join('') + ']><a>&lol9;</a>';
    const tree = parseXmlTree(laughs);
    expect(JSON.stringify(tree).length).toBeLessThan(100);
  });

  it('does not resolve an external entity', () => {
    const tree = parseXmlTree('<!DOCTYPE a [<!ENTITY x SYSTEM "file:///etc/passwd">]><a>&x;</a>');
    expect(JSON.stringify(tree)).not.toContain('root:');
  });

  it('answers null for an undelimited DOCTYPE and for a document nested too deeply', () => {
    expect(parseXmlTree('<!DOCTYPE a [')).toBeNull();
    expect(parseXmlTree(nested(10_000))).toBeNull();
  });

  it('reads nesting up to the cap (a bit over the cap is the parser\'s own off-by-one) and refuses a deep bomb', () => {
    expect(parseXmlTree(nested(XML_LIMITS.maxDepth - 1))).not.toBeNull();
    expect(parseXmlTree(nested(XML_LIMITS.maxDepth * 2))).toBeNull();
  });

  it('takes a parser with another depth cap', () => {
    const shallow = createSafeXmlParser(3);
    expect(parseXmlTree(nested(2), shallow)).not.toBeNull();
    expect(parseXmlTree(nested(10), shallow)).toBeNull();
  });
});
