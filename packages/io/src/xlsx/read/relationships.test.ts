import { parseRelationships, relsOfKind, relsPart, resolveRelationshipTarget } from './relationships';
import { relsXml } from './test-support';

describe('resolveRelationshipTarget', () => {
  it.each([
    ['xl/workbook.xml', 'worksheets/sheet1.xml', 'xl/worksheets/sheet1.xml'], // Excel: relative
    ['xl/workbook.xml', '/xl/worksheets/sheet1.xml', 'xl/worksheets/sheet1.xml'], // openpyxl: absolute
    ['xl/worksheets/sheet1.xml', '../drawings/drawing1.xml', 'xl/drawings/drawing1.xml'],
    ['xl/worksheets/sheet1.xml', '/xl/drawings/commentsDrawing1.vml', 'xl/drawings/commentsDrawing1.vml'],
    ['xl/workbook.xml', './styles.xml', 'xl/styles.xml'],
    ['xl/workbook.xml', 'worksheets\\sheet1.xml', 'xl/worksheets/sheet1.xml'], // backslashes (some .NET writers)
    ['xl/workbook.xml', 'worksheets/My%20Sheet.xml', 'xl/worksheets/My Sheet.xml'],
    ['xl/workbook.xml', 'worksheets/100%.xml', 'xl/worksheets/100%.xml'], // not valid percent-encoding: kept
    ['xl/workbook.xml', '  styles.xml  ', 'xl/styles.xml'],
    ['xl/workbook.xml', '../../../../x.xml', 'x.xml'], // cannot climb above the root
    ['xl/workbook.xml', 'worksheets/sheet1.xml#a', 'xl/worksheets/sheet1.xml'], // a fragment is not part of the name
    ['xl/workbook.xml', 'worksheets/a%23b.xml', 'xl/worksheets/a#b.xml'], // an encoded # is
    ['', 'xl/workbook.xml', 'xl/workbook.xml'], // the package's own relationships
    ['', '/xl/workbook.xml', 'xl/workbook.xml'],
  ])('%s + %s -> %s', (owner, target, expected) => expect(resolveRelationshipTarget(owner, target)).toBe(expected));
});

describe('relsPart', () => {
  it('names the relationships part of a part', () => expect(relsPart('xl/workbook.xml')).toBe('xl/_rels/workbook.xml.rels'));
});

describe('parseRelationships', () => {
  const rels = parseRelationships(
    (
      relsXml([
        ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
        ['rId1', 'worksheet', 'worksheets/dup.xml'],
        ['rId2', 'hyperlink', 'https://example.com/a b', 'External'],
        ['rId3', 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', '/xl/workbook.xml'],
        ['rId4', 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties', 'docProps/core.xml', 'external'],
      ]).replace('</Relationships>', '<Relationship Type="x" Target="no-id"/><Relationship Id="rId9" Type="x"/><Relationship Id="rId8" Target="no-type.xml"/></Relationships>')
    ),
    'xl/workbook.xml',
  );
  it('resolves internal targets, keeps external ones as written, skips entries without id or target and keeps the first of a repeated id', () => {
    expect(rels.map((r) => [r.id, r.kind, r.target, r.external])).toEqual([
      ['rId1', 'worksheet', 'xl/worksheets/sheet1.xml', false],
      ['rId2', 'hyperlink', 'https://example.com/a b', true],
      ['rId3', 'officeDocument', 'xl/workbook.xml', false],
      ['rId4', 'core-properties', 'docProps/core.xml', true],
      ['rId8', '', 'xl/no-type.xml', false],
    ]);
    expect(rels[2].type).toBe('http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument');
  });
  it('filters by kind regardless of case', () => {
    expect(relsOfKind(rels, 'OFFICEDOCUMENT').map((r) => r.id)).toEqual(['rId3']);
    expect(relsOfKind(rels, 'nothing')).toEqual([]);
  });
  it('is empty for a missing part', () => expect(parseRelationships(undefined, 'a.xml')).toEqual([]));
  it('reads only direct Relationship children of the root, and reads the same under any prefix', () => {
    const text = '<r:Relationships xmlns:r="urn:x"><r:Relationship Id="a" Target="t.xml"/><Other><Relationship Id="b" Target="u.xml"/></Other></r:Relationships>';
    expect(parseRelationships(text, 'xl/workbook.xml').map((r) => r.id)).toEqual(['a']);
  });
});

describe('targets with a URI scheme', () => {
  const rels = parseRelationships(
    relsXml([['rId1', 'worksheet', 'file:///x.xml'], ['rId2', 'worksheet', 'MAILTO:a@b.c'], ['rId3', 'worksheet', 'C:\\x.xml'], ['rId4', 'worksheet', 'a.b:c/x.xml']]),
    'xl/workbook.xml',
  );
  it('are kept as written and counted as external, even without TargetMode', () => {
    expect(rels.slice(0, 2).map((r) => [r.target, r.external])).toEqual([['file:///x.xml', true], ['MAILTO:a@b.c', true]]);
  });
  it('a Windows drive letter is not a scheme (a one letter scheme is a path), a dotted scheme is', () => {
    expect(rels[2].external).toBe(false);
    expect(rels[3]).toMatchObject({ target: 'a.b:c/x.xml', external: true });
  });
});

describe('parseRelationships: problems and the entry limit', () => {
  const many = (n: number) => relsXml(Array.from({ length: n }, (_, i) => [`r${i}`, 'worksheet', `s${i}.xml`] as [string, string, string]));
  it('keeps the first entries up to the limit and reports it once', () => {
    const problems: string[] = [];
    const rels = parseRelationships(many(10), 'xl/workbook.xml', { max: 4, problem: (w) => problems.push(w) });
    expect(rels.map((r) => r.id)).toEqual(['r0', 'r1', 'r2', 'r3']);
    expect(problems).toEqual(['over']);
  });
  it('is not over at exactly the limit, and a repeated id does not count', () => {
    const problems: string[] = [];
    const text = many(4).replace('</Relationships>', '<Relationship Id="r0" Target="again.xml"/></Relationships>');
    expect(parseRelationships(text, 'xl/workbook.xml', { max: 4, problem: (w) => problems.push(w) })).toHaveLength(4);
    expect(problems).toEqual([]);
  });
  it('reads what precedes damage and reports it', () => {
    const problems: string[] = [];
    const text = many(2).replace('</Relationships>', '<Relationship Id="x" Target="y.xml"></Oops>');
    expect(parseRelationships(text, 'xl/workbook.xml', { problem: (w) => problems.push(w) })).toHaveLength(3);
    expect(problems).toEqual(['damaged']);
    expect(() => parseRelationships(text, 'xl/workbook.xml')).not.toThrow();
  });
});

describe('parseRelationships: the kinds kept and the package budget', () => {
  const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const text = relsXml([['a', 'hyperlink', 'https://x/', 'External'], ['b', 'image', 'i.png'], ['c', 'hyperlink', 'https://y/', 'External'], ['b', 'hyperlink', 'again.xml'], ['d', 'comments', 'c.xml']]);

  it('keeps what `keep` names by kind and id, and the others are read (they count for the part limit) but not kept; a repeated id keeps the first, whatever it was', () => {
    const asked: Array<[string, string]> = [];
    const keep = (kind: string, id: string) => {
      asked.push([kind, id]);
      return kind === 'hyperlink' || id === 'd';
    };
    expect(parseRelationships(text, 'xl/worksheets/sheet1.xml', { keep }).map((r) => r.id)).toEqual(['a', 'c', 'd']);
    // the repeated `b` is not asked again: the first `b` (an image, not kept) stands
    expect(asked).toEqual([['hyperlink', 'a'], ['image', 'b'], ['hyperlink', 'c'], ['comments', 'd']]);
    expect(parseRelationships(text, 'xl/worksheets/sheet1.xml', { max: 2, keep, problem: () => undefined }).map((r) => r.id)).toEqual(['a']);
  });

  it('draws each entry read on the budget it is given, shared with the next part, and ends the read with problem("budget") when it is spent', () => {
    const budget = { left: 4 };
    const problems: string[] = [];
    expect(parseRelationships(text, 'x.xml', { budget, keep: (k) => k === 'hyperlink', problem: (w) => problems.push(w) }).map((r) => r.id)).toEqual(['a', 'c']);
    expect(problems).toEqual([]);
    expect(budget.left).toBe(0);
    expect(parseRelationships(text, 'x.xml', { budget, problem: (w) => problems.push(w) })).toEqual([]);
    expect(problems).toEqual(['budget']);
  });

  it('gives the entries of one type URI the same string, and keeps the kind as the last segment of it', () => {
    const [a, b] = parseRelationships(relsXml([['a', 'hyperlink', 'https://x/', 'External'], ['b', 'hyperlink', 'https://y/', 'External']]), 'x.xml');
    expect([a.type, a.kind]).toEqual([`${OD}/hyperlink`, 'hyperlink']);
    expect(b.type).toBe(a.type);
  });
});
