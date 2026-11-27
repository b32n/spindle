import { parseRelationships, relsOfKind, relsPart, resolveRelationshipTarget } from './relationships';
import { el, relsXml } from './test-support';

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
    el(
      relsXml([
        ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
        ['rId1', 'worksheet', 'worksheets/dup.xml'],
        ['rId2', 'hyperlink', 'https://example.com/a b', 'External'],
        ['rId3', 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument', '/xl/workbook.xml'],
        ['rId4', 'http://schemas.openxmlformats.org/officedocument/2006/relationships/metadata/core-properties', 'docProps/core.xml', 'external'],
      ]).replace('</Relationships>', '<Relationship Type="x" Target="no-id"/><Relationship Id="rId9" Type="x"/><Relationship Id="rId8" Target="no-type.xml"/></Relationships>'),
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
  it('is empty for a missing root', () => expect(parseRelationships(undefined, 'a.xml')).toEqual([]));
});

describe('targets with a URI scheme', () => {
  const rels = parseRelationships(
    el(relsXml([['rId1', 'worksheet', 'file:///x.xml'], ['rId2', 'worksheet', 'MAILTO:a@b.c'], ['rId3', 'worksheet', 'C:\\x.xml'], ['rId4', 'worksheet', 'a.b:c/x.xml']])),
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
