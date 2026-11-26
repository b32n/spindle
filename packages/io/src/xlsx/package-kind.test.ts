import { isEncryptedOle, isKnownPart, isOleFile } from './package-kind';

const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const utf16 = (s: string) => [...s].flatMap((c) => [c.charCodeAt(0), 0]);

describe('OLE detection', () => {
  it('recognises the compound-file signature and nothing shorter or different', () => {
    expect(isOleFile(new Uint8Array(OLE))).toBe(true);
    expect(isOleFile(new Uint8Array(OLE.slice(0, 7)))).toBe(false);
    expect(isOleFile(new Uint8Array([...OLE.slice(0, 7), 0]))).toBe(false);
    expect(isOleFile(new Uint8Array(0))).toBe(false);
    expect(isOleFile(new TextEncoder().encode('PK\u0003\u0004'))).toBe(false);
  });

  it('tells an encrypted package from a legacy workbook by its stream names', () => {
    expect(isEncryptedOle(new Uint8Array([...OLE, 1, 2, ...utf16('EncryptedPackage'), 3]))).toBe(true);
    expect(isEncryptedOle(new Uint8Array([...OLE, ...utf16('EncryptionInfo')]))).toBe(true);
    expect(isEncryptedOle(new Uint8Array([...OLE, ...utf16('Workbook')]))).toBe(false);
    // a near miss: the right prefix, a wrong tail, and a name stored as single bytes
    expect(isEncryptedOle(new Uint8Array([...OLE, ...utf16('EncryptedPackagX')]))).toBe(false);
    expect(isEncryptedOle(new Uint8Array([...OLE, ...new TextEncoder().encode('EncryptedPackage')]))).toBe(false);
    expect(isEncryptedOle(new Uint8Array([...OLE, ...utf16('Encrypted')]))).toBe(false);
    expect(isEncryptedOle(new Uint8Array(0))).toBe(false);
  });
});

describe('isKnownPart', () => {
  it.each([
    '[Content_Types].xml',
    '_rels/.rels',
    'xl/_rels/workbook.xml.rels',
    'xl/worksheets/_rels/sheet1.xml.rels',
    'docProps/core.xml',
    'xl/workbook.xml',
    'xl/styles.xml',
    'xl/sharedStrings.xml',
    'xl/calcChain.xml',
    'xl/metadata.xml',
    'xl/worksheets/sheet1.xml',
    'xl/theme/theme1.xml',
    'xl/printerSettings/printerSettings1.bin',
    'xl/comments1.xml',
    'xl/comments/comment1.xml',
    'xl/drawings/drawing1.xml',
    'xl/drawings/vmlDrawing1.vml',
    'xl/charts/chart1.xml',
    'xl/media/image1.png',
    'xl/threadedComments/threadedComment1.xml',
    'xl/persons/person.xml',
    'xl/tables/table1.xml',
    'xl/pivotTables/pivotTable1.xml',
    'xl/pivotCache/pivotCacheDefinition1.xml',
    'xl/slicers/slicer1.xml',
    'xl/slicerCaches/slicerCache1.xml',
    'xl/externalLinks/externalLink1.xml',
    'xl/ctrlProps/ctrlProp1.xml',
    'xl/activeX/activeX1.xml',
    'xl/embeddings/oleObject1.bin',
    'xl/queryTables/queryTable1.xml',
    'xl/richData/rdrichvalue.xml',
    'xl/model/item.data',
    'xl/chartsheets/sheet1.xml',
    'xl/connections.xml',
    'xl/cellimages.xml',
    'xl/vbaProject.bin',
    'xl/vbaProjectSignature.bin',
  ])('%s is known', (name) => {
    expect(isKnownPart(name)).toBe(true);
  });

  it.each([
    'customXml/item1.xml',
    'xl/dialogsheets/sheet1.xml',
    'xl/macrosheets/sheet1.xml',
    'xl/diagrams/data1.xml',
    'xl/revisions/revisionHeaders.xml',
    'xl/timelines/timeline1.xml',
    'xl/worksheets/sub/sheet.xml',
    'content.xml',
    'xl/volatileDependencies.xml',
  ])('%s is not known', (name) => {
    expect(isKnownPart(name)).toBe(false);
  });
});
