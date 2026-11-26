import { LEDGER_RULES, buildLedger, classifyPart } from '../../../../scripts/xlsx-fidelity/ledger';

const base = { lossIds: [] as string[], importFailed: false };

describe('classifyPart', () => {
  it.each([
    ['[Content_Types].xml', 'package.contentTypes'],
    ['_rels/.rels', 'package.relationships'],
    ['xl/_rels/workbook.xml.rels', 'package.relationships'],
    ['xl/worksheets/_rels/sheet1.xml.rels', 'package.relationships'],
    ['xl/workbook.xml', 'workbook'],
    ['xl/worksheets/sheet1.xml', 'worksheet'],
    ['xl/sharedStrings.xml', 'sharedStrings'],
    ['xl/styles.xml', 'styles'],
    ['docProps/core.xml', 'docProps'],
    ['docProps/app.xml', 'docProps'],
    ['xl/calcChain.xml', 'calcChain'],
  ])('%s is mapped by rule %s', (name, rule) => {
    expect(classifyPart(name, base)).toEqual({ name, status: 'mapped', rule });
  });

  it.each([
    ['xl/chartsheets/sheet1.xml', 'chartsheet', 'xlsx.charts'],
    ['xl/drawings/drawing1.xml', 'drawing', 'xlsx.images'],
    ['xl/media/image1.png', 'media', 'xlsx.images'],
    ['xl/charts/chart1.xml', 'chart', 'xlsx.charts'],
    ['xl/charts/style1.xml', 'chart', 'xlsx.charts'],
    ['xl/comments1.xml', 'comments', 'xlsx.comments'],
    ['xl/threadedComments/threadedComment1.xml', 'comments', 'xlsx.comments'],
    ['xl/persons/person.xml', 'comments', 'xlsx.comments'],
    ['xl/comments/comment1.xml', 'comments', 'xlsx.comments'],
    ['xl/threadedComments/threadedComment1.xml', 'comments', 'xlsx.comments.threaded'],
    ['xl/drawings/vmlDrawing1.vml', 'vmlDrawing', 'xlsx.comments'],
    ['xl/drawings/commentsDrawing1.vml', 'vmlDrawing', 'xlsx.comments'],
    ['xl/drawings/vmlDrawing2.vml', 'vmlDrawing', 'xlsx.objects.formControls'],
    ['xl/drawings/vmlDrawing3.vml', 'vmlDrawing', 'xlsx.images'],
    ['xl/drawings/drawing1.xml', 'drawing', 'xlsx.objects.shapes'],
    ['xl/drawings/drawing1.xml', 'drawing', 'xlsx.objects.textBoxes'],
    ['xl/drawings/drawing1.xml', 'drawing', 'xlsx.objects.formControls'],
    ['xl/drawings/drawing1.xml', 'drawing', 'xlsx.objects.embeddedObjects'],
    ['xl/media/image1.png', 'media', 'xlsx.objects.cellImages'],
    ['xl/theme/theme1.xml', 'theme', 'xlsx.styles.themeColors'],
    ['xl/printerSettings/printerSettings1.bin', 'printerSettings', 'xlsx.print.setup'],
    ['xl/metadata.xml', 'metadata', 'xlsx.formulas.dynamicArray'],
    ['xl/ctrlProps/ctrlProp1.xml', 'controls', 'xlsx.objects.formControls'],
    ['xl/activeX/activeX1.xml', 'controls', 'xlsx.objects.formControls'],
    ['xl/embeddings/oleObject1.bin', 'embedding', 'xlsx.objects.embeddedObjects'],
    ['xl/richData/rdrichvalue.xml', 'richData', 'xlsx.workbook.richData'],
    ['xl/richData/rdrichvalue.xml', 'richData', 'xlsx.objects.cellImages'],
    ['xl/cellimages.xml', 'richData', 'xlsx.images'],
    ['xl/connections.xml', 'externalData', 'xlsx.workbook.externalData'],
    ['xl/queryTables/queryTable1.xml', 'externalData', 'xlsx.workbook.externalData'],
    ['xl/model/item.data', 'externalData', 'xlsx.workbook.externalData'],
    ['customXml/item1.xml', 'unknownPart', 'xlsx.package.unaccountedParts'],
    ['something/else.bin', 'unknownPart', 'xlsx.package.unaccountedParts'],
    ['xl/tables/table1.xml', 'table', 'xlsx.tables'],
    ['xl/pivotTables/pivotTable1.xml', 'pivot', 'xlsx.pivotTables'],
    ['xl/pivotCache/pivotCacheDefinition1.xml', 'pivot', 'xlsx.pivotTables'],
    ['xl/slicers/slicer1.xml', 'slicer', 'xlsx.slicers'],
    ['xl/slicerCaches/slicerCache1.xml', 'slicer', 'xlsx.slicers'],
    ['xl/externalLinks/externalLink1.xml', 'externalLink', 'xlsx.externalLinks'],
    ['xl/vbaProject.bin', 'macro', 'xlsx.macros'],
    ['xl/vbaProjectSignature.bin', 'macro', 'xlsx.macros'],
  ])('%s is dropped when the loss report has %s', (name, rule, loss) => {
    expect(classifyPart(name, { ...base, lossIds: [loss] })).toEqual({
      name,
      status: 'dropped',
      rule,
      loss: [loss],
    });
  });

  it('is unaccounted when no reported loss covers a part a rule knows about', () => {
    expect(classifyPart('xl/charts/chart1.xml', base)).toEqual({
      name: 'xl/charts/chart1.xml',
      status: 'unaccounted',
      rule: 'chart',
    });
    expect(
      classifyPart('xl/drawings/drawing1.xml', { ...base, lossIds: ['xlsx.comments'] }).status
    ).toBe('unaccounted');
  });

  it('uses the first reported loss id of a rule that lists several', () => {
    expect(
      classifyPart('xl/drawings/drawing1.xml', { ...base, lossIds: ['xlsx.charts'] })
    ).toMatchObject({ status: 'dropped', loss: ['xlsx.charts'] });
    expect(
      classifyPart('xl/drawings/drawing1.xml', { ...base, lossIds: ['xlsx.charts', 'xlsx.images'] })
    ).toMatchObject({ loss: ['xlsx.images'] });
  });

  it.each([
    'xl/theme/theme1.xml',
    'xl/printerSettings/printerSettings1.bin',
    'xl/metadata.xml',
    'customXml/item1.xml',
    'something/else.bin',
  ])('%s is unaccounted while the loss report does not cover it', (name) => {
    expect(classifyPart(name, { ...base, lossIds: ['xlsx.charts'] })).toMatchObject({
      name,
      status: 'unaccounted',
    });
  });

  it('has no rule for an empty part name, so it is unaccounted', () => {
    expect(classifyPart('', base)).toEqual({ name: '', status: 'unaccounted', rule: 'none' });
  });

  it('does not map a worksheet in a sub-folder or a drawing relationship part by the sheet rules', () => {
    expect(classifyPart('xl/worksheets/sub/sheet.xml', base).status).toBe('unaccounted');
    expect(
      classifyPart('xl/worksheets/sub/sheet.xml', { ...base, lossIds: ['xlsx.package.unaccountedParts'] })
    ).toMatchObject({ status: 'dropped', rule: 'unknownPart' });
    expect(classifyPart('xl/drawings/_rels/drawing1.xml.rels', base)).toMatchObject({
      status: 'mapped',
      rule: 'package.relationships',
    });
  });

  it('counts mapped parts as unaccounted when the import failed, but not dropped ones', () => {
    expect(classifyPart('xl/workbook.xml', { ...base, importFailed: true })).toEqual({
      name: 'xl/workbook.xml',
      status: 'unaccounted',
      rule: 'workbook',
    });
    expect(
      classifyPart('xl/media/a.png', { lossIds: ['xlsx.images'], importFailed: true }).status
    ).toBe('dropped');
  });

  it('lets the opaque hook claim a part first', () => {
    expect(
      classifyPart('xl/charts/chart1.xml', { ...base, isOpaque: (n) => n.startsWith('xl/charts/') })
    ).toEqual({ name: 'xl/charts/chart1.xml', status: 'opaque', rule: 'opaque' });
    expect(classifyPart('xl/workbook.xml', { ...base, isOpaque: () => false }).status).toBe(
      'mapped'
    );
  });

  it('has unique rule ids per outcome kind', () => {
    const ids = LEDGER_RULES.map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe('buildLedger', () => {
  it('sorts parts and totals the statuses', () => {
    const ledger = buildLedger({
      parts: {
        kind: 'zip',
        names: [
          'xl/worksheets/sheet1.xml',
          'docProps/app.xml',
          '[Content_Types].xml',
          'xl/media/i.png',
          'xl/charts/c.xml',
        ],
      },
      lossIds: ['xlsx.images'],
      importFailed: false,
      isOpaque: (n) => n === 'xl/charts/c.xml',
    });
    expect(ledger.package).toBe('zip');
    expect(ledger.importError).toBe(false);
    expect(ledger.parts.map((p) => `${p.name}=${p.status}`)).toEqual([
      '[Content_Types].xml=mapped',
      'docProps/app.xml=mapped',
      'xl/charts/c.xml=opaque',
      'xl/media/i.png=dropped',
      'xl/worksheets/sheet1.xml=mapped',
    ]);
    expect(ledger.totals).toEqual({ mapped: 3, opaque: 1, dropped: 1, unaccounted: 0, total: 5 });
  });

  it('is empty for packages that cannot be read', () => {
    expect(
      buildLedger({ parts: { kind: 'ole', names: [] }, lossIds: [], importFailed: true })
    ).toEqual({
      package: 'ole',
      importError: true,
      totals: { mapped: 0, opaque: 0, dropped: 0, unaccounted: 0, total: 0 },
      parts: [],
    });
  });
});
