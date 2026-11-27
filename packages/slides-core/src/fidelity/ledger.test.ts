import { LEDGER_RULES, buildLedger, classifyPart } from '../../../../scripts/pptx-fidelity/ledger';

const base = { lossIds: [] as string[], importFailed: false };

describe('classifyPart', () => {
  it.each([
    ['[Content_Types].xml', 'package.contentTypes'],
    ['_rels/.rels', 'package.relationships'],
    ['ppt/_rels/presentation.xml.rels', 'package.relationships'],
    ['ppt/slides/_rels/slide1.xml.rels', 'package.relationships'],
    ['docProps/core.xml', 'docProps'],
    ['docProps/thumbnail.jpeg', 'docProps'],
    ['ppt/presentation.xml', 'presentation'],
    ['ppt/presProps.xml', 'presProps'],
    ['ppt/viewProps.xml', 'presProps'],
    ['ppt/tableStyles.xml', 'tableStyles'],
    ['ppt/printerSettings/printerSettings1.bin', 'printerSettings'],
    ['ppt/slides/slide1.xml', 'slide'],
    ['ppt/slideLayouts/slideLayout1.xml', 'layout'],
    ['ppt/slideMasters/slideMaster1.xml', 'master'],
    ['ppt/theme/theme1.xml', 'theme'],
    ['ppt/notesSlides/notesSlide1.xml', 'notesSlide'],
    ['ppt/notesMasters/notesMaster1.xml', 'notesMaster'],
    ['ppt/comments/comment1.xml', 'comments'],
    ['ppt/comments/modernComment_1.xml', 'comments'],
    ['ppt/commentAuthors.xml', 'comments'],
    ['ppt/authors.xml', 'comments'],
    ['ppt/media/image1.png', 'picture'],
    ['ppt/media/image2.jpeg', 'picture'],
    ['ppt/media/image3.svg', 'picture'],
  ])('%s is mapped by rule %s once an importer exists', (name, rule) => {
    expect(classifyPart(name, base)).toEqual({ name, status: 'mapped', rule });
  });

  it.each([
    ['ppt/handoutMasters/handoutMaster1.xml', 'handoutMaster', 'pptx.masters.handoutMaster'],
    ['ppt/charts/chart1.xml', 'chart', 'pptx.charts'],
    ['ppt/charts/style1.xml', 'chart', 'pptx.charts'],
    ['ppt/diagrams/data1.xml', 'diagram', 'pptx.smartArt'],
    ['ppt/diagrams/drawing1.xml', 'diagram', 'pptx.smartArt'],
    ['ppt/media/clip.mp4', 'audioVideo', 'pptx.media'],
    ['ppt/media/sound.WAV', 'audioVideo', 'pptx.media'],
    ['ppt/media/clip.mov', 'audioVideo', 'pptx.media'],
    ['ppt/media/image1.emf', 'vectorPicture', 'pptx.images.unsupportedFormat'],
    ['ppt/media/image2.wmf', 'vectorPicture', 'pptx.images.unsupportedFormat'],
    ['ppt/media/image3.EMZ', 'vectorPicture', 'pptx.images.unsupportedFormat'],
    ['ppt/embeddings/oleObject1.bin', 'embedding', 'pptx.embeddedObjects'],
    ['ppt/embeddings/Microsoft_Excel_Sheet1.xlsx', 'embedding', 'pptx.charts'],
    ['ppt/activeX/activeX1.xml', 'controls', 'pptx.embeddedObjects'],
    ['ppt/ink/ink1.xml', 'ink', 'pptx.ink'],
    ['ppt/vbaProject.bin', 'macro', 'pptx.macros'],
    ['ppt/vbaProjectSignature.bin', 'macro', 'pptx.macros'],
    ['_xmlsignatures/sig1.xml', 'signature', 'pptx.signatures'],
    ['customXml/item1.xml', 'unknownPart', 'pptx.package.unaccountedParts'],
    ['ppt/tags/tag1.xml', 'unknownPart', 'pptx.package.unaccountedParts'],
    ['ppt/fonts/font1.fntdata', 'unknownPart', 'pptx.package.unaccountedParts'],
  ])('%s is dropped by rule %s when the loss report has %s', (name, rule, loss) => {
    expect(classifyPart(name, { ...base, lossIds: [loss] })).toEqual({
      name,
      status: 'dropped',
      rule,
      loss: [loss],
    });
  });

  it('is unaccounted when no reported loss covers a part a rule knows about', () => {
    expect(classifyPart('ppt/charts/chart1.xml', base)).toEqual({
      name: 'ppt/charts/chart1.xml',
      status: 'unaccounted',
      rule: 'chart',
    });
    expect(classifyPart('ppt/embeddings/a.bin', { ...base, lossIds: ['pptx.ink'] }).status).toBe(
      'unaccounted'
    );
  });

  it('uses the first reported loss id of a rule that lists several', () => {
    expect(classifyPart('ppt/embeddings/a.bin', { ...base, lossIds: ['pptx.charts'] })).toMatchObject({
      status: 'dropped',
      loss: ['pptx.charts'],
    });
    expect(
      classifyPart('ppt/embeddings/a.bin', { ...base, lossIds: ['pptx.charts', 'pptx.embeddedObjects'] })
    ).toMatchObject({ loss: ['pptx.embeddedObjects'] });
  });

  it('maps relationship parts only inside a _rels folder', () => {
    expect(classifyPart('stray/file.rels', base).status).toBe('unaccounted');
    expect(classifyPart('stray/file.rels', { ...base, lossIds: ['pptx.package.unaccountedParts'] }).rule).toBe('unknownPart');
  });

  it('has no rule for an empty part name, so it is unaccounted', () => {
    expect(classifyPart('', base)).toEqual({ name: '', status: 'unaccounted', rule: 'none' });
  });

  it('does not map a slide in a sub-folder by the slide rule', () => {
    expect(classifyPart('ppt/slides/sub/slide.xml', base).status).toBe('unaccounted');
    expect(
      classifyPart('ppt/slides/sub/slide.xml', { ...base, lossIds: ['pptx.package.unaccountedParts'] })
    ).toMatchObject({ status: 'dropped', rule: 'unknownPart' });
  });

  it('counts mapped parts as unaccounted when the import failed (always, today), but not dropped ones', () => {
    expect(classifyPart('ppt/slides/slide1.xml', { ...base, importFailed: true })).toEqual({
      name: 'ppt/slides/slide1.xml',
      status: 'unaccounted',
      rule: 'slide',
    });
    expect(
      classifyPart('ppt/media/a.mp3', { lossIds: ['pptx.media'], importFailed: true }).status
    ).toBe('dropped');
  });

  it('lets the opaque hook claim a part first', () => {
    expect(
      classifyPart('ppt/charts/chart1.xml', { ...base, isOpaque: (n) => n.startsWith('ppt/charts/') })
    ).toEqual({ name: 'ppt/charts/chart1.xml', status: 'opaque', rule: 'opaque' });
    expect(classifyPart('ppt/slides/slide1.xml', { ...base, isOpaque: () => false }).status).toBe(
      'mapped'
    );
  });

  it('covers a kept-content part the store refused by the too-large item, and by nothing else it would not mention', () => {
    for (const part of ['ppt/charts/chart1.xml', 'ppt/diagrams/data1.xml', 'ppt/media/a.mp4', 'ppt/embeddings/o.bin', 'ppt/activeX/a.bin', 'ppt/ink/ink1.xml']) {
      expect(classifyPart(part, { lossIds: ['pptx.opaque.tooLarge'], importFailed: false })).toMatchObject({ status: 'dropped', loss: ['pptx.opaque.tooLarge'] });
      expect(classifyPart(part, { lossIds: [], importFailed: false }).status).toBe('unaccounted');
    }
  });

  it('has unique rule ids, and only pptx loss ids', () => {
    const ids = LEDGER_RULES.map((r) => r.id);
    expect(new Set(ids).size).toBe(ids.length);
    for (const r of LEDGER_RULES)
      if (r.outcome !== 'mapped') for (const id of r.outcome.dropped) expect(id).toMatch(/^pptx\./);
  });
});

describe('buildLedger', () => {
  it('sorts parts and totals the statuses', () => {
    const ledger = buildLedger({
      parts: {
        kind: 'zip',
        names: [
          'ppt/slides/slide1.xml',
          'docProps/app.xml',
          '[Content_Types].xml',
          'ppt/media/clip.mp4',
          'ppt/charts/c.xml',
        ],
      },
      lossIds: ['pptx.media'],
      importFailed: false,
      isOpaque: (n) => n === 'ppt/charts/c.xml',
    });
    expect(ledger.package).toBe('zip');
    expect(ledger.importError).toBe(false);
    expect(ledger.parts.map((p) => `${p.name}=${p.status}`)).toEqual([
      '[Content_Types].xml=mapped',
      'docProps/app.xml=mapped',
      'ppt/charts/c.xml=opaque',
      'ppt/media/clip.mp4=dropped',
      'ppt/slides/slide1.xml=mapped',
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
