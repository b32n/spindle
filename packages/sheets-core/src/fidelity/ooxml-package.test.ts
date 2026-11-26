import JSZip from 'jszip';
import {
  DEFAULT_MAX_ENTRY_BYTES,
  isWellFormed,
  readParts,
  relationshipTargets,
  zipEntrySizes,
} from '../../../../scripts/xlsx-fidelity/ooxml-package';

const zipOf = async (files: Record<string, string>, comment?: string) => {
  const z = new JSZip();
  for (const [n, d] of Object.entries(files)) z.file(n, d);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE', comment });
};
const u8 = (...b: number[]) => new Uint8Array(b);

describe('zipEntrySizes', () => {
  it('reads compressed and uncompressed sizes from the central directory', async () => {
    const bytes = await zipOf({ 'a.xml': 'x'.repeat(5000), 'dir/b.txt': 'hello' });
    const sizes = zipEntrySizes(bytes);
    expect(Object.keys(sizes).sort()).toEqual(['a.xml', 'dir/', 'dir/b.txt']);
    expect(sizes['a.xml'].uncompressed).toBe(5000);
    expect(sizes['a.xml'].compressed).toBeLessThan(100);
    expect(sizes['dir/b.txt'].uncompressed).toBe(5);
  });

  it('copes with an archive comment after the end record', async () => {
    const sizes = zipEntrySizes(await zipOf({ 'a.xml': 'abc' }, 'a comment'));
    expect(sizes['a.xml'].uncompressed).toBe(3);
  });

  it('returns nothing when there is no end-of-central-directory record', async () => {
    expect(zipEntrySizes(u8())).toEqual({});
    expect(zipEntrySizes(u8(1, 2, 3))).toEqual({});
    const bytes = await zipOf({ 'a.xml': 'abc'.repeat(100) });
    expect(zipEntrySizes(bytes.subarray(0, Math.floor(bytes.length / 2)))).toEqual({});
  });

  it('stops at a damaged central directory entry', async () => {
    const bytes = await zipOf({ 'a.xml': 'abc', 'b.xml': 'def' });
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const eocd = bytes.length - 22;
    const cd = view.getUint32(eocd + 16, true);
    const broken = bytes.slice();
    new DataView(broken.buffer).setUint32(cd, 0, true);
    expect(zipEntrySizes(broken)).toEqual({});
    const lying = bytes.slice();
    new DataView(lying.buffer).setUint32(eocd + 16, lying.length - 10, true);
    expect(zipEntrySizes(lying)).toEqual({});
    const many = bytes.slice();
    new DataView(many.buffer).setUint16(eocd + 10, 9, true);
    expect(Object.keys(zipEntrySizes(many)).sort()).toEqual(['a.xml', 'b.xml']);
  });
});

describe('readParts', () => {
  it('reads the names and the text of xml and rels parts', async () => {
    const p = await readParts(
      await zipOf({
        'xl/workbook.xml': '<a/>',
        'xl/_rels/workbook.xml.rels': '<r/>',
        'xl/media/i.png': 'png',
        'xl/dir/': '',
      })
    );
    expect(p.kind).toBe('zip');
    expect(p.names.sort()).toEqual([
      'xl/_rels/workbook.xml.rels',
      'xl/media/i.png',
      'xl/workbook.xml',
    ]);
    expect(Object.keys(p.xml).sort()).toEqual(['xl/_rels/workbook.xml.rels', 'xl/workbook.xml']);
    expect(p.oversized).toEqual([]);
    expect(p.sizes['xl/workbook.xml'].uncompressed).toBe(4);
  });

  it('leaves out parts above the size cap and lists them', async () => {
    const p = await readParts(
      await zipOf({ 'big.xml': '<a>' + ' '.repeat(100) + '</a>', 'small.xml': '<a/>' }),
      50
    );
    expect(p.oversized).toEqual(['big.xml']);
    expect(Object.keys(p.xml)).toEqual(['small.xml']);
    expect(DEFAULT_MAX_ENTRY_BYTES).toBe(64 * 1024 * 1024);
  });

  it('classifies OLE compound files, truncated zips and other bytes without throwing', async () => {
    expect((await readParts(u8(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0))).kind).toBe(
      'ole'
    );
    expect((await readParts(u8(0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe2))).kind).toBe(
      'unreadable'
    );
    const bytes = await zipOf({ 'a.xml': 'abc'.repeat(200) });
    const truncated = await readParts(bytes.subarray(0, 60));
    expect(truncated.kind).toBe('unreadable');
    expect(truncated.names).toEqual([]);
    expect((await readParts(u8())).kind).toBe('unreadable');
  });

  it('enforces the cap on the actual inflated size, not on the declared one', async () => {
    const bytes = await zipOf({
      'big.xml': '<a>' + 'x'.repeat(5000) + '</a>',
      'small.xml': '<a/>',
    });
    const lying = bytes.slice();
    const view = new DataView(lying.buffer);
    const eocd = lying.length - 22;
    let at = view.getUint32(eocd + 16, true);
    for (let n = view.getUint16(eocd + 10, true); n > 0; n--) {
      const nameLen = view.getUint16(at + 28, true);
      if (new TextDecoder().decode(lying.subarray(at + 46, at + 46 + nameLen)) === 'big.xml')
        view.setUint32(at + 24, 10, true); // declares 10 bytes
      at += 46 + nameLen + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    }
    const p = await readParts(lying, 100);
    expect(p.sizes['big.xml'].uncompressed).toBe(10);
    expect(p.oversized).toEqual(['big.xml']);
    expect(Object.keys(p.xml)).toEqual(['small.xml']);
  });

  it('lists an entry that cannot be inflated instead of throwing', async () => {
    const bytes = await zipOf({ 'bad.xml': '<a>' + 'x'.repeat(5000) + '</a>' });
    const broken = bytes.slice();
    const start = 30 + 'bad.xml'.length + 4; // inside the deflate stream
    for (let i = start; i < start + 8; i++) broken[i] = 0xff;
    const p = await readParts(broken);
    expect(p.kind).toBe('zip');
    expect(p.oversized).toEqual(['bad.xml']);
  });

  it('treats an entry missing from the size table as readable', async () => {
    const bytes = await zipOf({ 'a.xml': '<a/>', 'b.xml': '<b/>' });
    const lying = bytes.slice();
    new DataView(lying.buffer).setUint16(lying.length - 22 + 10, 1, true); // the end record claims one entry
    const p = await readParts(lying);
    expect(Object.keys(p.sizes)).toHaveLength(1);
    expect(Object.keys(p.xml).sort()).toEqual(['a.xml', 'b.xml']);
  });
});

describe('isWellFormed', () => {
  it.each([
    ['<a><b x="1"/><c>text</c></a>', true],
    ['<?xml version="1.0"?><!-- c --><a><![CDATA[<not> a tag]]></a>', true],
    ['<!DOCTYPE a [<!ENTITY e "v"><!ENTITY f "&e;&e;">]><a>&f;</a>', true],
    ['<!DOCTYPE a><a/>', true],
    ["<a x='>'/>", true],
    ['<a><b></a>', false],
    ['<a></a></a>', false],
    ['<a>', false],
    ['<a>text < more</a>', false],
    ['<a/><', false],
    ['', true],
  ])('%j -> %s', (xml, ok) => {
    expect(isWellFormed(xml)).toBe(ok);
  });
});

describe('relationshipTargets', () => {
  const rels = (...items: string[]) => `<Relationships>${items.join('')}</Relationships>`;
  it('resolves relative, absolute and parent targets against the owning part', () => {
    const xml = rels(
      '<Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/>',
      '<Relationship Id="rId2" Type="t" Target="/xl/styles.xml"/>',
      '<Relationship Id="rId3" Type="t" Target="../media/image%201.png"/>'
    );
    expect(relationshipTargets('xl/_rels/workbook.xml.rels', xml)).toEqual([
      { id: 'rId1', target: 'xl/worksheets/sheet1.xml' },
      { id: 'rId2', target: 'xl/styles.xml' },
      { id: 'rId3', target: 'media/image 1.png' },
    ]);
    expect(
      relationshipTargets(
        'xl/worksheets/_rels/sheet1.xml.rels',
        rels('<Relationship Id="a" Target="../media/i.png"/>')
      )[0].target
    ).toBe('xl/media/i.png');
    expect(
      relationshipTargets('_rels/.rels', rels('<Relationship Id="a" Target="xl/workbook.xml"/>'))[0]
        .target
    ).toBe('xl/workbook.xml');
  });

  it('treats a target that climbs out of the package root as pointing at a missing part', () => {
    const r = relationshipTargets(
      'xl/_rels/workbook.xml.rels',
      rels(
        '<Relationship Id="a" Target="../../x.xml"/>',
        '<Relationship Id="b" Target="/../y.xml"/>',
        '<Relationship Id="c" Target="./sheets/../z.xml"/>'
      )
    );
    expect(r.map((x) => x.target)).toEqual(['../../../x.xml', '..//../y.xml', 'xl/z.xml']);
  });

  it('gives null for external targets and for relationships without a target', () => {
    const xml = rels(
      '<Relationship Id="x" Target="https://example.com" TargetMode="External"/>',
      '<Relationship Id="y"/>',
      '<Relationship Target="a.xml"/>'
    );
    expect(relationshipTargets('xl/_rels/workbook.xml.rels', xml)).toEqual([
      { id: 'x', target: null },
      { id: 'y', target: null },
      { id: '', target: 'xl/a.xml' },
    ]);
  });

  it('keeps a target with broken percent-encoding as written', () => {
    expect(
      relationshipTargets('_rels/.rels', rels('<Relationship Id="a" Target="100%.xml"/>'))[0].target
    ).toBe('100%.xml');
  });
});
