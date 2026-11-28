import { attrs, colorXml, emu, esc, hex6, pct, relsXml, solidFillXml, xfrmXml } from './writer-xml';
import { bodyPrXml, levelXml, lstStyleXml, plainParagraphsXml, typefaceOf } from './writer-text';
import { themeXml } from './writer-theme';
import { FALLBACK_THEME } from './theme';

describe('writer-xml', () => {
  it('escapes the markup characters and drops what XML cannot carry', () => {
    expect(esc('a & b < c > "d"')).toBe('a &amp; b &lt; c &gt; &quot;d&quot;');
    expect(esc('a\u0000b\u000Bc￾d\tz\n')).toBe('abcd\tz\n');
  });

  it('writes attributes, leaving out undefined values and escaping strings', () => {
    expect(attrs({ a: 1, b: 'x"y', c: undefined, d: 0 })).toBe(' a="1" b="x&quot;y" d="0"');
    expect(attrs({})).toBe('');
  });

  it('converts units with rounding, and a hex colour to the upper case form (black for anything else)', () => {
    expect(emu(1)).toBe(9525);
    expect(emu(0.333)).toBe(3172);
    expect(pct(0.625)).toBe(62500);
    expect(hex6('#ab12Cd')).toBe('AB12CD');
    expect(hex6('ab12cd')).toBe('AB12CD');
    expect(hex6('red')).toBe('000000');
  });

  it('writes colours: a theme slot with its transforms, a literal with alpha', () => {
    expect(colorXml({ kind: 'theme', slot: 'accent1' })).toBe('<a:schemeClr val="accent1"/>');
    expect(colorXml({ kind: 'theme', slot: 'dk2', lumMod: 0.75, lumOff: 0.25, alpha: 0.5 })).toBe('<a:schemeClr val="dk2"><a:lumMod val="75000"/><a:lumOff val="25000"/><a:alpha val="50000"/></a:schemeClr>');
    expect(colorXml({ kind: 'rgb', hex: '#00ff00', alpha: 0.2 })).toBe('<a:srgbClr val="00FF00"><a:alpha val="20000"/></a:srgbClr>');
    expect(solidFillXml({ kind: 'rgb', hex: '#000000' })).toBe('<a:solidFill><a:srgbClr val="000000"/></a:solidFill>');
  });

  it('writes a frame with rotation and flips only when set, and relationships with external targets and explicit type URIs', () => {
    expect(xfrmXml({ x: 1, y: 2, w: 3, h: 4 })).toBe('<a:xfrm><a:off x="9525" y="19050"/><a:ext cx="28575" cy="38100"/></a:xfrm>');
    expect(xfrmXml({ x: 0, y: 0, w: 1, h: 1, rot: 90, flipH: true, flipV: true }, 'p:xfrm')).toContain('<p:xfrm rot="5400000" flipH="1" flipV="1">');
    const rels = relsXml([{ id: 'rId1', type: 'hyperlink', target: 'https://example.com/?a=1&b=2', external: true }, { id: 'rId2', type: 'x', typeUri: 'urn:t', target: 'a.xml' }]);
    expect(rels).toContain('Target="https://example.com/?a=1&amp;b=2" TargetMode="External"');
    expect(rels).toContain('Type="urn:t" Target="a.xml"');
  });
});

describe('writer-text', () => {
  it('maps the theme fonts to their references and leaves other families alone', () => {
    expect(typefaceOf('major')).toBe('+mj-lt');
    expect(typefaceOf('minor')).toBe('+mn-lt');
    expect(typefaceOf('Arial')).toBe('Arial');
  });

  it('writes an empty level as nothing and a level that states only attributes as an empty element', () => {
    expect(levelXml(0, {})).toBe('');
    expect(levelXml(2, { align: 'left' })).toBe('<a:lvl3pPr algn="l"/>');
    expect(lstStyleXml([{}, {}])).toBe('<a:lstStyle/>');
    expect(lstStyleXml([{}, { bold: false, italic: true }])).toBe('<a:lstStyle><a:lvl2pPr><a:defRPr b="0" i="1"></a:defRPr></a:lvl2pPr></a:lstStyle>');
  });

  it('writes spacing, bullets and tabs in schema order', () => {
    const xml = levelXml(0, { lineSpacingPts: 12, spaceBefore: 6, spaceAfterPct: 0.5, bullet: { kind: 'auto', scheme: 'romanLcPeriod' }, bulletSizePct: 1, tabStops: [{ pos: 10, align: 'ctr' }] });
    expect(xml).toBe('<a:lvl1pPr><a:lnSpc><a:spcPts val="1200"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef><a:spcAft><a:spcPct val="50000"/></a:spcAft><a:buSzPct val="100000"/><a:buAutoNum type="romanLcPeriod"/><a:tabLst><a:tab pos="95250" algn="ctr"/></a:tabLst></a:lvl1pPr>');
    expect(levelXml(0, { tabStops: [] })).toBe('');
  });

  it('writes body properties: nothing, plain, or with the three autofit modes', () => {
    expect(bodyPrXml(undefined)).toBe('<a:bodyPr/>');
    expect(bodyPrXml({}, 'bottom')).toBe('<a:bodyPr anchor="b"/>');
    expect(bodyPrXml({ anchor: 'top', anchorCtr: false }, 'bottom')).toBe('<a:bodyPr anchor="t" anchorCtr="0"/>');
    expect(bodyPrXml({ autofit: { kind: 'none' } })).toBe('<a:bodyPr><a:noAutofit/></a:bodyPr>');
    expect(bodyPrXml({ autofit: { kind: 'resize' } })).toBe('<a:bodyPr><a:spAutoFit/></a:bodyPr>');
    expect(bodyPrXml({ autofit: { kind: 'shrink' } })).toBe('<a:bodyPr><a:normAutofit/></a:bodyPr>');
  });

  it('writes plain paragraphs for a prompt, one per line, with an empty line as an empty paragraph', () => {
    expect(plainParagraphsXml('')).toBe('<a:p><a:endParaRPr lang="en-US"/></a:p>');
    expect(plainParagraphsXml('a<\n\nb')).toBe('<a:p><a:r><a:rPr lang="en-US"/><a:t>a&lt;</a:t></a:r></a:p><a:p><a:endParaRPr lang="en-US"/></a:p><a:p><a:r><a:rPr lang="en-US"/><a:t>b</a:t></a:r></a:p>');
  });
});

describe('writer-theme', () => {
  it('writes the twelve colours and the font scheme, with empty east Asian and complex faces when the theme names none', () => {
    const xml = themeXml(FALLBACK_THEME);
    expect(xml.match(/<a:srgbClr val="[0-9A-F]{6}"\/><\/a:(?:dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>/g)).toHaveLength(12);
    expect(xml).toContain('<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>');
    expect(xml).toContain('<a:fmtScheme name="Office">');
  });
});
