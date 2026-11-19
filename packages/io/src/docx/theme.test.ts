import JSZip from 'jszip';
import { importDocx } from './reader';
import { parseTheme, preferRecorded, themeColor, themeFont } from './theme';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const theme = (colors: string, fonts = '<a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont>') =>
  `<a:theme ${A}><a:themeElements><a:clrScheme name="t">${colors}</a:clrScheme><a:fontScheme name="f">${fonts}</a:fontScheme></a:themeElements></a:theme>`;
const OFFICE =
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>';

describe('parseTheme', () => {
  it('reads the colour scheme (srgb and system colours) and the major/minor Latin fonts', () => {
    const t = parseTheme(theme(OFFICE))!;
    expect(t.colors.get('dk1')).toBe('#000000');
    expect(t.colors.get('lt1')).toBe('#ffffff');
    expect(t.colors.get('accent1')).toBe('#4f81bd');
    expect(t.majorFont).toBe('Cambria');
    expect(t.minorFont).toBe('Calibri');
  });

  it('copes with whitespace between elements, empty typefaces and missing parts', () => {
    const t = parseTheme(theme('\n  <a:accent1>\n    <a:srgbClr val="112233"/>\n  </a:accent1>\n', '<a:majorFont><a:latin typeface=""/></a:majorFont>'))!;
    expect(t.colors.get('accent1')).toBe('#112233');
    expect(t.majorFont).toBeUndefined();
    expect(t.minorFont).toBeUndefined();
    expect(parseTheme(undefined)).toBeUndefined();
    expect(parseTheme('<<<')).toBeUndefined();
    expect(parseTheme('<nope/>')).toBeUndefined();
  });

  it('ignores colours that are not six hex digits', () => {
    expect(parseTheme(theme('<a:accent1><a:srgbClr val="zzz"/></a:accent1>'))!.colors.has('accent1')).toBe(false);
  });
});

describe('themeFont / themeColor', () => {
  const t = parseTheme(theme(OFFICE));
  it('maps major/minor font references', () => {
    expect(themeFont(t, 'majorHAnsi')).toBe('Cambria');
    expect(themeFont(t, 'minorAscii')).toBe('Calibri');
    expect(themeFont(t, 'minorBidi')).toBe('Calibri');
    expect(themeFont(t, 'other')).toBeUndefined();
    expect(themeFont(undefined, 'minorHAnsi')).toBeUndefined();
    expect(themeFont(t, undefined)).toBeUndefined();
  });

  it('maps every theme colour name to its slot', () => {
    expect(themeColor(t, 'text1')).toBe('#000000');
    expect(themeColor(t, 'dark1')).toBe('#000000');
    expect(themeColor(t, 'background1')).toBe('#ffffff');
    expect(themeColor(t, 'light1')).toBe('#ffffff');
    expect(themeColor(t, 'text2')).toBe('#1f497d');
    expect(themeColor(t, 'background2')).toBe('#eeece1');
    expect(themeColor(t, 'accent1')).toBe('#4f81bd');
    expect(themeColor(t, 'hyperlink')).toBe('#0000ff');
    expect(themeColor(t, 'followedHyperlink')).toBe('#800080');
    expect(themeColor(t, 'accent5')).toBeUndefined(); // not in this theme
    expect(themeColor(t, 'none')).toBeUndefined();
    expect(themeColor(undefined, 'accent1')).toBeUndefined();
  });

  it('applies shade and tint to the luminance, matching the values Word records', () => {
    // From Word's own default styles: accent1 4F81BD shade BF is 365F91, shade 7F is 243F60; tint 33 is about DBE5F1.
    const close = (got: string | undefined, want: string) => {
      const g = [1, 3, 5].map((i) => parseInt(got!.slice(i, i + 2), 16));
      const w = [0, 2, 4].map((i) => parseInt(want.slice(i, i + 2), 16));
      expect(g.every((c, i) => Math.abs(c - w[i]) <= 1)).toBe(true);
    };
    close(themeColor(t, 'accent1', undefined, 'BF'), '365F91');
    close(themeColor(t, 'accent1', undefined, '7F'), '243F60');
    close(themeColor(t, 'accent1', '33'), 'DBE5F1');
    close(themeColor(t, 'text1', 'BF'), '404040');
    expect(themeColor(t, 'accent1', 'FF')).toBe('#4f81bd');
    expect(themeColor(t, 'accent1', undefined, '00')).toBe('#000000');
    expect(themeColor(t, 'accent1', 'zz')).toBe('#4f81bd'); // not a number: ignored
  });
});

describe('preferRecorded', () => {
  it("takes Word's recorded colour when it is within a few steps, the computed one otherwise", () => {
    expect(preferRecorded('#366091', '365F91')).toBe('#365f91');
    expect(preferRecorded('#8f2a20', 'C0392B')).toBe('#8f2a20');
    expect(preferRecorded('#8f2a20', undefined)).toBe('#8f2a20');
    expect(preferRecorded('#8f2a20', 'auto')).toBe('#8f2a20');
  });
});

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
async function importWithTheme(styles: string, body: string, themeXml?: string) {
  const zip = new JSZip();
  zip.file('word/document.xml', `<?xml version="1.0"?><w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`);
  zip.file('word/styles.xml', `<?xml version="1.0"?><w:styles ${NS}>${styles}</w:styles>`);
  if (themeXml) zip.file('word/theme/theme1.xml', themeXml);
  const result = await importDocx(await zip.generateAsync({ type: 'uint8array' }));
  const block = result.sections[0].blocks[0];
  return 'content' in block ? (block.content[0] as unknown as Record<string, unknown>) : {};
}
const para = (rPr: string) => `<w:p><w:r><w:rPr>${rPr}</w:rPr><w:t>x</w:t></w:r></w:p>`;
const pick = (r: Record<string, unknown>) => ({ fontFamily: r.fontFamily, color: r.color, backgroundColor: r.backgroundColor });

describe('theme resolution in runs', () => {
  it('resolves theme fonts from docDefaults, styles and direct properties; theme beats the literal font', async () => {
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:styleId="Big"><w:name w:val="Big"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr></w:style>';
    const t = theme(OFFICE);
    expect(pick(await importWithTheme(styles, '<w:p><w:r><w:t>x</w:t></w:r></w:p>', t)).fontFamily).toBe('Calibri');
    expect(pick(await importWithTheme(styles, '<w:p><w:pPr><w:pStyle w:val="Big"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>', t)).fontFamily).toBe('Cambria');
    expect(pick(await importWithTheme(styles, para('<w:rFonts w:ascii="Georgia" w:asciiTheme="majorHAnsi"/>'), t)).fontFamily).toBe('Cambria');
    expect(pick(await importWithTheme(styles, para('<w:rFonts w:ascii="Georgia"/>'), t)).fontFamily).toBe('Georgia');
  });

  it('falls back to the literal font and colour when there is no theme or it lacks the slot', async () => {
    expect(pick(await importWithTheme('', para('<w:rFonts w:ascii="Georgia" w:asciiTheme="minorHAnsi"/><w:color w:val="112233" w:themeColor="accent1"/>'))).fontFamily).toBe('Georgia');
    expect(pick(await importWithTheme('', para('<w:color w:val="112233" w:themeColor="accent1"/>'))).color).toBe('#112233');
    expect(pick(await importWithTheme('', para('<w:color w:val="112233" w:themeColor="accent5"/>'), theme(OFFICE))).color).toBe('#112233');
    expect(pick(await importWithTheme('', para('<w:rFonts w:asciiTheme="minorHAnsi"/>'), theme(OFFICE, ''))).fontFamily).toBeUndefined();
  });

  it('resolves theme colours with tint and shade, preferring the value Word recorded when it is close', async () => {
    const t = theme(OFFICE);
    // Word-recorded 365F91 is within a step of the computed shade.
    expect(pick(await importWithTheme('', para('<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>'), t)).color).toBe('#365f91');
    // A producer that left the base colour next to a shade: the theme computation wins.
    const darker = pick(await importWithTheme('', para('<w:color w:val="4F81BD" w:themeColor="accent1" w:themeShade="7F"/>'), t)).color as string;
    expect(darker).toMatch(/^#2[34]3[ef]6[01]$/);
    expect(pick(await importWithTheme('', para('<w:color w:val="auto" w:themeColor="text1"/>'), t)).color).toBe('#000000');
  });

  it('resolves theme colours in styles and in shading fills', async () => {
    const styles = '<w:style w:type="character" w:styleId="Acc"><w:name w:val="Acc"/><w:rPr><w:color w:themeColor="accent2"/><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFill="background2"/></w:rPr></w:style>';
    const r = pick(await importWithTheme(styles, para('<w:rStyle w:val="Acc"/>'), theme(OFFICE)));
    expect(r.color).toBe('#c0504d');
    expect(r.backgroundColor).toBe('#eeece1');
  });
});
