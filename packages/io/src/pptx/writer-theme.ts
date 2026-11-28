import type { PptxTheme, PptxThemeSlot } from '@b32nio/spindle-shared/export/pptx';
import { esc, hex6, NS_A, XML_DECLARATION } from './writer-xml';
import { THEME_SLOTS } from './color';

// `ppt/theme/theme1.xml` from the deck theme (plan K12). The colour scheme and the font scheme are the deck's; the format scheme
// (fill, line and effect styles that `p:style` references point at) is Office's usual one, which the editor does not model.

const SOLID = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';

const FMT_SCHEME = `<a:fmtScheme name="Office"><a:fillStyleLst>${SOLID}<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${SOLID}<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>`;

const color = (slot: PptxThemeSlot, theme: PptxTheme): string => `<a:${slot}><a:srgbClr val="${hex6(theme.colors[slot])}"/></a:${slot}>`;

const font = (latin: string, ea: string | undefined, cs: string | undefined): string =>
  `<a:latin typeface="${esc(latin)}"/><a:ea typeface="${esc(ea ?? '')}"/><a:cs typeface="${esc(cs ?? '')}"/>`;

/** The theme part: 12 colours, the major and minor font families (with the east Asian and complex-script ones when the deck names them). */
export function themeXml(theme: PptxTheme): string {
  const f = theme.fonts;
  return (
    `${XML_DECLARATION}<a:theme xmlns:a="${NS_A}" name="${esc(theme.name)}"><a:themeElements>` +
    `<a:clrScheme name="${esc(theme.name)}">${THEME_SLOTS.map((s) => color(s, theme)).join('')}</a:clrScheme>` +
    `<a:fontScheme name="${esc(theme.name)}"><a:majorFont>${font(f.major, f.majorEastAsian, f.majorComplex)}</a:majorFont><a:minorFont>${font(f.minor, f.minorEastAsian, f.minorComplex)}</a:minorFont></a:fontScheme>` +
    `${FMT_SCHEME}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
  );
}
