import { isWellFormed } from '../fidelity-core/ooxml-package';
import type { OoxmlParts } from '../fidelity-core/ooxml-package';
import type { Construct } from '../fidelity-core/types';
import {
  maxGroupDepth,
  missingSlideParts,
  relationshipCycles,
  relationshipEdges,
  slideSize,
} from './structure';

/** PPTX construct registry: each probe counts one construct in the package XML (0 = absent). */
export type PptxConstruct = Construct<OoxmlParts>;

const count = (s: string, re: RegExp) => s.match(re)?.length ?? 0;
const SLIDE = /^ppt\/slides\/slide\d+\.xml$/;
const LAYOUT = /^ppt\/slideLayouts\/slideLayout\d+\.xml$/;
const MASTER = /^ppt\/slideMasters\/slideMaster\d+\.xml$/;
const THEME = /^ppt\/theme\/theme\d+\.xml$/;
const partsOf = (p: OoxmlParts, names: RegExp) =>
  Object.entries(p.xml)
    .filter(([n]) => names.test(n))
    .map(([, x]) => x);
const inParts = (p: OoxmlParts, names: RegExp, re: RegExp) =>
  partsOf(p, names).reduce((sum, x) => sum + count(x, re), 0);
const inSlides = (p: OoxmlParts, re: RegExp) => inParts(p, SLIDE, re);
const countNames = (p: OoxmlParts, re: RegExp) => p.names.filter((n) => re.test(n)).length;

/** Opening tag of `tag` carrying every attribute pattern, in any order (each pattern is a regex source such as `lvl="[1-8]"`). */
const el = (tag: string, ...attrs: string[]) =>
  new RegExp(`<${tag}\\b${attrs.map((a) => `(?=[^>]*\\b(?:${a}))`).join('')}[^>]*>`, 'g');
/** The rest of an opening tag that is not self-closing: attributes, then `>` (a `/>` never matches). */
const OPEN = '(?:[^>/]|/(?!>))*>';
const tagRe = (...names: string[]) => new RegExp(`<a:(?:${names.join('|')})\\b`, 'g');

const PROPS = new RegExp(`<p:spPr\\b${OPEN}([\\s\\S]*?)</p:spPr>`, 'g');
const OUTLINE = new RegExp(`<a:ln\\b${OPEN}[\\s\\S]*?</a:ln>`, 'g');
const RUN_PROPS = `<a:rPr\\b${OPEN}(?:(?!</a:rPr>)[\\s\\S])*?`;

/** Shape property blocks of the slides, with the outline removed (fills) and the outlines apart. */
function shapeProps(p: OoxmlParts): { fills: string[]; lines: string[] } {
  const fills: string[] = [];
  const lines: string[] = [];
  for (const x of partsOf(p, SLIDE)) {
    for (const m of x.matchAll(PROPS)) {
      lines.push(...(m[1].match(OUTLINE) ?? []));
      fills.push(m[1].replace(OUTLINE, ''));
    }
  }
  return { fills, lines };
}
const inFills = (p: OoxmlParts, re: RegExp) =>
  shapeProps(p).fills.reduce((sum, x) => sum + count(x, re), 0);
const inLines = (p: OoxmlParts, re: RegExp) =>
  shapeProps(p).lines.reduce((sum, x) => sum + count(x, re), 0);

const presetNames = (p: OoxmlParts) =>
  partsOf(p, SLIDE).flatMap((x) =>
    [...x.matchAll(el('a:prstGeom', 'prst="([^"]*)"'))].map((m) => /prst="([^"]*)"/.exec(m[0])?.[1])
  );

const sizeIs = (kind: string) => (p: OoxmlParts) => (slideSize(p)?.kind === kind ? 1 : 0);
const placeholderTypes = (p: OoxmlParts) =>
  new Set(
    [...partsOf(p, LAYOUT), ...partsOf(p, MASTER)].flatMap((x) =>
      [...x.matchAll(/<p:ph\b[^>]*>/g)].map((m) => /\btype="([^"]*)"/.exec(m[0])?.[1] ?? 'obj')
    )
  );
const masterThemes = (p: OoxmlParts) =>
  new Set(
    relationshipEdges(p)
      .filter((e) => MASTER.test(e.from) && e.type.endsWith('/theme'))
      .map((e) => e.to)
  );

/** True when the OLE container holds a directory entry with this name (stored as UTF-16LE). */
function ole(p: OoxmlParts, name: string): boolean {
  return Buffer.from(p.raw).includes(Buffer.from(name, 'utf16le'));
}

function brokenRels(p: OoxmlParts): number {
  return relationshipEdges(p).filter((e) => !p.names.includes(e.to)).length;
}

export const CONSTRUCTS: PptxConstruct[] = [
  {
    id: 'slides.multiple',
    label: 'Several slides',
    count: (p) => (countNames(p, SLIDE) > 1 ? countNames(p, SLIDE) : 0),
  },
  {
    id: 'text.basic',
    label: 'Text runs',
    count: (p) => inSlides(p, /<a:t>[^<]+<\/a:t>/g),
  },
  {
    id: 'text.runs',
    label: 'Run formatting (bold, italic, underline, strike)',
    count: (p) =>
      inSlides(p, el('a:rPr', '(?:b|i)="(?:1|true)"|u="(?!none")|strike="(?!noStrike")')),
  },
  {
    id: 'text.alignment',
    label: 'Paragraph alignment other than left',
    count: (p) => inSlides(p, el('a:pPr', 'algn="(?:ctr|r|just|dist)"')),
  },
  {
    id: 'text.levels',
    label: 'Paragraph levels (lvl 1 to 8)',
    count: (p) => inSlides(p, el('a:pPr', 'lvl="[1-8]"')),
  },
  { id: 'text.bullets', label: 'Bullet characters', count: (p) => inSlides(p, tagRe('buChar')) },
  {
    id: 'text.numbering',
    label: 'Auto-numbered paragraphs',
    count: (p) => inSlides(p, tagRe('buAutoNum')),
  },
  {
    id: 'text.bulletsCustom',
    label: 'Bullet colour or size',
    count: (p) => inSlides(p, tagRe('buClr', 'buSzPct', 'buSzPts')),
  },
  {
    id: 'text.bulletsPicture',
    label: 'Picture bullets',
    count: (p) => inSlides(p, tagRe('buBlip')),
  },
  {
    id: 'text.hyperlinks',
    label: 'Hyperlinks on text runs',
    count: (p) => inSlides(p, new RegExp(`${RUN_PROPS}<a:hlinkClick\\b`, 'g')),
  },
  {
    id: 'text.spacing',
    label: 'Paragraph spacing',
    count: (p) => inSlides(p, tagRe('spcBef', 'spcAft', 'lnSpc')),
  },
  {
    id: 'text.highlight',
    label: 'Highlighted text',
    count: (p) => inSlides(p, tagRe('highlight')),
  },
  { id: 'text.tabStops', label: 'Tab stops', count: (p) => inSlides(p, /<a:tab\b/g) },
  {
    id: 'text.columns',
    label: 'Text columns',
    count: (p) => inSlides(p, el('a:bodyPr', 'numCol="(?:[2-9]|[1-9]\\d)"')),
  },
  {
    id: 'text.vertical',
    label: 'Vertical text direction',
    count: (p) => inSlides(p, el('a:bodyPr', 'vert="(?!horz")[^"]+"')),
  },
  {
    id: 'text.autofit',
    label: 'Shrink-on-overflow autofit (fontScale)',
    count: (p) => inSlides(p, el('a:normAutofit', '(?:fontScale|lnSpcReduction)="')),
  },
  {
    id: 'text.autofitShape',
    label: 'Resize shape to fit text',
    count: (p) => inSlides(p, tagRe('spAutoFit')),
  },
  {
    id: 'text.insets',
    label: 'Custom text insets',
    count: (p) => inSlides(p, el('a:bodyPr', '[lrtb]Ins="')),
  },
  {
    id: 'text.fields',
    label: 'Fields (slide number, date)',
    count: (p) => inSlides(p, tagRe('fld')),
  },
  {
    id: 'text.baseline',
    label: 'Superscript and subscript',
    count: (p) => inSlides(p, el('a:rPr', 'baseline="-?[1-9]')),
  },
  {
    id: 'text.caps',
    label: 'All caps and small caps',
    count: (p) => inSlides(p, el('a:rPr', 'cap="(?:all|small)"')),
  },
  {
    id: 'text.charSpacing',
    label: 'Character spacing',
    count: (p) => inSlides(p, el('a:rPr', 'spc="-?[1-9]')),
  },
  {
    id: 'masters.multiple',
    label: 'More than one slide master',
    count: (p) => (countNames(p, MASTER) > 1 ? countNames(p, MASTER) : 0),
  },
  {
    id: 'masters.handoutMaster',
    label: 'Handout master',
    count: (p) => countNames(p, /^ppt\/handoutMasters\/handoutMaster\d+\.xml$/),
  },
  {
    id: 'masters.txStyles',
    label: 'Master text styles',
    count: (p) => inParts(p, MASTER, /<p:txStyles>/g),
  },
  {
    id: 'masters.background',
    label: 'Master with its own background',
    count: (p) => inParts(p, MASTER, /<p:bgPr\b/g),
  },
  {
    id: 'masters.shapes',
    label: 'Non-placeholder shapes on a master',
    count: (p) => Math.max(0, inParts(p, MASTER, /<p:sp>/g) - inParts(p, MASTER, /<p:ph\b/g)),
  },
  {
    id: 'layouts.placeholders',
    label: 'Layout placeholders',
    count: (p) => inParts(p, LAYOUT, /<p:ph\b/g),
  },
  {
    id: 'layouts.placeholderKinds',
    label: 'Distinct placeholder types',
    count: (p) => placeholderTypes(p).size,
  },
  {
    id: 'layouts.background',
    label: 'Layout with its own background',
    count: (p) => inParts(p, LAYOUT, /<p:bg>/g),
  },
  {
    id: 'layouts.hideMasterShapes',
    label: 'Layout that hides master shapes',
    count: (p) => inParts(p, LAYOUT, el('p:sldLayout', 'showMasterSp="(?:0|false)"')),
  },
  {
    id: 'layouts.inheritance',
    label: 'Slide placeholders that inherit their position',
    count: (p) =>
      inSlides(
        p,
        /<p:ph\b[^>]*\/>\s*<\/p:nvPr>\s*<\/p:nvSpPr>\s*<p:spPr(?:\s*\/>|\s*>\s*<\/p:spPr>)/g
      ),
  },
  {
    id: 'layouts.footerPlaceholders',
    label: 'Date, footer and slide-number placeholders on slides',
    count: (p) => inSlides(p, el('p:ph', 'type="(?:dt|ftr|sldNum)"')),
  },
  {
    id: 'theme.multiple',
    label: 'Masters with different themes',
    count: (p) => (masterThemes(p).size > 1 ? masterThemes(p).size : 0),
  },
  {
    id: 'theme.colorMap',
    label: 'Colour map that is not the default',
    count: (p) =>
      inParts(
        p,
        MASTER,
        el('p:clrMap', 'bg1="(?!lt1")|tx1="(?!dk1")|bg2="(?!lt2")|tx2="(?!dk2")')
      ) + inParts(p, /^ppt\/(?:slides|slideLayouts)\/[^/]+\.xml$/, tagRe('overrideClrMapping')),
  },
  {
    id: 'theme.fontScheme',
    label: 'Theme with its own font scheme',
    count: (p) => inParts(p, THEME, /<a:(?:major|minor)Font>\s*<a:latin typeface="(?!Calibri)/g),
  },
  {
    id: 'theme.fontsEastAsian',
    label: 'Theme East Asian fonts',
    count: (p) => inParts(p, THEME, /<a:ea typeface="[^"]/g),
  },
  {
    id: 'theme.effectScheme',
    label: 'Theme effect styles with effects',
    count: (p) =>
      inParts(
        p,
        THEME,
        /<a:effectStyle>\s*<a:effectLst>\s*<a:(?:outerShdw|innerShdw|prstShdw|glow|reflection|softEdge|blur)\b/g
      ),
  },
  {
    id: 'fills.themeColors',
    label: 'Theme colours',
    count: (p) => inSlides(p, tagRe('schemeClr')),
  },
  {
    id: 'fills.colorTransforms',
    label: 'Colour transforms (lumMod, tint, shade, satMod)',
    count: (p) => inSlides(p, tagRe('lumMod', 'lumOff', 'tint', 'shade', 'satMod', 'hueMod')),
  },
  { id: 'fills.solid', label: 'Solid shape fills', count: (p) => inFills(p, tagRe('solidFill')) },
  {
    id: 'fills.noFill',
    label: 'Shapes with no fill',
    count: (p) => inFills(p, /<a:noFill\s*\/>/g),
  },
  { id: 'fills.gradient', label: 'Gradient fills', count: (p) => inFills(p, tagRe('gradFill')) },
  {
    id: 'fills.gradientPath',
    label: 'Path (radial, rectangular) gradients',
    count: (p) => inFills(p, /<a:gradFill\b(?:(?!<\/a:gradFill>)[\s\S])*?<a:path\b/g),
  },
  { id: 'fills.pattern', label: 'Pattern fills', count: (p) => inFills(p, tagRe('pattFill')) },
  { id: 'fills.picture', label: 'Picture fills', count: (p) => inFills(p, tagRe('blipFill')) },
  { id: 'fills.alpha', label: 'Transparency (alpha)', count: (p) => inSlides(p, tagRe('alpha')) },
  { id: 'fills.background', label: 'Slide backgrounds', count: (p) => inSlides(p, /<p:bg>/g) },
  {
    id: 'lines.dash',
    label: 'Dashed lines',
    count: (p) => inLines(p, el('a:prstDash', 'val="(?!solid")')),
  },
  {
    id: 'lines.arrowheads',
    label: 'Arrowheads',
    count: (p) => inLines(p, el('a:(?:headEnd|tailEnd)', 'type="(?!none")')),
  },
  {
    id: 'lines.joinCap',
    label: 'Line caps and joins',
    count: (p) => inLines(p, /<a:ln\b[^>]*\bcap="(?:rnd|sq)"|<a:(?:bevel|miter)\b/g),
  },
  {
    id: 'lines.compound',
    label: 'Compound lines',
    count: (p) => inLines(p, el('a:ln', 'cmpd="(?!sng")')),
  },
  {
    id: 'effects.shadow',
    label: 'Shadows',
    count: (p) => inSlides(p, tagRe('outerShdw', 'innerShdw', 'prstShdw')),
  },
  { id: 'effects.glow', label: 'Glow', count: (p) => inSlides(p, tagRe('glow')) },
  { id: 'effects.reflection', label: 'Reflection', count: (p) => inSlides(p, tagRe('reflection')) },
  { id: 'effects.softEdge', label: 'Soft edges', count: (p) => inSlides(p, tagRe('softEdge')) },
  { id: 'effects.blur', label: 'Blur', count: (p) => inSlides(p, tagRe('blur')) },
  {
    id: 'effects.scene3d',
    label: '3-D scene (camera, lights)',
    count: (p) => inSlides(p, tagRe('scene3d')),
  },
  {
    id: 'effects.sp3d',
    label: '3-D shape (extrusion, material)',
    count: (p) => inSlides(p, tagRe('sp3d')),
  },
  { id: 'effects.bevel', label: 'Bevels', count: (p) => inSlides(p, tagRe('bevelT', 'bevelB')) },
  {
    id: 'effects.text',
    label: 'Text effects',
    count: (p) => inSlides(p, new RegExp(`${RUN_PROPS}<a:effectLst\\b`, 'g')),
  },
  { id: 'shapes.presets', label: 'Preset geometry shapes', count: (p) => presetNames(p).length },
  {
    id: 'shapes.presetKinds',
    label: 'Distinct preset geometries',
    count: (p) => new Set(presetNames(p)).size,
  },
  {
    id: 'shapes.adjustments',
    label: 'Preset geometries with adjustments',
    count: (p) =>
      partsOf(p, SLIDE).reduce(
        (sum, x) => sum + count(x, /<a:prstGeom\b[^>]*>\s*<a:avLst>\s*<a:gd\b/g),
        0
      ),
  },
  { id: 'shapes.custGeom', label: 'Custom geometry', count: (p) => inSlides(p, tagRe('custGeom')) },
  {
    id: 'shapes.flip',
    label: 'Flipped shapes',
    count: (p) => inSlides(p, el('a:xfrm', 'flip[HV]="(?:1|true)"')),
  },
  {
    id: 'shapes.rotation',
    label: 'Rotated shapes',
    count: (p) => inSlides(p, el('a:xfrm', 'rot="-?[1-9]')),
  },
  {
    id: 'shapes.textBox',
    label: 'Text boxes',
    count: (p) => inSlides(p, el('p:cNvSpPr', 'txBox="(?:1|true)"')),
  },
  { id: 'sizes.4x3', label: 'Slide size 4:3', count: sizeIs('4x3') },
  { id: 'sizes.16x9', label: 'Slide size 16:9', count: sizeIs('16x9') },
  { id: 'sizes.a4', label: 'Slide size A4', count: sizeIs('a4') },
  { id: 'sizes.custom', label: 'Custom slide size', count: sizeIs('custom') },
  { id: 'sizes.portrait', label: 'Portrait slides', count: sizeIs('portrait') },
  {
    id: 'edge.truncatedZip',
    label: 'Truncated zip (starts as a zip, cannot be read)',
    count: (p) =>
      p.kind === 'unreadable' &&
      p.raw[0] === 0x50 &&
      p.raw[1] === 0x4b &&
      p.raw[2] === 0x03 &&
      p.raw[3] === 0x04
        ? 1
        : 0,
  },
  {
    id: 'edge.encrypted',
    label: 'Encrypted package (OLE with EncryptionInfo)',
    count: (p) => (p.kind === 'ole' && ole(p, 'EncryptionInfo') ? 1 : 0),
  },
  {
    id: 'edge.legacyBinary',
    label: 'Legacy .ppt binary (OLE with a PowerPoint Document stream)',
    count: (p) => (p.kind === 'ole' && ole(p, 'PowerPoint Document') ? 1 : 0),
  },
  {
    id: 'edge.noPresentation',
    label: 'Package without ppt/presentation.xml',
    count: (p) =>
      p.kind === 'zip' &&
      p.names.includes('[Content_Types].xml') &&
      !p.names.includes('ppt/presentation.xml')
        ? 1
        : 0,
  },
  {
    id: 'edge.slideMissing',
    label: 'Listed slide whose part is missing',
    count: missingSlideParts,
  },
  { id: 'edge.brokenRels', label: 'Relationships that point nowhere', count: brokenRels },
  {
    id: 'edge.invalidXml',
    label: 'Part that is not well-formed XML',
    count: (p) => Object.values(p.xml).filter((x) => !isWellFormed(x)).length,
  },
  {
    id: 'edge.manySlides',
    label: '1000 or more slides declared',
    count: (p) => (count(p.xml['ppt/presentation.xml'] ?? '', /<p:sldId\b/g) >= 1000 ? 1 : 0),
  },
  {
    id: 'edge.deepGroups',
    label: 'Groups nested 20 levels deep or more',
    count: (p) => partsOf(p, SLIDE).filter((x) => maxGroupDepth(x) >= 20).length,
  },
  { id: 'edge.relCycle', label: 'Cycles in the relationship graph', count: relationshipCycles },
];

export function findConstruct(id: string): PptxConstruct | undefined {
  return CONSTRUCTS.find((c) => c.id === id);
}
