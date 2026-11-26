import { CONSTRUCTS, findConstruct } from '../../../../scripts/pptx-fidelity/constructs';
import type { OoxmlParts } from '../../../../scripts/fidelity-core/ooxml-package';

type Xml = Record<string, string>;
const mk = (xml: Xml = {}, extra: Partial<OoxmlParts> = {}): OoxmlParts => ({
  kind: 'zip',
  raw: new Uint8Array(),
  names: Object.keys(xml),
  xml,
  sizes: {},
  oversized: [],
  ...extra,
});
const slide = (body: string, n = 1): Xml => ({
  [`ppt/slides/slide${n}.xml`]: `<p:sld>${body}</p:sld>`,
});
const spPr = (body: string) => slide(`<p:spPr>${body}</p:spPr>`);
const master = (body: string, n = 1): Xml => ({
  [`ppt/slideMasters/slideMaster${n}.xml`]: `<p:sldMaster>${body}</p:sldMaster>`,
});
const layout = (body: string, attrs = ''): Xml => ({
  'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout${attrs}>${body}</p:sldLayout>`,
});
const theme = (body: string): Xml => ({ 'ppt/theme/theme1.xml': `<a:theme>${body}</a:theme>` });
const pres = (body: string): Xml => ({
  'ppt/presentation.xml': `<p:presentation>${body}</p:presentation>`,
});
const names = (...n: string[]): OoxmlParts => mk({}, { names: n });
const ole = (name: string) =>
  mk({}, { kind: 'ole', raw: new Uint8Array(Buffer.from(`HEADER${name}`, 'utf16le')) });
const rel = (id: string, type: string, target: string, extra = '') =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"${extra}/>`;
const rels = (name: string, ...items: string[]): Xml => ({
  [name]: `<Relationships>${items.join('')}</Relationships>`,
});
const masterTheme = (n: number, theme: number): Xml =>
  rels(
    `ppt/slideMasters/_rels/slideMaster${n}.xml.rels`,
    rel('rId1', 'theme', `../theme/theme${theme}.xml`)
  );
const size = (attrs: string) => mk(pres(`<p:sldSz ${attrs}/>`));
const chart = (body: string): Xml => ({
  'ppt/charts/chart1.xml': `<c:chartSpace>${body}</c:chartSpace>`,
});
const notesPart = (body: string): Xml => ({
  'ppt/notesSlides/notesSlide1.xml': `<p:notes>${body}</p:notes>`,
});
const twoSlides = (a: string, b: string): Xml => ({ ...slide(a, 1), ...slide(b, 2) });
const cxn = (...blocks: string[]) =>
  mk(slide(blocks.map((b) => `<p:cxnSp>${b}</p:cxnSp>`).join('')));
const GROUP_XFRM = (ext: string, chExt: string) =>
  `<p:grpSpPr><a:xfrm>${ext}${chExt}</a:xfrm></p:grpSpPr>`;
const clrMap = (attrs: string, n: number) => master(`<p:clrMap ${attrs}/>`, n);

/** One minimal package per construct, with the count the probe must return. */
const POSITIVE: Array<[string, OoxmlParts, number]> = [
  ['slides.multiple', mk({ ...slide('', 1), ...slide('', 2) }), 2],
  ['text.basic', mk(slide('<a:t>Hi</a:t><a:t></a:t><a:t>Ho</a:t>')), 2],
  [
    'text.runs',
    mk(
      slide(
        '<a:rPr b="1"/><a:rPr lang="x" i="true"/><a:rPr u="sng"/><a:rPr strike="sngStrike"/><a:rPr b="0" u="none" strike="noStrike"/><a:rPr lang="en-US"/>'
      )
    ),
    4,
  ],
  [
    'text.alignment',
    mk(slide('<a:pPr algn="ctr"/><a:pPr algn="l"/><a:pPr lvl="1" algn="just"/>')),
    2,
  ],
  [
    'text.levels',
    mk(slide('<a:pPr lvl="2"/><a:pPr lvl="0"/><a:pPr marL="1" lvl="8"/><a:pPr lvl="9"/>')),
    2,
  ],
  ['text.bullets', mk(slide('<a:buChar char="x"/><a:buNone/>')), 1],
  ['text.numbering', mk(slide('<a:buAutoNum type="arabicPeriod"/>')), 1],
  [
    'text.bulletsCustom',
    mk(slide('<a:buClr/><a:buSzPct val="1"/><a:buSzPts val="1"/><a:buFont/>')),
    3,
  ],
  ['text.bulletsPicture', mk(slide('<a:buBlip><a:blip/></a:buBlip>')), 1],
  [
    'text.hyperlinks',
    mk(
      slide(
        '<a:rPr lang="x"><a:hlinkClick r:id="rId1"/></a:rPr><a:rPr/><a:cNvPr id="1"><a:hlinkClick r:id="rId2"/></a:cNvPr><a:rPr><a:solidFill/><a:hlinkClick/></a:rPr>'
      )
    ),
    2,
  ],
  ['text.spacing', mk(slide('<a:lnSpc/><a:spcBef/><a:spcAft/>')), 3],
  ['text.highlight', mk(slide('<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>')), 1],
  ['text.tabStops', mk(slide('<a:tabLst><a:tab pos="1" algn="l"/></a:tabLst>')), 1],
  [
    'text.columns',
    mk(slide('<a:bodyPr numCol="2"/><a:bodyPr numCol="1"/><a:bodyPr numCol="12"/>')),
    2,
  ],
  ['text.vertical', mk(slide('<a:bodyPr vert="vert270"/><a:bodyPr vert="horz"/><a:bodyPr/>')), 1],
  [
    'text.autofit',
    mk(
      slide(
        '<a:normAutofit fontScale="50000"/><a:normAutofit/><a:normAutofit lnSpcReduction="10000"/>'
      )
    ),
    2,
  ],
  ['text.autofitShape', mk(slide('<a:spAutoFit/><a:noAutofit/>')), 1],
  [
    'text.insets',
    mk(slide('<a:bodyPr lIns="0"/><a:bodyPr wrap="square" bIns="5"/><a:bodyPr/>')),
    2,
  ],
  ['text.fields', mk(slide('<a:fld id="{1}" type="slidenum"><a:t>1</a:t></a:fld>')), 1],
  [
    'text.baseline',
    mk(slide('<a:rPr baseline="-25000"/><a:rPr baseline="0"/><a:rPr baseline="30000"/>')),
    2,
  ],
  ['text.caps', mk(slide('<a:rPr cap="all"/><a:rPr cap="none"/><a:rPr cap="small"/>')), 2],
  ['text.charSpacing', mk(slide('<a:rPr spc="300"/><a:rPr spc="0"/><a:rPr spc="-100"/>')), 2],
  ['masters.multiple', mk({ ...master('', 1), ...master('', 2), ...master('', 3) }), 3],
  ['masters.handoutMaster', names('ppt/handoutMasters/handoutMaster1.xml'), 1],
  ['masters.txStyles', mk(master('<p:txStyles></p:txStyles>')), 1],
  ['masters.background', mk(master('<p:bg><p:bgPr><a:noFill/></p:bgPr></p:bg>')), 1],
  [
    'masters.shapes',
    mk(
      master(
        '<p:sp><p:nvPr><p:ph type="title"/></p:nvPr></p:sp><p:sp><p:nvPr/></p:sp><p:sp></p:sp>'
      )
    ),
    2,
  ],
  ['layouts.placeholders', mk(layout('<p:ph type="title"/><p:ph idx="1"/>')), 2],
  [
    'layouts.placeholderKinds',
    mk({
      ...layout('<p:ph type="title"/><p:ph idx="1"/><p:ph type="dt" idx="2"/>'),
      ...master('<p:ph type="title"/>'),
    }),
    3,
  ],
  ['layouts.background', mk(layout('<p:bg><p:bgPr/></p:bg>')), 1],
  ['layouts.hideMasterShapes', mk(layout('', ' type="blank" showMasterSp="0"')), 1],
  [
    'layouts.inheritance',
    mk(
      slide(
        '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:nvPr><p:ph idx="1"/>\n</p:nvPr>\n</p:nvSpPr>\n<p:spPr>\n</p:spPr><p:nvPr><p:ph idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm/></p:spPr>'
      )
    ),
    2,
  ],
  [
    'layouts.footerPlaceholders',
    mk(
      slide(
        '<p:ph type="dt" idx="10"/><p:ph sz="quarter" type="ftr"/><p:ph type="sldNum"/><p:ph type="title"/>'
      )
    ),
    3,
  ],
  ['theme.multiple', mk({ ...masterTheme(1, 1), ...masterTheme(2, 2), ...masterTheme(3, 2) }), 2],
  [
    'theme.colorMap',
    mk({
      ...clrMap('bg1="dk1" tx1="dk1" bg2="lt2" tx2="dk2"', 1),
      ...clrMap('bg1="lt1" tx1="lt1" bg2="lt2" tx2="dk2"', 2),
      ...clrMap('bg1="lt1" tx1="dk1" bg2="dk2" tx2="dk2"', 3),
      ...clrMap('tx2="lt2" bg1="lt1" tx1="dk1" bg2="lt2"', 4),
      ...clrMap('bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"', 5),
      'ppt/slides/slide1.xml':
        '<p:sld><a:overrideClrMapping bg1="dk1"/><a:masterClrMapping/></p:sld>',
      'ppt/slideLayouts/slideLayout1.xml':
        '<p:sldLayout><a:overrideClrMapping bg1="dk1"/></p:sldLayout>',
    }),
    6,
  ],
  [
    'theme.fontScheme',
    mk(
      theme(
        '<a:majorFont><a:latin typeface="Georgia"/></a:majorFont><a:minorFont>\n<a:latin typeface="Calibri"/></a:minorFont>'
      )
    ),
    1,
  ],
  ['theme.fontsEastAsian', mk(theme('<a:ea typeface="MS Gothic"/><a:ea typeface=""/>')), 1],
  [
    'theme.effectScheme',
    mk(
      theme(
        '<a:effectStyle><a:effectLst><a:outerShdw/></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle>\n<a:effectLst>\n<a:softEdge rad="1"/></a:effectLst></a:effectStyle>'
      )
    ),
    2,
  ],
  ['fills.themeColors', mk(slide('<a:schemeClr val="accent1"/><a:srgbClr val="000000"/>')), 1],
  [
    'fills.colorTransforms',
    mk(slide('<a:lumMod val="1"/><a:tint val="1"/><a:alpha val="1"/><a:hueMod val="1"/>')),
    3,
  ],
  ['fills.solid', mk(spPr('<a:solidFill/><a:ln><a:solidFill/></a:ln>')), 1],
  ['fills.noFill', mk(spPr('<a:noFill/><a:ln><a:noFill/></a:ln>')), 1],
  ['fills.gradient', mk(spPr('<a:gradFill><a:gsLst/></a:gradFill><a:ln><a:gradFill/></a:ln>')), 1],
  [
    'fills.gradientPath',
    mk(
      spPr(
        '<a:gradFill><a:path path="circle"/></a:gradFill><a:gradFill><a:lin/></a:gradFill><a:custGeom><a:path w="1"/></a:custGeom>'
      )
    ),
    1,
  ],
  ['fills.pattern', mk(spPr('<a:pattFill prst="pct5"/>')), 1],
  ['fills.picture', mk(spPr('<a:blipFill><a:blip/></a:blipFill>')), 1],
  ['fills.alpha', mk(slide('<a:alpha val="50000"/>')), 1],
  ['fills.background', mk(slide('<p:bg><p:bgPr/></p:bg>')), 1],
  [
    'lines.dash',
    mk(
      slide(
        '<p:spPr><a:ln><a:prstDash val="dash"/></a:ln></p:spPr><p:spPr><a:ln><a:prstDash val="solid"/></a:ln></p:spPr>'
      )
    ),
    1,
  ],
  [
    'lines.arrowheads',
    mk(spPr('<a:ln><a:headEnd type="none"/><a:tailEnd w="med" type="triangle"/></a:ln>')),
    1,
  ],
  [
    'lines.joinCap',
    mk(spPr('<a:ln cap="sq"><a:bevel/></a:ln><a:ln cap="flat"><a:round/></a:ln>')),
    2,
  ],
  ['lines.compound', mk(spPr('<a:ln cmpd="dbl"><a:solidFill/></a:ln><a:ln cmpd="sng"></a:ln>')), 1],
  ['effects.shadow', mk(slide('<a:outerShdw/><a:innerShdw/><a:prstShdw prst="shdw1"/>')), 3],
  ['effects.glow', mk(slide('<a:glow rad="1"/>')), 1],
  ['effects.reflection', mk(slide('<a:reflection/>')), 1],
  ['effects.softEdge', mk(slide('<a:softEdge rad="1"/>')), 1],
  ['effects.blur', mk(slide('<a:blur rad="1"/>')), 1],
  ['effects.scene3d', mk(slide('<a:scene3d><a:camera/></a:scene3d>')), 1],
  ['effects.sp3d', mk(slide('<a:sp3d/>')), 1],
  ['effects.bevel', mk(slide('<a:sp3d><a:bevelT/><a:bevelB/></a:sp3d>')), 2],
  ['effects.text', mk(slide('<a:rPr lang="x"><a:effectLst/></a:rPr><a:rPr/><a:effectLst/>')), 1],
  [
    'shapes.presets',
    mk(
      slide('<a:prstGeom prst="rect"/><a:prstGeom prst="rect"/><a:prstGeom fmla="x" prst="star5"/>')
    ),
    3,
  ],
  [
    'shapes.presetKinds',
    mk(
      slide('<a:prstGeom prst="rect"/><a:prstGeom prst="rect"/><a:prstGeom fmla="x" prst="star5"/>')
    ),
    2,
  ],
  [
    'shapes.adjustments',
    mk(
      slide(
        '<a:prstGeom prst="a"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst></a:prstGeom><a:prstGeom prst="b"><a:avLst/></a:prstGeom>'
      )
    ),
    1,
  ],
  ['shapes.custGeom', mk(slide('<a:custGeom><a:pathLst/></a:custGeom>')), 1],
  [
    'shapes.flip',
    mk(slide('<a:xfrm flipH="1"/><a:xfrm rot="5" flipV="true"/><a:xfrm flipH="0"/>')),
    2,
  ],
  [
    'shapes.rotation',
    mk(slide('<a:xfrm rot="5400000"/><a:xfrm rot="0"/><a:xfrm rot="-60000"/>')),
    2,
  ],
  ['shapes.textBox', mk(slide('<p:cNvSpPr txBox="1"/><p:cNvSpPr/><p:cNvSpPr txBox="0"/>')), 1],
  [
    'images.png',
    names('ppt/media/a.png', 'ppt/media/b.PNG', 'ppt/media/c.png.bak', 'ppt/other/d.png'),
    2,
  ],
  ['images.jpeg', names('ppt/media/a.jpg', 'ppt/media/b.jpeg', 'ppt/media/c.jpgx'), 2],
  ['images.gif', names('ppt/media/a.gif'), 1],
  ['images.svg', names('ppt/media/a.svg', 'ppt/media/a.png'), 1],
  ['images.emf', names('ppt/media/a.emf'), 1],
  ['images.wmf', names('ppt/media/a.wmf'), 1],
  [
    'images.crop',
    mk(slide('<a:srcRect l="25000" t="0"/><a:srcRect/><a:srcRect b="-5000"/><a:srcRect r="0"/>')),
    2,
  ],
  [
    'images.transform',
    mk(
      slide(
        '<p:pic><a:xfrm rot="5400000"/></p:pic><p:pic><a:xfrm flipH="1"/></p:pic><p:pic><a:xfrm rot="0"/></p:pic><a:xfrm rot="5400000"/>'
      )
    ),
    2,
  ],
  ['images.linked', mk(slide('<a:blip r:link="rId1"/><a:blip r:embed="rId2"/>')), 1],
  [
    'images.altText',
    mk(
      slide(
        '<p:pic><p:cNvPr id="1" descr="x"/></p:pic><p:pic><p:cNvPr id="2" descr=""/></p:pic><p:pic><p:cNvPr id="3"/></p:pic><p:sp><p:cNvPr descr="y"/></p:sp>'
      )
    ),
    1,
  ],
  [
    'images.recolor',
    mk(
      slide(
        '<a:duotone/><a:grayscl/><a:lum bright="1"/><a:lumMod val="1"/><a:clrChange/><a:biLevel/>'
      )
    ),
    5,
  ],
  ['tables.basic', mk(slide('<a:tbl><a:tblPr/></a:tbl>')), 1],
  [
    'tables.merge',
    mk(
      slide(
        '<a:tc gridSpan="2"/><a:tc rowSpan="3"/><a:tc hMerge="1"/><a:tc vMerge="true"/><a:tc gridSpan="1"/><a:tc/>'
      )
    ),
    4,
  ],
  [
    'tables.banding',
    mk(slide('<a:tblPr bandRow="1"/><a:tblPr firstRow="1" bandCol="true"/><a:tblPr bandRow="0"/>')),
    2,
  ],
  [
    'tables.headerRow',
    mk(slide('<a:tblPr firstRow="1"/><a:tblPr lastCol="1"/><a:tblPr bandRow="1"/>')),
    2,
  ],
  ['tables.styleRef', mk(slide('<a:tableStyleId>{5C22544A}</a:tableStyleId>')), 1],
  [
    'tables.customStyles',
    mk({
      'ppt/tableStyles.xml':
        '<a:tblStyleLst def="x"><a:tblStyle styleId="1"/><a:tblStyle styleId="2"/></a:tblStyleLst>',
    }),
    2,
  ],
  [
    'tables.borders',
    mk(slide('<a:lnL/><a:lnR/><a:lnT/><a:lnB/><a:lnTlToBr/><a:lnBlToTr/><a:ln/>')),
    6,
  ],
  [
    'tables.cellFill',
    mk(
      slide(
        '<a:tcPr anchor="ctr"><a:solidFill/></a:tcPr><a:tcPr><a:lnL/></a:tcPr><a:tcPr/><a:tcPr><a:noFill/></a:tcPr>'
      )
    ),
    1,
  ],
  ['groups.basic', mk(slide('<p:grpSp><p:grpSpPr/></p:grpSp><p:grpSp>')), 2],
  [
    'groups.nested',
    mk(
      twoSlides('<p:grpSp><p:grpSp></p:grpSp></p:grpSp>', '<p:grpSp></p:grpSp><p:grpSp></p:grpSp>')
    ),
    1,
  ],
  [
    'groups.scaled',
    mk(
      slide(
        [
          GROUP_XFRM('<a:ext cx="0" cy="0"/>', '<a:chExt cx="0" cy="0"/>'),
          GROUP_XFRM('<a:ext cx="10" cy="10"/>', '<a:chExt cx="20" cy="20"/>'),
          GROUP_XFRM('<a:ext cx="10" cy="10"/>', '<a:chExt cx="10" cy="20"/>'),
          GROUP_XFRM('<a:ext cx="10" cy="10"/>', ''),
          '<p:grpSpPr/>',
          GROUP_XFRM('', '<a:chExt cx="1" cy="1"/>'),
        ].join('')
      )
    ),
    2,
  ],
  [
    'connectors.straight',
    cxn(
      '<a:prstGeom prst="straightConnector1"/>',
      '<a:prstGeom prst="line"/>',
      '<a:prstGeom prst="bentConnector3"/>'
    ),
    2,
  ],
  [
    'connectors.elbow',
    cxn(
      '<a:prstGeom prst="bentConnector3"/>',
      '<a:prstGeom prst="bentConnector2"/>',
      '<a:prstGeom prst="curvedConnector3"/>'
    ),
    2,
  ],
  [
    'connectors.curved',
    cxn(
      '<a:prstGeom prst="curvedConnector3"/>',
      '<a:prstGeom prst="curvedConnector5"/>',
      '<a:prstGeom prst="bentConnector3"/>'
    ),
    2,
  ],
  [
    'connectors.connected',
    cxn('<a:stCxn id="1" idx="0"/>', '<a:endCxn id="2" idx="1"/>', '<a:prstGeom prst="line"/>'),
    2,
  ],
  ['charts.bar', mk(chart('<c:barChart><c:ser/></c:barChart><c:barChart/>')), 2],
  ['charts.line', mk(chart('<c:lineChart/>')), 1],
  ['charts.pie', mk(chart('<c:pieChart/>')), 1],
  ['charts.scatter', mk(chart('<c:scatterChart/>')), 1],
  ['charts.cache', mk(chart('<c:numCache><c:pt/></c:numCache><c:numRef/>')), 1],
  [
    'charts.embeddedWorkbook',
    mk({
      ...rels(
        'ppt/charts/_rels/chart1.xml.rels',
        rel('rId1', 'package', '../embeddings/a.xlsx'),
        rel('rId2', 'themeOverride', '../theme/themeOverride1.xml')
      ),
      ...rels('ppt/slides/_rels/slide1.xml.rels', rel('rId1', 'package', '../embeddings/b.xlsx')),
    }),
    1,
  ],
  ['smartart.diagram', mk(slide('<dgm:relIds r:dm="rId1" r:lo="rId2"/>')), 1],
  ['smartart.drawing', names('ppt/diagrams/drawing1.xml', 'ppt/diagrams/data1.xml'), 1],
  ['media.audio', mk(slide('<a:audioFile r:link="rId1"/>')), 1],
  ['media.video', mk(slide('<a:videoFile r:link="rId1"/>')), 1],
  [
    'transitions.basic',
    mk(slide('<p:transition spd="med"/><mc:Choice Requires="p14"><p:transition/></mc:Choice>')),
    2,
  ],
  [
    'transitions.timed',
    mk(slide('<p:transition advTm="5000"/><p:transition advClick="0"/><p:transition spd="med"/>')),
    2,
  ],
  [
    'transitions.extended',
    mk(
      slide(
        '<mc:Choice xmlns:p14="x" Requires="p14"><p:transition/></mc:Choice><mc:Choice Requires="p14" xmlns:p14="x">\n<p:transition spd="slow"/></mc:Choice><mc:Choice Requires="a14"><p:transition/></mc:Choice>'
      )
    ),
    2,
  ],
  ['animations.timing', mk(slide('<p:timing><p:tnLst/></p:timing>')), 1],
  [
    'animations.entrance',
    mk(slide('<p:cTn id="1" presetClass="entr"/><p:cTn presetClass="exit"/>')),
    1,
  ],
  [
    'animations.exit',
    mk(slide('<p:cTn id="1" presetClass="entr"/><p:cTn presetClass="exit"/>')),
    1,
  ],
  [
    'animations.emphasis',
    mk(slide('<p:cTn presetClass="emph" id="2"/><p:cTn presetClass="path"/>')),
    1,
  ],
  ['animations.motionPath', mk(slide('<p:animMotion origin="layout"/>')), 1],
  [
    'animations.build',
    mk(
      slide(
        '<p:bldP spid="3" build="p"/><p:bldP spid="4" build="whole"/><p:bldP spid="5" grpId="0"/>'
      )
    ),
    1,
  ],
  [
    'notes.slides',
    names(
      'ppt/notesSlides/notesSlide1.xml',
      'ppt/notesSlides/_rels/notesSlide1.xml.rels',
      'ppt/notesSlides/notesSlide2.xml'
    ),
    2,
  ],
  ['notes.master', names('ppt/notesMasters/notesMaster1.xml'), 1],
  [
    'notes.rich',
    mk(notesPart('<a:rPr b="1"/><a:rPr lang="x"/><a:buChar char="x"/><a:buAutoNum type="a"/>')),
    3,
  ],
  [
    'comments.legacy',
    mk({
      'ppt/comments/comment1.xml': '<p:cmLst><p:cm authorId="0"/><p:cm authorId="1"/></p:cmLst>',
    }),
    2,
  ],
  [
    'comments.modern',
    mk({
      'ppt/comments/modernComment_1.xml':
        '<p188:cmLst><p188:cm id="1"><p188:replyLst/></p188:cm></p188:cmLst>',
    }),
    1,
  ],
  [
    'comments.replies',
    mk({
      'ppt/comments/modernComment_1.xml':
        '<p188:cm><p188:reply id="1"/><p188:reply id="2"/><p188:replyLst/></p188:cm>',
    }),
    2,
  ],
  [
    'comments.authors',
    names('ppt/commentAuthors.xml', 'ppt/authors.xml', 'ppt/other/authors.xml'),
    2,
  ],
  [
    'slides.hidden',
    mk({
      'ppt/slides/slide1.xml': '<p:sld show="0"/>',
      'ppt/slides/slide2.xml': '<p:sld showMasterSp="0" show="false"/>',
      'ppt/slides/slide3.xml': '<p:sld show="1"/>',
      'ppt/slides/slide4.xml': '<p:sld/>',
    }),
    2,
  ],
  ['sections.list', mk(pres('<p14:section name="a"/><p14:section name="b"/><p14:sectionLst/>')), 2],
  [
    'sections.customShows',
    mk(pres('<p:custShowLst><p:custShow name="a"/><p:custShow name="b"/></p:custShowLst>')),
    2,
  ],
  ['text.rtl', mk(slide('<a:pPr rtl="1"/><a:pPr rtl="0"/><a:pPr algn="r" rtl="true"/>')), 2],
  ['text.cjk', mk(slide('<a:t>日本語 abc</a:t>')), 3],
  [
    'text.lang',
    mk(slide('<a:rPr lang="he-IL"/><a:rPr lang="en-US"/><a:rPr dirty="0" lang="ja-JP"/><a:rPr/>')),
    2,
  ],
  [
    'objects.wordArt',
    mk(slide('<a:prstTxWarp prst="textArchUp"/><a:prstTxWarp prst="textNoShape"/>')),
    1,
  ],
  ['objects.ink', names('ppt/ink/ink1.xml', 'ppt/ink/_rels/ink1.xml.rels'), 1],
  ['objects.ole', mk(slide('<p:oleObj name="Worksheet"/>')), 1],
  ['objects.equations', mk(slide('<a14:m><m:oMathPara/></a14:m>')), 1],
  [
    'objects.actions',
    mk(
      slide(
        '<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=nextslide"/><a:hlinkClick r:id="rId1"/><a:hlinkClick action="ppaction://macro?name=M"/>'
      )
    ),
    2,
  ],
  [
    'objects.macroActions',
    mk(
      slide(
        '<a:hlinkClick action="ppaction://macro?name=M"/><a:hlinkClick action="ppaction://hlinksldjump"/>'
      )
    ),
    1,
  ],
  ['macros.vba', names('ppt/vbaProject.bin'), 1],
  ['sizes.4x3', size('cx="9144000" cy="6858000" type="screen4x3"'), 1],
  ['sizes.16x9', size('type="x" cy="6858000" cx="12192000"'), 1],
  ['sizes.a4', size('cx="10692000" cy="7560000" type="A4"'), 1],
  ['sizes.custom', size('cx="10058400" cy="6400800" type="custom"'), 1],
  ['sizes.portrait', size('cx="6858000" cy="9144000"'), 1],
  [
    'edge.truncatedZip',
    mk({}, { kind: 'unreadable', raw: new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2]) }),
    1,
  ],
  ['edge.encrypted', ole('EncryptionInfo'), 1],
  ['edge.legacyBinary', ole('PowerPoint Document'), 1],
  [
    'edge.noPresentation',
    mk({ '[Content_Types].xml': '<Types/>', 'ppt/slides/slide1.xml': '<p:sld/>' }),
    1,
  ],
  [
    'edge.slideMissing',
    mk({
      ...pres(
        '<p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId r:id="rId3"/><p:sldId r:id="rId4"/><p:sldId id="1"/></p:sldIdLst>'
      ),
      ...rels(
        'ppt/_rels/presentation.xml.rels',
        rel('rId1', 'slide', 'slides/slide1.xml'),
        rel('rId2', 'slide', 'slides/slide2.xml'),
        rel('rId4', 'slide', 'http://example.com/', ' TargetMode="External"')
      ),
      ...slide('', 1),
    }),
    4,
  ],
  [
    'edge.brokenRels',
    mk({
      ...rels(
        'ppt/slides/_rels/slide1.xml.rels',
        rel('rId1', 'image', '../media/image1.png'),
        rel('rId2', 'image', '../media/missing.png'),
        rel('rId3', 'oleObject', '../../../outside.bin'),
        rel('rId4', 'hyperlink', 'https://example.com', ' TargetMode="External"')
      ),
      'ppt/media/image1.png': '',
    }),
    2,
  ],
  [
    'edge.invalidXml',
    mk({ 'ppt/slides/slide1.xml': '<p:sld><p:sp></p:sld>', 'ppt/slides/slide2.xml': '<p:sld/>' }),
    1,
  ],
  [
    'edge.manySlides',
    mk(pres(`<p:sldIdLst>${'<p:sldId id="1" r:id="rId1"/>'.repeat(1000)}</p:sldIdLst>`)),
    1,
  ],
  [
    'edge.deepGroups',
    mk({
      ...slide('<p:grpSp>'.repeat(20) + '</p:grpSp>'.repeat(20), 1),
      ...slide('<p:grpSp>'.repeat(19) + '</p:grpSp>'.repeat(19), 2),
    }),
    1,
  ],
  [
    'edge.relCycle',
    mk({
      ...rels('ppt/slides/_rels/slide1.xml.rels', rel('rId1', 'slide', 'slide2.xml')),
      ...rels(
        'ppt/slides/_rels/slide2.xml.rels',
        rel('rId1', 'slide', 'slide1.xml'),
        rel('rId2', 'slide', 'slide2.xml')
      ),
    }),
    2,
  ],
];

describe('PPTX construct probes', () => {
  it('has one positive sample per registry entry and unique ids and labels', () => {
    expect(POSITIVE.map(([id]) => id).sort()).toEqual(CONSTRUCTS.map((c) => c.id).sort());
    expect(new Set(CONSTRUCTS.map((c) => c.id)).size).toBe(CONSTRUCTS.length);
    expect(new Set(CONSTRUCTS.map((c) => c.label)).size).toBe(CONSTRUCTS.length);
    for (const c of CONSTRUCTS) expect(c.id).toMatch(/^[a-z]+(?:\.[A-Za-z0-9]+)*$/);
  });

  it.each(POSITIVE)('%s counts its construct', (id, parts, expected) => {
    expect(findConstruct(id)?.count(parts)).toBe(expected);
  });

  it.each(CONSTRUCTS.map((c) => [c.id, c] as const))(
    '%s counts nothing in an empty package',
    (_id, c) => {
      expect(c.count(mk())).toBe(0);
      expect(c.count(mk({}, { kind: 'unreadable' }))).toBe(0);
      expect(
        c.count(mk({}, { oversized: ['ppt/slides/slide1.xml', 'ppt/presentation.xml'] }))
      ).toBe(0);
      expect(c.count(mk({}, { kind: 'ole' }))).toBe(0);
    }
  );

  it('does not mistake look-alikes for constructs', () => {
    const none = (id: string, p: OoxmlParts) =>
      expect({ id, n: findConstruct(id)?.count(p) }).toEqual({ id, n: 0 });
    none('slides.multiple', mk(slide('')));
    none('images.crop', mk(slide('<a:srcRect/><a:srcRect l="0" t="0"/>')));
    none('images.recolor', mk(slide('<a:lumMod val="1"/><a:lumOff val="1"/>')));
    none(
      'tables.cellFill',
      mk(slide('<a:tcPr/><a:tcPr><a:lnL/></a:tcPr><a:tcPr><a:noFill/></a:tcPr>'))
    );
    none('tables.customStyles', mk({ 'ppt/tableStyles.xml': '<a:tblStyleLst def="x"/>' }));
    none('groups.nested', mk(slide('<p:grpSp><p:grpSpPr/></p:grpSp>')));
    none('connectors.connected', cxn('<a:prstGeom prst="line"/>'));
    none(
      'charts.embeddedWorkbook',
      mk(
        rels(
          'ppt/charts/_rels/chart1.xml.rels',
          rel('rId1', 'themeOverride', '../theme/themeOverride1.xml')
        )
      )
    );
    none('transitions.extended', mk(slide('<mc:Choice Requires="p14"><p:fade/></mc:Choice>')));
    none('notes.rich', mk(notesPart('<a:rPr lang="x" b="0" u="none"/>')));
    none('slides.hidden', mk(slide('<p:sldId show="0"/>')));
    none('text.cjk', mk(slide('<a:t>ASCII only</a:t>')));
    none('animations.build', mk(slide('<p:bldP spid="4" build="whole"/>')));
    none('objects.actions', mk(slide('<a:hlinkClick r:id="rId1"/>')));
    none('masters.multiple', mk(master('')));
    none('masters.shapes', mk(master('<p:ph type="title"/>')));
    none('theme.multiple', mk({ ...masterTheme(1, 1), ...masterTheme(2, 1) }));
    none(
      'theme.multiple',
      mk(
        rels(
          'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
          rel('rId1', 'theme', '../theme/theme1.xml'),
          rel('rId2', 'theme', '../theme/theme2.xml')
        )
      )
    );
    none(
      'theme.colorMap',
      mk(clrMap('bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"', 1))
    );
    none(
      'theme.fontScheme',
      mk(theme('<a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>'))
    );
    none('theme.effectScheme', mk(theme('<a:effectStyle><a:effectLst/></a:effectStyle>')));
    none('text.hyperlinks', mk(slide('<a:rPr/><a:cNvPr><a:hlinkClick/></a:cNvPr>')));
    none('effects.text', mk(slide('<a:rPr lang="x"/><p:spPr><a:effectLst/></p:spPr>')));
    none(
      'fills.solid',
      mk(slide('<p:spPr/><p:ln><a:solidFill/></p:ln><p:spPr><a:ln><a:solidFill/></a:ln></p:spPr>'))
    );
    none('lines.dash', mk(slide('<a:prstDash val="dash"/>')));
    none(
      'lines.dash',
      mk(slide('<p:spPr><a:ln/></p:spPr><p:spPr><a:ln><a:prstDash val="solid"/></a:ln></p:spPr>'))
    );
    none('fills.picture', mk(slide('<p:blipFill><a:blip/></p:blipFill>')));
    none(
      'layouts.inheritance',
      mk(slide('<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm/></p:spPr>'))
    );
    none(
      'edge.noPresentation',
      mk({ 'ppt/presentation.xml': '<p:presentation/>', '[Content_Types].xml': '<Types/>' })
    );
    none('edge.noPresentation', mk({ 'ppt/slides/slide1.xml': '<p:sld/>' }));
    none('edge.noPresentation', mk({ '[Content_Types].xml': '<Types/>' }, { kind: 'ole' }));
    none(
      'edge.truncatedZip',
      mk({}, { kind: 'unreadable', raw: new Uint8Array([0x50, 0x4b, 0x01, 0x02]) })
    );
    none(
      'edge.truncatedZip',
      mk({}, { kind: 'zip', raw: new Uint8Array([0x50, 0x4b, 0x03, 0x04]) })
    );
    none(
      'edge.encrypted',
      mk({}, { kind: 'zip', raw: new Uint8Array(Buffer.from('EncryptionInfo', 'utf16le')) })
    );
    none('edge.legacyBinary', ole('Other'));
    none(
      'edge.manySlides',
      mk(pres(`<p:sldIdLst>${'<p:sldId id="1" r:id="rId1"/>'.repeat(999)}</p:sldIdLst>`))
    );
    none('edge.brokenRels', mk({ 'ppt/_rels/presentation.xml.rels': '<Relationships/>' }));
    none(
      'edge.relCycle',
      mk({
        ...rels(
          'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
          rel('rId1', 'slideMaster', '../slideMasters/slideMaster1.xml')
        ),
        ...rels(
          'ppt/slideMasters/_rels/slideMaster1.xml.rels',
          rel('rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml')
        ),
        ...rels(
          'ppt/notesSlides/_rels/notesSlide1.xml.rels',
          rel('rId1', 'slide', '../slides/slide1.xml')
        ),
        ...rels(
          'ppt/slides/_rels/slide1.xml.rels',
          rel('rId1', 'notesSlide', '../notesSlides/notesSlide1.xml')
        ),
      })
    );
  });

  it('reads attributes in any order', () => {
    const count = (id: string, p: OoxmlParts) => findConstruct(id)?.count(p);
    expect(count('text.columns', mk(slide('<a:bodyPr rtlCol="0" numCol="3" spcCol="1"/>')))).toBe(
      1
    );
    expect(count('text.runs', mk(slide('<a:rPr lang="en-US" dirty="0" b="1"/>')))).toBe(1);
    expect(
      count(
        'lines.arrowheads',
        mk(spPr('<a:ln><a:tailEnd w="lg" len="lg" type="stealth"/></a:ln>'))
      )
    ).toBe(1);
    expect(count('shapes.flip', mk(slide('<a:xfrm rot="0" flipV="1"/>')))).toBe(1);
    expect(
      count('layouts.footerPlaceholders', mk(slide('<p:ph idx="11" type="ftr" sz="quarter"/>')))
    ).toBe(1);
  });

  it('counts lines and fills of every shape separately, also next to self-closing property blocks', () => {
    const p = mk(
      slide(
        '<p:spPr/><p:spPr><a:solidFill/><a:ln w="1"><a:noFill/></a:ln></p:spPr><p:spPr><a:ln/></p:spPr><p:spPr><a:gradFill><a:lin/></a:gradFill><a:ln><a:solidFill/></a:ln></p:spPr>'
      )
    );
    expect(findConstruct('fills.solid')?.count(p)).toBe(1);
    expect(findConstruct('fills.gradient')?.count(p)).toBe(1);
    expect(findConstruct('fills.noFill')?.count(p)).toBe(0);
  });

  it('finds constructs by id', () => {
    expect(findConstruct('text.basic')?.label).toBe('Text runs');
    expect(findConstruct('nope')).toBeUndefined();
  });
});
