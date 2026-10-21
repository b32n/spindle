// Regenerates the committed E2E fixtures in e2e/fixtures.
//   node scripts/testing/make-fixtures.mjs
//
// Office files are REAL LibreOffice output: the sources (flat ODF .fods/.fodt, written
// by this script so the diff is reviewable) are converted headless with
// `soffice --convert-to xlsx|docx`. Needs LibreOffice (`soffice` on PATH).
// CSV files and the test image are generated directly. See e2e/fixtures/README.md.
import { execFileSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, writeFileSync, copyFileSync, rmSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import os from 'node:os';
import path from 'node:path';

const outDir = path.resolve(import.meta.dirname, '..', '..', 'e2e', 'fixtures');
mkdirSync(path.join(outDir, 'src'), { recursive: true });
const write = (name, data) => writeFileSync(path.join(outDir, name), data);

// ---------------------------------------------------------------- PNG (no deps)
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
export function makePng(w, h, rgb) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4);
    len.writeUInt32BE(data.length);
    const td = Buffer.concat([Buffer.from(type), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(td));
    return Buffer.concat([len, td, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 2; // 8-bit RGB
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    const o = y * (w * 3 + 1);
    for (let x = 0; x < w; x++) {
      const p = o + 1 + x * 3;
      const [r, g, b] = rgb(x, y);
      raw[p] = r; raw[p + 1] = g; raw[p + 2] = b;
    }
  }
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0)),
  ]);
}
const logo = makePng(120, 60, (x, y) => (x < 60 ? [30, 90 + y, 200] : [230, 120, 20 + x]));

// ------------------------------------------------------------------- XLSX source
const fods = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
 xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
 xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
 <office:settings>
  <config:config-item-set config:name="ooo:view-settings">
   <config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>
    <config:config-item config:name="ViewId" config:type="string">view1</config:config-item>
    <config:config-item config:name="ActiveTable" config:type="string">Sales</config:config-item>
    <config:config-item-map-named config:name="Tables">
     <config:config-item-map-entry config:name="Sales">
      <config:config-item config:name="HorizontalSplitMode" config:type="short">2</config:config-item>
      <config:config-item config:name="VerticalSplitMode" config:type="short">2</config:config-item>
      <config:config-item config:name="HorizontalSplitPosition" config:type="int">1</config:config-item>
      <config:config-item config:name="VerticalSplitPosition" config:type="int">1</config:config-item>
      <config:config-item config:name="ActiveSplitRange" config:type="short">2</config:config-item>
      <config:config-item config:name="PositionLeft" config:type="int">0</config:config-item>
      <config:config-item config:name="PositionRight" config:type="int">1</config:config-item>
      <config:config-item config:name="PositionTop" config:type="int">0</config:config-item>
      <config:config-item config:name="PositionBottom" config:type="int">1</config:config-item>
     </config:config-item-map-entry>
    </config:config-item-map-named>
   </config:config-item-map-entry></config:config-item-map-indexed>
  </config:config-item-set>
 </office:settings>
 <office:automatic-styles>
  <number:currency-style style:name="N-USD"><number:text>$</number:text><number:number number:decimal-places="2" number:min-integer-digits="1" number:grouping="true"/></number:currency-style>
  <number:percentage-style style:name="N-PCT"><number:number number:decimal-places="1" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>
  <number:date-style style:name="N-DATE"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>
  <number:number-style style:name="N-KG"><number:number number:decimal-places="1" number:min-integer-digits="1" number:grouping="true"/><number:text> kg</number:text></number:number-style>
  <number:number-style style:name="N-ZERO"><number:number number:decimal-places="0" number:min-integer-digits="5"/></number:number-style>
  <style:style style:name="co-a" style:family="table-column"><style:table-column-properties style:column-width="4.2cm"/></style:style>
  <style:style style:name="co-b" style:family="table-column"><style:table-column-properties style:column-width="1.9cm"/></style:style>
  <style:style style:name="co-c" style:family="table-column"><style:table-column-properties style:column-width="3.4cm"/></style:style>
  <style:style style:name="co-d" style:family="table-column"><style:table-column-properties style:column-width="2.6cm"/></style:style>
  <style:style style:name="hdr" style:family="table-cell">
   <style:table-cell-properties fo:background-color="#1f4e79" fo:border="1.0pt solid #000000"/>
   <style:paragraph-properties fo:text-align="center"/>
   <style:text-properties fo:font-weight="bold" fo:color="#ffffff"/>
  </style:style>
  <style:style style:name="usd" style:family="table-cell" style:data-style-name="N-USD"/>
  <style:style style:name="usd-b" style:family="table-cell" style:data-style-name="N-USD"><style:text-properties fo:font-weight="bold"/><style:table-cell-properties fo:border-top="1.5pt solid #000000"/></style:style>
  <style:style style:name="pct" style:family="table-cell" style:data-style-name="N-PCT"/>
  <style:style style:name="date" style:family="table-cell" style:data-style-name="N-DATE"/>
  <style:style style:name="kg" style:family="table-cell" style:data-style-name="N-KG"/>
  <style:style style:name="zero" style:family="table-cell" style:data-style-name="N-ZERO"/>
  <style:style style:name="banner" style:family="table-cell">
   <style:table-cell-properties fo:background-color="#ffe699"/>
   <style:paragraph-properties fo:text-align="center"/>
   <style:text-properties fo:font-style="italic"/>
  </style:style>
  <style:style style:name="red" style:family="table-cell"><style:text-properties fo:color="#c00000" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
  <style:style style:name="right" style:family="table-cell"><style:paragraph-properties fo:text-align="end"/></style:style>
  <style:style style:name="bold" style:family="table-cell"><style:text-properties fo:font-weight="bold"/></style:style>
 </office:automatic-styles>
 <office:body><office:spreadsheet>
  <table:table table:name="Sales">
   <table:table-column table:style-name="co-a"/><table:table-column table:style-name="co-b"/>
   <table:table-column table:style-name="co-c"/><table:table-column table:style-name="co-c"/>
   <table:table-column table:style-name="co-d"/><table:table-column table:style-name="co-c"/><table:table-column table:style-name="co-c"/>
   <table:table-row>
    ${['Item', 'Qty', 'Price', 'Total', 'Share', 'Date', 'Weight'].map((h) => `<table:table-cell table:style-name="hdr" office:value-type="string"><text:p>${h}</text:p></table:table-cell>`).join('')}
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Widget</text:p></table:table-cell>
    <table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" office:value-type="currency" office:currency="USD" office:value="2.5"><text:p>$2.50</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" table:formula="of:=[.B2]*[.C2]" office:value-type="currency" office:currency="USD" office:value="25"><text:p>$25.00</text:p></table:table-cell>
    <table:table-cell table:style-name="pct" table:formula="of:=[.D2]/[.D$5]" office:value-type="percentage" office:value="0.25"><text:p>25.0%</text:p></table:table-cell>
    <table:table-cell table:style-name="date" office:value-type="date" office:date-value="2024-03-15"><text:p>2024-03-15</text:p></table:table-cell>
    <table:table-cell table:style-name="kg" office:value-type="float" office:value="1234.5"><text:p>1,234.5 kg</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Gadget</text:p></table:table-cell>
    <table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" office:value-type="currency" office:currency="USD" office:value="12.75"><text:p>$12.75</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" table:formula="of:=[.B3]*[.C3]" office:value-type="currency" office:currency="USD" office:value="51"><text:p>$51.00</text:p></table:table-cell>
    <table:table-cell table:style-name="pct" table:formula="of:=[.D3]/[.D$5]" office:value-type="percentage" office:value="0.51"><text:p>51.0%</text:p></table:table-cell>
    <table:table-cell table:style-name="date" office:value-type="date" office:date-value="2024-12-01"><text:p>2024-12-01</text:p></table:table-cell>
    <table:table-cell table:style-name="zero" office:value-type="float" office:value="42"><text:p>00042</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Gizmo</text:p></table:table-cell>
    <table:table-cell office:value-type="float" office:value="8"><text:p>8</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" office:value-type="currency" office:currency="USD" office:value="3"><text:p>$3.00</text:p></table:table-cell>
    <table:table-cell table:style-name="usd" table:formula="of:=[.B4]*[.C4]" office:value-type="currency" office:currency="USD" office:value="24"><text:p>$24.00</text:p></table:table-cell>
    <table:table-cell table:style-name="pct" table:formula="of:=[.D4]/[.D$5]" office:value-type="percentage" office:value="0.24"><text:p>24.0%</text:p></table:table-cell>
    <table:table-cell table:style-name="date" office:value-type="date" office:date-value="2025-01-31"><text:p>2025-01-31</text:p></table:table-cell>
    <table:table-cell table:style-name="kg" office:value-type="float" office:value="0.5"><text:p>0.5 kg</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell table:style-name="bold" office:value-type="string"><text:p>Total</text:p></table:table-cell>
    <table:table-cell table:style-name="bold" table:formula="of:=SUM([.B2:.B4])" office:value-type="float" office:value="22"><text:p>22</text:p></table:table-cell>
    <table:table-cell/>
    <table:table-cell table:style-name="usd-b" table:formula="of:=SUM([.D2:.D4])" office:value-type="currency" office:currency="USD" office:value="100"><text:p>$100.00</text:p></table:table-cell>
    <table:table-cell table:style-name="pct" table:formula="of:=SUM([.E2:.E4])" office:value-type="percentage" office:value="1"><text:p>100.0%</text:p></table:table-cell>
   </table:table-row>
   <table:table-row><table:table-cell/></table:table-row>
   <table:table-row>
    <table:table-cell table:style-name="banner" table:number-columns-spanned="4" table:number-rows-spanned="2" office:value-type="string"><text:p>Merged banner</text:p></table:table-cell>
    <table:covered-table-cell table:number-columns-repeated="3"/>
   </table:table-row>
   <table:table-row><table:covered-table-cell table:number-columns-repeated="4"/></table:table-row>
   <table:table-row>
    <table:table-cell table:style-name="red" office:value-type="string"><text:p>Red underlined note</text:p></table:table-cell>
    <table:table-cell/>
    <table:table-cell table:style-name="right" office:value-type="string"><text:p>right</text:p></table:table-cell>
   </table:table-row>
  </table:table>
  <table:table table:name="Costs">
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Sales total</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[$Sales.D5]" office:value-type="float" office:value="100"><text:p>100</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>20% cost</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[$Sales.D5]*0.2" office:value-type="float" office:value="20"><text:p>20</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Margin</text:p></table:table-cell>
    <table:table-cell table:formula="of:=[.B1]-[.B2]" office:value-type="float" office:value="80"><text:p>80</text:p></table:table-cell>
   </table:table-row>
  </table:table>
  <table:table table:name="Notes ünï">
   <table:table-row>
    <table:table-cell office:value-type="string"><text:p>Café — naïve 日本語</text:p></table:table-cell>
   </table:table-row>
  </table:table>
 </office:spreadsheet></office:body>
</office:document>
`;

// -------------------------------------------------------------------- DOCX source
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const fodt = `<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
 <office:meta><dc:title>Fixture Document</dc:title></office:meta>
 <office:font-face-decls/>
 <office:styles>
  <style:style style:name="Standard" style:family="paragraph"><style:text-properties style:font-name="Liberation Serif" fo:font-size="12pt"/></style:style>
  <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:next-style-name="Standard" style:default-outline-level="1"><style:text-properties fo:font-size="20pt" fo:font-weight="bold"/></style:style>
  <style:style style:name="Heading_20_2" style:display-name="Heading 2" style:family="paragraph" style:next-style-name="Standard" style:default-outline-level="2"><style:text-properties fo:font-size="16pt" fo:font-weight="bold"/></style:style>
  <style:style style:name="Heading_20_3" style:display-name="Heading 3" style:family="paragraph" style:next-style-name="Standard" style:default-outline-level="3"><style:text-properties fo:font-size="14pt" fo:font-weight="bold"/></style:style>
  <style:style style:name="Header" style:family="paragraph"/>
  <style:style style:name="Footer" style:family="paragraph"/>
  <style:style style:name="Internet_20_link" style:display-name="Internet link" style:family="text"><style:text-properties fo:color="#0000ff" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
  <text:list-style style:name="Bullets">
   ${[1, 2, 3].map((l) => `<text:list-level-style-bullet text:level="${l}" text:bullet-char="•"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${l * 1.27}cm" fo:text-indent="-0.635cm" fo:margin-left="${l * 1.27}cm"/></style:list-level-properties></text:list-level-style-bullet>`).join('')}
  </text:list-style>
  <text:list-style style:name="Numbers">
   ${[1, 2, 3].map((l) => `<text:list-level-style-number text:level="${l}" style:num-suffix="." style:num-format="${l === 2 ? 'a' : '1'}"><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="${l * 1.27}cm" fo:text-indent="-0.635cm" fo:margin-left="${l * 1.27}cm"/></style:list-level-properties></text:list-level-style-number>`).join('')}
  </text:list-style>
 </office:styles>
 <office:automatic-styles>
  <style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  <style:style style:name="I" style:family="text"><style:text-properties fo:font-style="italic"/></style:style>
  <style:style style:name="U" style:family="text"><style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
  <style:style style:name="BIU" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic" style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
  <style:style style:name="PB" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:break-before="page"/></style:style>
  <style:style style:name="Tbl" style:family="table"><style:table-properties style:width="14cm" table:align="left"/></style:style>
  <style:style style:name="Tbl.A" style:family="table-column"><style:table-column-properties style:column-width="5cm"/></style:style>
  <style:style style:name="Tbl.B" style:family="table-column"><style:table-column-properties style:column-width="4.5cm"/></style:style>
  <style:style style:name="Cell" style:family="table-cell"><style:table-cell-properties fo:padding="0.1cm" fo:border="0.5pt solid #000000"/></style:style>
  <style:style style:name="Img" style:family="graphic"><style:graphic-properties style:wrap="none" text:anchor-type="as-char" style:vertical-pos="top" style:vertical-rel="baseline"/></style:style>
  <style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21.59cm" fo:page-height="27.94cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="2.5cm" fo:margin-right="2.5cm"/><style:header-style/><style:footer-style/></style:page-layout>
 </office:automatic-styles>
 <office:master-styles>
  <style:master-page style:name="Standard" style:page-layout-name="pm1">
   <style:header><text:p text:style-name="Header">Fixture header text</text:p></style:header>
   <style:footer><text:p text:style-name="Footer">Fixture footer — page <text:page-number text:select-page="current"/></text:p></style:footer>
  </style:master-page>
 </office:master-styles>
 <office:body><office:text>
  <text:h text:style-name="Heading_20_1" text:outline-level="1">Fixture Title</text:h>
  <text:p text:style-name="Standard">Plain text, <text:span text:style-name="B">bold</text:span>, <text:span text:style-name="I">italic</text:span>, <text:span text:style-name="U">underline</text:span> and <text:span text:style-name="BIU">all three</text:span>.</text:p>
  <text:h text:style-name="Heading_20_2" text:outline-level="2">Lists</text:h>
  <text:list text:style-name="Bullets">
   <text:list-item><text:p text:style-name="Standard">Bullet one</text:p>
    <text:list><text:list-item><text:p text:style-name="Standard">Nested bullet</text:p></text:list-item></text:list>
   </text:list-item>
   <text:list-item><text:p text:style-name="Standard">Bullet two</text:p></text:list-item>
  </text:list>
  <text:list text:style-name="Numbers">
   <text:list-item><text:p text:style-name="Standard">Step one</text:p>
    <text:list><text:list-item><text:p text:style-name="Standard">Sub step</text:p></text:list-item></text:list>
   </text:list-item>
   <text:list-item><text:p text:style-name="Standard">Step two</text:p></text:list-item>
  </text:list>
  <text:h text:style-name="Heading_20_3" text:outline-level="3">Table, link, image</text:h>
  <table:table table:name="T1" table:style-name="Tbl">
   <table:table-column table:style-name="Tbl.A"/><table:table-column table:style-name="Tbl.B"/>
   <table:table-row>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard"><text:span text:style-name="B">Name</text:span></text:p></table:table-cell>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard"><text:span text:style-name="B">Value</text:span></text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard">Alpha</text:p></table:table-cell>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard">1</text:p></table:table-cell>
   </table:table-row>
   <table:table-row>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard">Beta</text:p></table:table-cell>
    <table:table-cell table:style-name="Cell" office:value-type="string"><text:p text:style-name="Standard">2</text:p></table:table-cell>
   </table:table-row>
  </table:table>
  <text:p text:style-name="Standard">Visit <text:a xlink:type="simple" xlink:href="https://example.com/spindle" text:style-name="Internet_20_link">the example site</text:a> for details.</text:p>
  <text:p text:style-name="Standard"><draw:frame draw:style-name="Img" draw:name="Logo" text:anchor-type="as-char" svg:width="4cm" svg:height="2cm"><draw:image><office:binary-data>${logo.toString('base64')}</office:binary-data></draw:image><svg:title>Fixture logo</svg:title></draw:frame></text:p>
  <text:p text:style-name="PB">Text after the page break.</text:p>
 </office:text></office:body>
</office:document>
`;

// -------------------------------------------------------------------- CSV files
const quoting = [
  'id,name,notes,amount',
  '1,"Smith, John","He said ""hello"" twice",1234.5',
  '2,"Multi\nline cell","Line one\nLine two",-7',
  '3,plain,,0.1',
  '4,"  padded  ","trailing comma,",007',
  '5,"=1+1 text","quote "" at end""",1e3',
].join('\n');
write('quoting.csv', quoting + '\n');
write('unicode.csv', [
  'language,greeting,notes',
  'French,Café crème,"à la carte, s\'il vous plaît"',
  'German,Grüße aus Köln,Straße',
  'Japanese,こんにちは世界,日本語のテキスト',
  'Chinese,你好，世界,简体中文',
  'Emoji,🚀 Launch 🎉,"family 👨‍👩‍👧‍👦, and flag 🇯🇵"',
  'Arabic,مرحبا بالعالم,نص من اليمين إلى اليسار',
  'Hebrew,שלום עולם,עברית',
].join('\r\n') + '\r\n');
write('semicolon.csv', 'name;qty;price\nWidget;10;2,5\nGadget;4;12,75\n'); // European Excel export
write('tabs.tsv.csv', 'name\tqty\tprice\nWidget\t10\t2.5\n'); // tab-delimited export saved as .csv
write('notes.txt', 'This is a plain text file, not a spreadsheet or a document.\n');
write('logo.png', logo);
write('logo.svg', `<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60" viewBox="0 0 120 60"><rect width="60" height="60" fill="#1e5ac8"/><rect x="60" width="60" height="60" fill="#e67814"/><text x="60" y="35" font-size="14" text-anchor="middle" fill="#fff">SVG</text></svg>\n`);

// --------------------------------------------------------------- LibreOffice step
write('src/styled-multisheet.fods', fods);
write('src/rich-document.fodt', fodt);
const tmp = mkdtempSync(path.join(os.tmpdir(), 'spindle-fixtures-'));
const profile = `-env:UserInstallation=file://${tmp}/profile`;
function convert(src, ext, filter, outName) {
  execFileSync('soffice', [profile, '--headless', '--convert-to', `${ext}:${filter}`, '--outdir', tmp, path.join(outDir, 'src', src)], { stdio: 'inherit' });
  copyFileSync(path.join(tmp, src.replace(/\.[^.]+$/, `.${ext}`)), path.join(outDir, outName));
}
convert('styled-multisheet.fods', 'xlsx', 'Calc MS Excel 2007 XML', 'styled-multisheet.xlsx');
convert('rich-document.fodt', 'docx', 'MS Word 2007 XML', 'rich-document.docx');
// Image variants of logo.png (JPEG / GIF / WebP) via Pillow; LibreOffice's graphic export would
// place the picture on a full page instead of keeping its 120x60 size.
execFileSync(
  'python3',
  ['-c', 'import sys\nfrom PIL import Image\nsrc, out = sys.argv[1], sys.argv[2]\nim = Image.open(src).convert("RGB")\nim.save(out + ".jpg", quality=90)\nim.save(out + ".gif")\nim.save(out + ".webp", lossless=True)', path.join(outDir, 'logo.png'), path.join(outDir, 'logo')],
  { stdio: 'inherit' }
);
rmSync(tmp, { recursive: true, force: true });

// Library-generated variant (exceljs): LibreOffice's headless conversion drops the
// frozen-pane view setting, so frozen panes + widths live in this file instead.
const { default: ExcelJS } = await import('exceljs');
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('Frozen', { views: [{ state: 'frozen', xSplit: 1, ySplit: 2 }] });
ws.columns = [{ width: 30 }, { width: 12 }, { width: 18 }];
ws.getRow(1).values = ['Frozen header', 'Qty', 'Amount'];
ws.getRow(1).font = { bold: true };
ws.getRow(2).values = ['Second frozen row', 1, 2];
for (let r = 3; r <= 40; r++) ws.getRow(r).values = [`Row ${r}`, r, r * 1.5];
ws.getCell('D3').value = { formula: 'SUM(B3:C3)', result: 4.5 };
ws.getCell('C3').numFmt = '0.00';
await wb.xlsx.writeFile(path.join(outDir, 'exceljs-frozen.xlsx'));
console.log('fixtures written to', outDir);
