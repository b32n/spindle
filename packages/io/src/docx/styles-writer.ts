import JSZip from 'jszip';
import type { DocNamedStyle, DocParagraphStyle, DocRunBase } from '@b32nio/spindle-shared/export/docx';

// ---------------------------------------------------------------------------
// Named styles -> `w:style` elements of word/styles.xml.
//
// The `docx` library writes its own built-in styles (Normal, Heading1..6, Hyperlink, ...) and cannot be told to
// replace them, so the document's named styles are written here and spliced into the library's styles part: a
// built-in with the same id is replaced by the document's definition, every other style is appended.
// ---------------------------------------------------------------------------

const escapeXml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const hex = (color: string): string => color.replace('#', '').toUpperCase();
const twips = (pt: number): number => Math.round(pt * 20);

const onOff = (tag: string, value: boolean | undefined): string => (value === undefined ? '' : value ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`);

/** Run properties in the order of the schema (`CT_RPr`). */
function rPr(text: DocRunBase | undefined): string {
  if (!text) return '';
  const fonts = [text.fontFamily, text.fontFamilyEastAsia, text.fontFamilyComplex].some((f) => f != null)
    ? `<w:rFonts${text.fontFamily != null ? ` w:ascii="${escapeXml(text.fontFamily)}" w:hAnsi="${escapeXml(text.fontFamily)}"` : ''}${
        (text.fontFamilyEastAsia ?? text.fontFamily) != null ? ` w:eastAsia="${escapeXml((text.fontFamilyEastAsia ?? text.fontFamily) as string)}"` : ''
      }${(text.fontFamilyComplex ?? text.fontFamily) != null ? ` w:cs="${escapeXml((text.fontFamilyComplex ?? text.fontFamily) as string)}"` : ''}/>`
    : '';
  const size = text.fontSize != null ? `<w:sz w:val="${Math.round(text.fontSize * 2)}"/><w:szCs w:val="${Math.round(text.fontSize * 2)}"/>` : '';
  const vert = text.superscript ? '<w:vertAlign w:val="superscript"/>' : text.subscript ? '<w:vertAlign w:val="subscript"/>' : '';
  const body =
    fonts +
    onOff('b', text.bold) +
    onOff('i', text.italic) +
    onOff('caps', text.allCaps) +
    onOff('smallCaps', text.smallCaps) +
    onOff('strike', text.strikethrough) +
    onOff('vanish', text.hidden) +
    (text.color ? `<w:color w:val="${hex(text.color)}"/>` : '') +
    size +
    (text.underline === undefined ? '' : text.underline ? '<w:u w:val="single"/>' : '<w:u w:val="none"/>') +
    (text.backgroundColor ? `<w:shd w:val="clear" w:color="auto" w:fill="${hex(text.backgroundColor)}"/>` : '') +
    vert;
  return body ? `<w:rPr>${body}</w:rPr>` : '';
}

const LINE_KEYWORD_TWIPS = { single: 240, onePointFive: 360, double: 480 } as const;
const JC = { left: 'left', center: 'center', right: 'right', justify: 'both' } as const;

/** Paragraph properties in the order of the schema (`CT_PPrBase`). */
function pPr(paragraph: DocParagraphStyle | undefined, outlineLevel: number | undefined): string {
  const p = paragraph ?? {};
  const custom = p.lineSpacingCustom;
  const line =
    custom != null
      ? custom.rule === 'multiple'
        ? ` w:line="${Math.round(custom.value * 240)}" w:lineRule="auto"`
        : ` w:line="${twips(custom.value)}" w:lineRule="${custom.rule === 'exactly' ? 'exact' : 'atLeast'}"`
      : p.lineSpacing
        ? ` w:line="${LINE_KEYWORD_TWIPS[p.lineSpacing]}" w:lineRule="auto"`
        : '';
  const spacing =
    p.spaceBeforePt != null || p.spaceAfterPt != null || line
      ? `<w:spacing${p.spaceBeforePt != null ? ` w:before="${twips(p.spaceBeforePt)}"` : ''}${p.spaceAfterPt != null ? ` w:after="${twips(p.spaceAfterPt)}"` : ''}${line}/>`
      : '';
  const first = p.firstLineIndentPt;
  const ind =
    p.leftIndentPt != null || p.rightIndentPt != null || first != null
      ? `<w:ind${p.leftIndentPt != null ? ` w:left="${twips(p.leftIndentPt)}"` : ''}${p.rightIndentPt != null ? ` w:right="${twips(p.rightIndentPt)}"` : ''}${
          first != null ? (first < 0 ? ` w:hanging="${twips(-first)}"` : ` w:firstLine="${twips(first)}"`) : ''
        }/>`
      : '';
  const body =
    (p.keepWithNext ? '<w:keepNext/>' : '') +
    (p.keepTogether ? '<w:keepLines/>' : '') +
    (p.pageBreakBefore ? '<w:pageBreakBefore/>' : '') +
    spacing +
    ind +
    (p.alignment ? `<w:jc w:val="${JC[p.alignment]}"/>` : '') +
    (outlineLevel !== undefined && outlineLevel >= 0 && outlineLevel <= 9 ? `<w:outlineLvl w:val="${outlineLevel === 0 ? 9 : Math.trunc(outlineLevel) - 1}"/>` : '');
  return body ? `<w:pPr>${body}</w:pPr>` : '';
}

const KINDS = new Set(['paragraph', 'character', 'table']);

/** A style id that is safe to write and whose definition is usable, as the document's named styles hold it. */
export function usableNamedStyles(styles: Record<string, DocNamedStyle> | undefined): Record<string, DocNamedStyle> {
  const out: Record<string, DocNamedStyle> = {};
  for (const [id, style] of Object.entries(styles ?? {})) {
    if (id && style && typeof style.name === 'string' && KINDS.has(style.kind)) out[id] = style;
  }
  return out;
}

function styleXml(id: string, style: DocNamedStyle, known: Set<string>): string {
  const ref = (tag: string, target: string | undefined) =>
    target && (tag === 'next' || target !== id) && known.has(target) ? `<w:${tag} w:val="${escapeXml(target)}"/>` : '';
  const props = style.kind === 'table' ? '' : pPr(style.kind === 'paragraph' ? style.paragraph : undefined, style.kind === 'paragraph' ? style.outlineLevel : undefined);
  return (
    `<w:style w:type="${style.kind}"${style.isDefault ? ' w:default="1"' : ''} w:styleId="${escapeXml(id)}">` +
    `<w:name w:val="${escapeXml(style.name || id)}"/>` +
    ref('basedOn', style.basedOn) +
    ref('next', style.kind === 'paragraph' ? style.next : undefined) +
    ref('link', style.kind === 'table' ? undefined : style.link) +
    (style.kind === 'table' ? '' : '<w:qFormat/>') +
    props +
    rPr(style.text) +
    '</w:style>'
  );
}

/** Replace/append the document's named styles in the `docx` output. Returns the bytes unchanged when there are none. */
export async function addNamedStyles(bytes: Uint8Array, styles: Record<string, DocNamedStyle>): Promise<Uint8Array> {
  const ids = Object.keys(styles);
  if (ids.length === 0) return bytes;
  const zip = await JSZip.loadAsync(bytes);
  const file = zip.file('word/styles.xml');
  if (!file) return bytes;
  let xml = await file.async('string');
  // A style may be based on, linked to or followed by any style the file will define: the document's own, the
  // library's built-ins, and the default character style Word implies.
  const known = new Set([...ids, 'DefaultParagraphFont', ...[...xml.matchAll(/<w:style\b[^>]*\bw:styleId="([^"]+)"/g)].map((m) => m[1])]);
  for (const id of ids) {
    const pattern = new RegExp(`<w:style\\b[^>]*\\bw:styleId="${id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*>[\\s\\S]*?</w:style>`, 'g');
    xml = xml.replace(pattern, '');
  }
  xml = xml.replace('</w:styles>', `${ids.map((id) => styleXml(id, styles[id], known)).join('')}</w:styles>`);
  zip.file('word/styles.xml', xml);
  return new Uint8Array(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));
}
