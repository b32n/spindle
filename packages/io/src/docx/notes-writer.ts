import JSZip from 'jszip';
import type { DocInline } from '@b32nio/spindle-shared/export/docx';
import type { NoteKind } from './notes';

// ---------------------------------------------------------------------------
// Footnotes and endnotes on export.
//
// The `docx` library writes a footnotes part and an endnotes part with only the separator notes. The notes of the
// document are generated here and put into those parts (zip post-processing; no network, no new dependency): each
// note's text is written as a paragraph by the same writer as the body (so its formatting, links and pictures are
// written the same way), taken out of that scratch document and wrapped as a note, with the links and pictures it
// needs added to the part's own relationships.
// ---------------------------------------------------------------------------

/** A note collected while the body was written: its text and custom mark. */
export interface WrittenNote {
  content: DocInline[];
  customMark?: string;
}

export type WrittenNotes = Record<NoteKind, WrittenNote[]>;

/** Writes the given inline contents as one paragraph each into a scratch .docx and returns its bytes. */
export type ScratchWriter = (bodies: DocInline[][]) => Promise<Uint8Array>;

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const HYPERLINK_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const escapeXml = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** The note's own reference mark, a superscript; direct formatting, because the reference styles may not be in styles.xml. */
const markRun = (kind: NoteKind) => `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:${kind}Ref/></w:r>`;

/** The paragraphs of a scratch document's body, in order. A note has no tables, so paragraphs do not nest. */
function bodyParagraphs(documentXml: string): string[] {
  const body = documentXml.slice(documentXml.indexOf('<w:body>') + '<w:body>'.length, documentXml.lastIndexOf('</w:body>'));
  return body.match(/<w:p\b[^>]*\/>|<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
}

/** A note paragraph with the reference mark (or the custom mark) and the space Word puts after it at its start. */
function withMark(paragraph: string, kind: NoteKind, customMark: string | undefined): string {
  const lead = (customMark ? `<w:r><w:t xml:space="preserve">${escapeXml(customMark)}</w:t></w:r>` : markRun(kind)) + '<w:r><w:t xml:space="preserve"> </w:t></w:r>';
  if (/^<w:p\b[^>]*\/>$/.test(paragraph)) return paragraph.replace(/\/>$/, `>${lead}</w:p>`);
  const pPrEnd = paragraph.indexOf('</w:pPr>');
  if (paragraph.startsWith('<w:p>') || paragraph.startsWith('<w:p ')) {
    const afterOpen = paragraph.indexOf('>') + 1;
    const at = paragraph.slice(afterOpen).startsWith('<w:pPr>') && pPrEnd >= 0 ? pPrEnd + '</w:pPr>'.length : afterOpen;
    return paragraph.slice(0, at) + lead + paragraph.slice(at);
  }
  return paragraph;
}

/** The hyperlink and picture relationships of a scratch document, with the pictures copied into `zip` under a name of their own. */
async function takeRelationships(scratch: JSZip, zip: JSZip, kind: NoteKind): Promise<string> {
  const relsXml = (await scratch.file('word/_rels/document.xml.rels')?.async('string')) ?? '';
  const out: string[] = [];
  for (const rel of relsXml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
    const attr = (name: string) => rel.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
    const type = attr('Type');
    const id = attr('Id');
    const target = attr('Target');
    if (!id || target === undefined) continue;
    if (type === HYPERLINK_REL) {
      out.push(rel);
    } else if (type === IMAGE_REL) {
      const source = scratch.file(`word/${target}`);
      if (!source) continue;
      const name = `word/media/note-${kind}-${target.replace(/^.*\//, '')}`;
      zip.file(name, await source.async('uint8array'));
      out.push(`<Relationship Id="${id}" Type="${IMAGE_REL}" Target="${name.replace(/^word\//, '')}"/>`);
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}">${out.join('')}</Relationships>`;
}

/** Put the notes of the document into the notes parts, and tell the settings where their separators are. */
export async function addNoteParts(bytes: Uint8Array, notes: WrittenNotes, scratchWriter: ScratchWriter): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  for (const kind of ['footnote', 'endnote'] as const) {
    const list = notes[kind];
    if (list.length === 0) continue;
    const scratch = await JSZip.loadAsync(await scratchWriter(list.map((n) => n.content)));
    const paragraphs = bodyParagraphs((await scratch.file('word/document.xml')?.async('string')) ?? '');
    const part = zip.file(`word/${kind}s.xml`);
    const existing = part ? await part.async('string') : '';
    const head = existing.slice(0, existing.indexOf(`<w:${kind} `));
    if (!head.includes(`<w:${kind}s`)) continue;
    const separator = (type: string, id: number) =>
      `<w:${kind} w:type="${type}" w:id="${id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${type}/></w:r></w:p></w:${kind}>`;
    const written = list.map((note, i) => `<w:${kind} w:id="${i + 1}">${withMark(paragraphs[i] ?? '<w:p/>', kind, note.customMark)}</w:${kind}>`);
    zip.file(`word/${kind}s.xml`, `${head}${separator('separator', -1)}${separator('continuationSeparator', 0)}${written.join('')}</w:${kind}s>`);
    zip.file(`word/_rels/${kind}s.xml.rels`, await takeRelationships(scratch, zip, kind));
  }
  const settings = zip.file('word/settings.xml');
  if (settings) {
    let xml = await settings.async('string');
    const props = (['footnote', 'endnote'] as const)
      .filter((kind) => notes[kind].length > 0 && !xml.includes(`<w:${kind}Pr`))
      .map((kind) => `<w:${kind}Pr><w:${kind} w:id="-1"/><w:${kind} w:id="0"/></w:${kind}Pr>`)
      .join('');
    if (props) {
      // CT_Settings puts the note properties before the compatibility settings.
      xml = xml.includes('<w:compat>') ? xml.replace('<w:compat>', `${props}<w:compat>`) : xml.replace('</w:settings>', `${props}</w:settings>`);
      zip.file('word/settings.xml', xml);
    }
  }
  return zip.generateAsync({ type: 'uint8array' });
}
