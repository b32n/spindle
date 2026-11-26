import { isKeptBlockControl, sdtContent } from './content-controls';
import { attrsOf, childrenOf, find, flattenText, tagOf, type XmlNode } from './xml-tree';

// ---------------------------------------------------------------------------
// The `TOC` field (a table of contents): where it starts and ends in a list of block elements, which of its runs are the
// field's own markup (begin, instruction, separator, end) and which are the entries it shows. The field starts in the first
// entry paragraph and ends after the last (in it, or in a paragraph of its own), with the entries' own fields (PAGEREF) inside.
// ---------------------------------------------------------------------------

export interface TocRange {
  /** Index of the paragraph the field starts in, and of the one it ends in (in the list given). */
  start: number;
  end: number;
  /** The instruction, e.g. `TOC \o "1-3" \h \z \u`. */
  instruction: string;
  /** The runs that hold the field's own begin, instruction, separator and end (not the entries' fields). */
  markers: Set<XmlNode>;
}

type FieldMark = { type: 'begin' | 'separate' | 'end' | 'instr'; run: XmlNode; text?: string };

/** The field characters and instruction text in an element, in document order, with the run each is in. */
function fieldMarks(node: XmlNode, run: XmlNode | undefined, out: FieldMark[]): void {
  for (const child of childrenOf(node)) {
    const tag = tagOf(child);
    if (!tag || tag === 'w:pPr' || tag === '#text') continue;
    if (tag === 'w:r') {
      fieldMarks(child, child, out);
    } else if (tag === 'w:fldChar' && run) {
      const type = attrsOf(child)['@_w:fldCharType'];
      if (type === 'begin' || type === 'separate' || type === 'end') out.push({ type, run });
    } else if (tag === 'w:instrText' && run) {
      out.push({ type: 'instr', run, text: flattenText(child) });
    } else {
      fieldMarks(child, run, out);
    }
  }
}

/** The field kind: the first word of the instruction, upper case. */
const kindOf = (instruction: string): string => instruction.trim().toUpperCase().split(/[\s\\"]/)[0];

/**
 * The `TOC` fields in a list of block elements (the children of the body or of a content control), outermost only. A field
 * is left out when something other than paragraphs lies between its start and end, or when a section ends inside it (the
 * table of contents could not be one block).
 */
export function findTocRanges(nodes: XmlNode[]): TocRange[] {
  const ranges: TocRange[] = [];
  let depth = 0;
  let candidate: { start: number; instruction: string; separated: boolean; markers: Set<XmlNode>; valid: boolean } | null = null;
  nodes.forEach((node, index) => {
    const tag = tagOf(node);
    if (tag !== 'w:p') {
      // Anything but a paragraph inside the field means it is not a table of contents we can keep as one block.
      if (candidate && tag !== 'w:bookmarkStart' && tag !== 'w:bookmarkEnd' && tag !== 'w:proofErr') candidate.valid = false;
      return;
    }
    const pPr = find(childrenOf(node), 'w:pPr');
    if (candidate && pPr && find(childrenOf(pPr), 'w:sectPr')) candidate.valid = false;
    const marks: FieldMark[] = [];
    fieldMarks(node, undefined, marks);
    for (const mark of marks) {
      if (mark.type === 'begin') {
        depth++;
        if (depth === 1) candidate = { start: index, instruction: '', separated: false, markers: new Set([mark.run]), valid: true };
      } else if (mark.type === 'instr') {
        if (candidate && depth === 1 && !candidate.separated) {
          candidate.instruction += mark.text ?? '';
          candidate.markers.add(mark.run);
        }
      } else if (mark.type === 'separate') {
        if (candidate && depth === 1) {
          candidate.separated = true;
          candidate.markers.add(mark.run);
        }
      } else if (depth > 0) {
        if (depth === 1 && candidate) {
          candidate.markers.add(mark.run);
          const instruction = candidate.instruction.replace(/\s+/g, ' ').trim();
          if (candidate.valid && candidate.separated && kindOf(instruction) === 'TOC') {
            ranges.push({ start: candidate.start, end: index, instruction, markers: candidate.markers });
          }
          candidate = null;
        }
        depth--;
      }
    }
  });
  return ranges;
}

/** A copy of an element without the given runs, wherever they are in it (the element itself when none is). */
export function withoutMarkers(node: XmlNode, markers: Set<XmlNode>): XmlNode {
  const tag = tagOf(node);
  if (!tag) return node;
  let changed = false;
  const kept = childrenOf(node).flatMap((child) => {
    if (markers.has(child)) {
      changed = true;
      return [];
    }
    const next = withoutMarkers(child, markers);
    if (next !== child) changed = true;
    return [next];
  });
  return changed ? ({ ...node, [tag]: kept } as XmlNode) : node;
}

/** Whether a paragraph shows nothing (no runs, links or fields), once the field's own runs are taken out. */
export function isEmptyParagraph(node: XmlNode): boolean {
  return childrenOf(node).every((child) => tagOf(child) === 'w:pPr');
}

/**
 * The lists of block elements a table of contents can be in, for a body: the body's own children, and the content of each
 * content control that is kept as a block (custom XML and the controls that are not kept are read as their content).
 */
export function tocLists(bodyChildren: XmlNode[]): XmlNode[][] {
  const lists: XmlNode[][] = [];
  const visit = (list: XmlNode[]): void => {
    const flat: XmlNode[] = [];
    const add = (nodes: XmlNode[]): void => {
      for (const node of nodes) {
        const tag = tagOf(node);
        if (tag === 'w:sdt' && isKeptBlockControl(node)) {
          flat.push(node);
          visit(sdtContent(node));
        } else if (tag === 'w:sdt') {
          add(sdtContent(node));
        } else if (tag === 'w:customXml') {
          add(childrenOf(node).filter((child) => tagOf(child) !== 'w:customXmlPr'));
        } else {
          flat.push(node);
        }
      }
    };
    add(list);
    lists.push(flat);
  };
  visit(bodyChildren);
  return lists;
}
