import type { DocContentControl, DocContentControlItem, DocContentControlType } from '@b32nio/spindle-shared/export/docx';
import { attrsOf, childrenOf, find, findDeep, tagOf, walk, type XmlNode } from './xml-tree';

/**
 * Content controls (`w:sdt`). Block-level and run-level controls with one of the supported types are kept as a
 * wrapper around their content; the rest (repeating sections, data-bound, equation, citation, bibliography and group
 * controls, and controls wrapping table rows or cells) are read as their content only and reported by the detector.
 */

const UNSUPPORTED_PROPERTIES: Record<string, string> = {
  'w15:repeatingSection': 'repeating section',
  'w15:repeatingSectionItem': 'repeating section item',
  'w:dataBinding': 'data-bound control',
  'w:equation': 'equation control',
  'w:citation': 'citation control',
  'w:bibliography': 'bibliography control',
  'w:group': 'group control',
};

const isOff = (v: string | undefined) => v === '0' || v === 'false' || v === 'off';
const isTrue = (v: string | undefined) => v === '1' || v === 'true' || v === 'on';

export type ParsedSdt = { control: DocContentControl } | { unsupported: string };

/** The properties of a `w:sdt`, or why it cannot be kept as a control. */
export function parseSdt(sdt: XmlNode): ParsedSdt {
  const pr = find(childrenOf(sdt), 'w:sdtPr');
  const props = pr ? childrenOf(pr) : [];
  for (const prop of props) {
    const tag = tagOf(prop);
    if (tag && UNSUPPORTED_PROPERTIES[tag]) return { unsupported: UNSUPPORTED_PROPERTIES[tag] };
  }
  const val = (tag: string): string | undefined => {
    const el = find(props, tag);
    return el ? attrsOf(el)['@_w:val'] : undefined;
  };

  let type: DocContentControlType = 'richText';
  let items: DocContentControlItem[] | undefined;
  let dateFormat: string | undefined;
  let dateValue: string | undefined;
  let checked: boolean | undefined;
  let multiline: boolean | undefined;
  let gallery: string | undefined;

  const listItems = (el: XmlNode): DocContentControlItem[] =>
    childrenOf(el)
      .filter((k) => tagOf(k) === 'w:listItem')
      .map((k) => {
        const a = attrsOf(k);
        const text = a['@_w:displayText'] ?? a['@_w:value'] ?? '';
        return { text, value: a['@_w:value'] ?? text };
      });

  const text = find(props, 'w:text');
  const dropDown = find(props, 'w:dropDownList');
  const comboBox = find(props, 'w:comboBox');
  const date = find(props, 'w:date');
  const checkbox = find(props, 'w14:checkbox');
  const docPart = find(props, 'w:docPartObj') ?? find(props, 'w:docPartList');
  if (text) {
    type = 'text';
    if (isTrue(attrsOf(text)['@_w:multiLine'])) multiline = true;
  } else if (dropDown) {
    type = 'dropDown';
    items = listItems(dropDown);
  } else if (comboBox) {
    type = 'comboBox';
    items = listItems(comboBox);
  } else if (date) {
    type = 'date';
    const format = find(childrenOf(date), 'w:dateFormat');
    dateFormat = format ? attrsOf(format)['@_w:val'] : undefined;
    dateValue = attrsOf(date)['@_w:fullDate'];
  } else if (checkbox) {
    type = 'checkbox';
    const state = find(childrenOf(checkbox), 'w14:checked');
    checked = state ? !isOff(attrsOf(state)['@_w14:val']) : false;
  } else if (find(props, 'w:picture')) {
    type = 'picture';
  } else if (docPart) {
    type = 'docPart';
    const galleryEl = find(childrenOf(docPart), 'w:docPartGallery');
    gallery = galleryEl ? attrsOf(galleryEl)['@_w:val'] : undefined;
  }

  const lockValue = val('w:lock');
  const lock = lockValue === 'sdtLocked' || lockValue === 'contentLocked' || lockValue === 'sdtContentLocked' ? lockValue : undefined;
  const placeholderEl = find(props, 'w:placeholder');
  const placeholderDoc = placeholderEl ? find(childrenOf(placeholderEl), 'w:docPart') : undefined;
  const placeholder = placeholderDoc ? attrsOf(placeholderDoc)['@_w:val'] : undefined;
  const showing = find(props, 'w:showingPlcHdr');
  const alias = val('w:alias');
  const tag = val('w:tag');

  return {
    control: {
      type,
      ...(tag && { tag }),
      ...(alias && { alias }),
      ...(lock && { lock }),
      ...(placeholder && { placeholder }),
      ...(showing && !isOff(attrsOf(showing)['@_w:val']) && { showingPlaceholder: true }),
      ...(items && { items }),
      ...(dateFormat && { dateFormat }),
      ...(dateValue && { dateValue }),
      ...(checked !== undefined && { checked }),
      ...(multiline && { multiline }),
      ...(gallery && { gallery }),
    },
  };
}

/** The nodes a `w:sdt` wraps. */
export function sdtContent(sdt: XmlNode): XmlNode[] {
  const content = find(childrenOf(sdt), 'w:sdtContent');
  return content ? childrenOf(content) : [];
}

/**
 * Whether a block-level `w:sdt` (a child of the body or of a table cell) is kept as a wrapper. A section break inside
 * it would split the wrapper across sections, so such a control is read as its content only.
 */
export function isKeptBlockControl(sdt: XmlNode): boolean {
  return 'control' in parseSdt(sdt) && !findDeep(sdt, 'w:sectPr');
}

/**
 * Controls of a document part that are read as their content only, by reason (`row-level control`, `repeating
 * section`, ...). A control directly under a table (it wraps rows) or a row (it wraps cells) cannot be kept.
 */
export function unsupportedControls(root: XmlNode[]): Map<string, number> {
  const reasons = new Map<string, number>();
  walk(root, (node, tag, path) => {
    if (tag !== 'w:sdt') return;
    let container: string | undefined;
    for (let i = path.length - 1; i >= 0; i--) {
      if (path[i] !== 'w:sdtContent' && path[i] !== 'w:sdt') {
        container = path[i];
        break;
      }
    }
    let reason: string | undefined;
    if (container === 'w:tbl') reason = 'row-level control';
    else if (container === 'w:tr') reason = 'cell-level control';
    else {
      const parsed = parseSdt(node);
      if ('unsupported' in parsed) reason = parsed.unsupported;
      else if ((container === 'w:body' || container === 'w:tc') && findDeep(node, 'w:sectPr')) reason = 'control around a section break';
    }
    if (reason) reasons.set(reason, (reasons.get(reason) ?? 0) + 1);
  });
  return reasons;
}
