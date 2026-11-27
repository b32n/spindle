import { XMLValidator } from 'fast-xml-parser';
import JSZip from 'jszip';
import { relsPart, resolveTarget } from '../common/ooxml-package';
import { descendants, kid, kids, parseXml, walk, type XmlElement } from '../xlsx/xml';

// Structural lints of a written presentation (plan A4): the mistakes that make PowerPoint offer to repair a file and that a
// string-built package is prone to. They run in Jest over what the writer produced (and over hand-broken packages, to prove
// each one fires); they read the finished package, not the writer's inputs, so they also catch a bug in the packaging itself.
// The list is the one of the plan: unique shape ids per part, relationship ids and targets, content types, the child order of
// `p:spTree`, and the id ranges of the slide, master and layout lists.

const SHAPE_CHILDREN = new Set(['sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic', 'contentPart', 'AlternateContent']);
const DRAWN_PARTS = /^ppt\/(?:slides|slideLayouts|slideMasters|notesSlides|notesMasters)\/[^/]+\.xml$/;

/** The parts of the package and the content types they are declared with. */
async function contentTypes(zip: JSZip, problems: string[]): Promise<{ defaults: Set<string>; overrides: Set<string> } | undefined> {
  const text = await zip.file('[Content_Types].xml')?.async('string');
  const root = text === undefined || XMLValidator.validate(text) !== true ? null : parseXml(text);
  if (!root) {
    problems.push('[Content_Types].xml is missing or not well-formed');
    return undefined;
  }
  return {
    defaults: new Set(kids(root, 'Default').map((d) => (d.attrs.Extension ?? '').toLowerCase())),
    overrides: new Set(kids(root, 'Override').map((o) => (o.attrs.PartName ?? '').replace(/^\//, ''))),
  };
}

function lintShapeTree(part: string, root: XmlElement, problems: string[]): void {
  const ids = new Set<string>();
  for (const c of descendants(root, 'cNvPr')) {
    const id = c.attrs.id ?? '';
    if (!/^\d+$/.test(id)) problems.push(`${part}: a shape has no numeric id`);
    else if (ids.has(id)) problems.push(`${part}: shape id ${id} is used twice`);
    ids.add(id);
  }
  // each shape tree (the slide's and every group's): `nvGrpSpPr`, `grpSpPr`, the shapes, and an `extLst` last
  const trees: XmlElement[] = [];
  walk(root, (e) => {
    if (e.name === 'spTree' || e.name === 'grpSp') trees.push(e);
  });
  for (const tree of trees) {
    const order = tree.children.map((c) => c.name);
    const headerOk = order[0] === 'nvGrpSpPr';
    const second = order[1] === 'grpSpPr';
    if (!headerOk || !second) problems.push(`${part}: ${tree.name} must start with nvGrpSpPr and grpSpPr`);
    const body = order.slice(2);
    const ext = body.indexOf('extLst');
    if (ext !== -1 && ext !== body.length - 1) problems.push(`${part}: ${tree.name} has children after extLst`);
    for (const name of body) if (name !== 'extLst' && !SHAPE_CHILDREN.has(name)) problems.push(`${part}: ${tree.name} has an unexpected child ${name}`);
  }
}

/** Every attribute that refers to a relationship (`r:id`, `r:embed`, `r:link`...) must name one the part has. */
function lintRelationshipRefs(part: string, root: XmlElement, ids: Set<string>, problems: string[]): void {
  walk(root, (e) => {
    for (const [attr, id] of Object.entries(e.rattrs ?? {})) if (!ids.has(id)) problems.push(`${part}: ${e.name}@r:${attr} refers to ${id}, which the part's relationships do not have`);
  });
}

function lintIdList(part: string, root: XmlElement, list: string, item: string, min: number, max: number, taken: Set<string>, problems: string[]): void {
  for (const e of kids(kid(root, list), item)) {
    const id = e.attrs.id ?? '';
    const n = Number(id);
    if (!/^\d+$/.test(id) || n < min || n > max) problems.push(`${part}: ${item} id ${id} is outside ${min}..${max}`);
    else if (taken.has(id)) problems.push(`${part}: ${item} id ${id} is not unique`);
    taken.add(id);
  }
}

/**
 * The problems of a written package, as sentences; an empty list means every lint passed. `bytes` is a zip.
 */
export async function lintPptxPackage(bytes: Uint8Array | JSZip): Promise<string[]> {
  const zip = bytes instanceof JSZip ? bytes : await JSZip.loadAsync(bytes);
  const problems: string[] = [];
  const parts = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const types = await contentTypes(zip, problems);

  if (types) {
    for (const part of parts) {
      if (part === '[Content_Types].xml') continue;
      const ext = part.slice(part.lastIndexOf('.') + 1).toLowerCase();
      if (!types.overrides.has(part) && !types.defaults.has(ext)) problems.push(`${part}: no content type (no Override, and no Default for .${ext})`);
    }
    for (const part of types.overrides) if (!zip.file(part)) problems.push(`[Content_Types].xml declares ${part}, which the package does not have`);
  }

  const idsTaken = { slide: new Set<string>(), master: new Set<string>() };
  for (const part of parts) {
    if (!/\.(?:xml|rels)$/.test(part)) continue;
    const text = await (zip.file(part) as JSZip.JSZipObject).async('string');
    const root = XMLValidator.validate(text) === true ? parseXml(text) : null;
    if (!root) {
      problems.push(`${part}: not well-formed XML`);
      continue;
    }
    if (part.endsWith('.rels')) {
      const ids = new Set<string>();
      const owner = part === '_rels/.rels' ? '' : part.replace(/_rels\/([^/]+)\.rels$/, '$1');
      for (const r of kids(root, 'Relationship')) {
        const id = r.attrs.Id ?? '';
        if (ids.has(id)) problems.push(`${part}: relationship id ${id} is used twice`);
        ids.add(id);
        if (r.attrs.TargetMode !== 'External' && !zip.file(resolveTarget(owner, r.attrs.Target ?? ''))) problems.push(`${part}: ${id} targets ${r.attrs.Target}, which the package does not have`);
      }
      continue;
    }
    const relsText = await zip.file(relsPart(part))?.async('string');
    const rels = relsText === undefined ? null : parseXml(relsText);
    lintRelationshipRefs(part, root, new Set(kids(rels ?? undefined, 'Relationship').map((r) => r.attrs.Id ?? '')), problems);
    if (DRAWN_PARTS.test(part)) lintShapeTree(part, kid(root, 'cSld') ?? root, problems);
    if (root.name === 'presentation') {
      lintIdList(part, root, 'sldIdLst', 'sldId', 256, 2147483647, idsTaken.slide, problems);
      lintIdList(part, root, 'sldMasterIdLst', 'sldMasterId', 2147483648, 4294967295, idsTaken.master, problems);
    }
    if (root.name === 'sldMaster') lintIdList(part, root, 'sldLayoutIdLst', 'sldLayoutId', 2147483648, 4294967295, idsTaken.master, problems);
    // a slide has one layout, a layout one master, a master one theme
    const needs: Record<string, string> = { sld: 'slideLayout', sldLayout: 'slideMaster', sldMaster: 'theme' };
    const want = needs[root.name];
    if (want && kids(rels ?? undefined, 'Relationship').filter((r) => (r.attrs.Type ?? '').endsWith(`/${want}`)).length !== 1) problems.push(`${part}: needs exactly one ${want} relationship`);
  }
  return problems;
}
