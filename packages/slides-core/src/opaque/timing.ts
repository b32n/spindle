// Pruning of a slide's timing XML (`Slide.opaque.timing`, plan K11). The timing tree names shapes by their PPTX id (`spTgt
// spid`, and the build list `bldP spid` ...): deleting a shape must remove the effects that animate it, or the file would
// point at a shape that is not there (PowerPoint reports such a file as damaged). The XML is not modelled otherwise, so it is
// edited as a small tree of tags and text that serialises back to exactly what it was where nothing is removed.

const MAX_DEPTH = 256;
const SKIPPED: ReadonlyArray<readonly [string, string]> = [
  ['<!--', '-->'],
  ['<![CDATA[', ']]>'],
  ['<?', '?>'],
  ['<!', '>'],
];

interface Token {
  index: number;
  whole: string;
  closing: boolean;
  qname: string;
  attrs: string;
}

/**
 * The tags of an XML text (comments, CDATA, processing instructions and DOCTYPE are text to this tree; a tag is a name and its
 * attribute text, whose quotes may hold `>`), or `undefined` when a construct is not terminated. One pass of `indexOf`, so a
 * peer's hostile text (this XML comes from the shared document) cannot make it super-linear.
 */
function tokenize(xml: string): Token[] | undefined {
  const out: Token[] = [];
  for (let i = xml.indexOf('<'); i >= 0; i = xml.indexOf('<', i)) {
    const skip = SKIPPED.find(([open]) => xml.startsWith(open, i));
    if (skip) {
      const end = xml.indexOf(skip[1], i + skip[0].length);
      if (end < 0) return undefined;
      i = end + skip[1].length;
      continue;
    }
    const closing = xml[i + 1] === '/';
    let at = i + (closing ? 2 : 1);
    const nameAt = at;
    while (at < xml.length && !/[\s/>]/.test(xml[at])) at++;
    if (at === nameAt) return undefined;
    const attrsAt = at;
    for (; at < xml.length && xml[at] !== '>'; at++) {
      if (xml[at] === '"' || xml[at] === "'") {
        at = xml.indexOf(xml[at], at + 1);
        if (at < 0) return undefined;
      }
    }
    if (at >= xml.length) return undefined;
    out.push({ index: i, whole: xml.slice(i, at + 1), closing, qname: xml.slice(nameAt, attrsAt), attrs: xml.slice(attrsAt, at) });
    i = at + 1;
  }
  return out;
}
const SPID = /\sspid\s*=\s*(?:"([^"]*)"|'([^']*)')/;
const PRESET_CLASS = /\spresetClass\s*=/;

interface Tag {
  tag: string;
  local: string;
  open: string;
  close: string;
  children: Array<Tag | string>;
  parent?: Tag;
}

const localName = (qname: string): string => qname.slice(qname.indexOf(':') + 1);
const isTag = (n: Tag | string): n is Tag => typeof n !== 'string';

/** The tree of an XML text, or `undefined` when its tags do not nest, nest too deep or are not terminated (the text is then left alone). */
function parse(xml: string): Tag | undefined {
  const root: Tag = { tag: '', local: '', open: '', close: '', children: [] };
  let at = root;
  let last = 0;
  let depth = 0;
  const tokens = tokenize(xml);
  if (!tokens) return undefined;
  for (const { index, whole, closing, qname, attrs } of tokens) {
    if (index > last) at.children.push(xml.slice(last, index));
    last = index + whole.length;
    if (closing) {
      if (at.tag !== qname) return undefined;
      at.close = whole;
      at = at.parent as Tag;
      depth--;
    } else if (attrs.endsWith('/')) at.children.push({ tag: qname, local: localName(qname), open: whole, close: '', children: [], parent: at });
    else {
      const t: Tag = { tag: qname, local: localName(qname), open: whole, close: '', children: [], parent: at };
      at.children.push(t);
      at = t;
      if (++depth > MAX_DEPTH) return undefined;
    }
  }
  if (at !== root) return undefined;
  if (last < xml.length) root.children.push(xml.slice(last));
  return root;
}

const serialize = (t: Tag): string => t.open + t.children.map((c) => (isTag(c) ? serialize(c) : c)).join('') + t.close;

const spidOf = (t: Tag): string | undefined => {
  const m = SPID.exec(t.open);
  return m ? (m[1] ?? m[2]) : undefined;
};

const BUILD = new Set(['bldP', 'bldGraphic', 'bldDgm', 'bldOleChart']);
const BEHAVIOR = new Set(['set', 'anim', 'animEffect', 'animMotion', 'animClr', 'animScale', 'animRot', 'cmd', 'audio', 'video']);
const CONTAINER = new Set(['par', 'seq', 'excl']);

const elementChildren = (t: Tag): Tag[] => t.children.filter(isTag);
const child = (t: Tag, local: string): Tag | undefined => elementChildren(t).find((c) => c.local === local);

/** Every tag of the tree, parents before children. */
function* walk(t: Tag): Generator<Tag> {
  for (const c of elementChildren(t)) {
    yield c;
    yield* walk(c);
  }
}

const ancestors = function* (t: Tag): Generator<Tag> {
  for (let a = t.parent; a?.parent; a = a.parent) yield a;
};

/** Whether a `par` is a preset effect; its `cTn` is looked up once per `par` (a lookup per hit is quadratic on a hostile sibling list). */
function isPreset(par: Tag, cache: Map<Tag, boolean>): boolean {
  let v = cache.get(par);
  if (v === undefined) {
    v = PRESET_CLASS.test(child(par, 'cTn')?.open ?? '');
    cache.set(par, v);
  }
  return v;
}

/** The tag to remove for a reference to a deleted shape: its build entry, or the effect (or interactive sequence, or behaviour) the reference belongs to. */
function doomed(ref: Tag, presets: Map<Tag, boolean>): Tag | undefined {
  if (BUILD.has(ref.local)) return ref;
  let inCondition = false;
  let behavior: Tag | undefined;
  for (const a of ancestors(ref)) {
    if (a.local === 'cond') inCondition = true;
    else if (inCondition && CONTAINER.has(a.local)) return a;
    else if (a.local === 'par' && isPreset(a, presets)) return a;
    else if (!behavior && BEHAVIOR.has(a.local)) behavior = a;
  }
  return behavior;
}

/** The shape ids (`pptxId`s) a timing XML refers to. */
export function timingTargets(xml: string): Set<string> {
  const out = new Set<string>();
  const root = parse(xml);
  if (root)
    for (const t of walk(root)) {
      const id = t.local === 'spTgt' || BUILD.has(t.local) ? spidOf(t) : undefined;
      if (id !== undefined) out.add(id);
    }
  return out;
}

/**
 * Remove what refers to the shapes `removed` (their `pptxId`s) from a timing XML: the effects that animate them, interactive
 * sequences they trigger, their build entries, and every container (and finally the whole timing) that is left empty by it.
 * Returns the pruned XML (the text itself, untouched, when nothing refers to a removed shape) or `undefined` when no animation
 * remains. XML that does not parse is returned as it is.
 */
export function pruneTiming(xml: string, removed: ReadonlySet<string>): string | undefined {
  const root = parse(xml);
  if (!root) return xml;
  const hits = [...walk(root)].filter((t) => (t.local === 'spTgt' || BUILD.has(t.local)) && removed.has(spidOf(t) ?? ''));
  if (hits.length === 0) return xml;
  const gone = new Set<Tag>();
  const presets = new Map<Tag, boolean>();
  for (const ref of hits) {
    const d = doomed(ref, presets);
    if (d) gone.add(d);
  }
  if (gone.size === 0) return xml;
  // removals are collected and applied in one pass per parent (a filter per removal is quadratic on a hostile sibling list)
  const dead = new Set<Tag>(gone);
  const live = (t: Tag): Tag[] => elementChildren(t).filter((c) => !dead.has(c));
  const touched = new Set<Tag>();
  for (const d of gone) for (const a of ancestors(d)) touched.add(a);
  // containers whose child list the removal emptied go too, innermost first (a removed container touches its own parents, so the emptiness cascades)
  const all = [...walk(root)].reverse();
  for (const t of all) {
    if (!touched.has(t) || dead.has(t)) continue;
    const list = CONTAINER.has(t.local) ? child(child(t, 'cTn') ?? t, 'childTnLst') : t.local === 'bldLst' || t.local === 'tnLst' ? t : undefined;
    if (list && live(list).length === 0) dead.add(t);
  }
  for (const t of [root, ...all]) if (!dead.has(t) && t.children.some((c) => isTag(c) && dead.has(c))) t.children = t.children.filter((c) => !(isTag(c) && dead.has(c)));
  const timing = elementChildren(root)[0];
  const remaining = timing && child(timing, 'tnLst');
  return remaining && elementChildren(remaining).length > 0 ? serialize(root) : undefined;
}
