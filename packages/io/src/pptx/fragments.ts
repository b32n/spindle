// The slide-level XML the editor does not model (plan K11): the transition and the timing tree (animations, media playback),
// cut out of the slide's source text as written. The parsed tree drops namespace declarations, ordering details and attributes
// nobody read, so the fragments are taken from the text: byte for byte, with the namespaces the slide root declares added to the
// fragment's own tag so that it stands alone.

import type { PptxSlideOpaque } from '@b32nio/spindle-shared/export/pptx';

/** Comments, CDATA, processing instructions and DOCTYPE are skipped; a tag is a name and its attribute text (quotes may hold `>`). */
const TOKEN = /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<![^>]*>|<(\/?)([^\s/>]+)((?:"[^"]*"|'[^']*'|[^>"'])*)>/g;
const XMLNS = /\sxmlns(?::[\w.-]+)?\s*=\s*(?:"[^"]*"|'[^']*')/g;
const TRANSITION = /<(?:[\w.-]+:)?transition[\s>/]/;

const local = (qname: string): string => qname.slice(qname.indexOf(':') + 1);

interface Child {
  name: string;
  start: number;
  end: number;
}

/** The root's namespace declarations and the spans of its direct children (the text is well formed: the caller checked). */
function topLevel(text: string): { declarations: string[]; children: Child[] } {
  const children: Child[] = [];
  let declarations: string[] = [];
  let depth = 0;
  let open = { name: '', start: 0 };
  for (const m of text.matchAll(TOKEN)) {
    const [whole, closing, qname, attrs] = m;
    if (qname === undefined) continue;
    const end = m.index + whole.length;
    if (closing) {
      depth--;
      if (depth === 1) children.push({ ...open, end });
      continue;
    }
    if (depth === 0) declarations = attrs.match(XMLNS) ?? [];
    if (depth === 1) {
      open = { name: local(qname), start: m.index };
      if (attrs.endsWith('/')) children.push({ ...open, end });
    }
    if (!attrs.endsWith('/')) depth++;
  }
  return { declarations, children };
}

/** `<p:timing>` becomes `<p:timing xmlns:p="..." ...>`: the declarations the tag does not already carry are added. */
function standalone(fragment: string, declarations: string[]): string {
  const tag = (/^<[^\s/>]+/.exec(fragment) as RegExpExecArray)[0];
  const head = (/^<(?:"[^"]*"|'[^']*'|[^>"'])*>/.exec(fragment) as RegExpExecArray)[0];
  const missing = declarations.filter((d) => {
    const prefix = /xmlns(:[\w.-]+)?/.exec(d)?.[0] as string;
    return !new RegExp(`\\s${prefix.replace(/\./g, '\\.')}\\s*=`).test(head);
  });
  return tag + missing.join('') + fragment.slice(tag.length);
}

/** The transition (`p:transition`, or the `mc:AlternateContent` wrapping a newer one with its fallback) and the timing of a slide's source text. */
export function slideFragments(text: string): PptxSlideOpaque {
  const { declarations, children } = topLevel(text);
  const out: PptxSlideOpaque = {};
  for (const c of children) {
    const fragment = text.slice(c.start, c.end);
    if (c.name === 'timing' && out.timing === undefined) out.timing = standalone(fragment, declarations);
    else if (out.transition === undefined && (c.name === 'transition' || (c.name === 'AlternateContent' && TRANSITION.test(fragment)))) out.transition = standalone(fragment, declarations);
  }
  return out;
}
