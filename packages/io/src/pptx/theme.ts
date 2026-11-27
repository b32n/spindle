import type { PptxTheme, PptxThemeSlot } from '@b32nio/spindle-shared/export/pptx';
import { kid, type XmlElement } from '../xlsx/xml';
import { THEME_SLOTS, colorElement, parseColor } from './color';

/** Office's default theme (2013+), used for a slot or font the theme part does not define. */
export const FALLBACK_THEME: PptxTheme = {
  name: 'Office Theme',
  colors: {
    dk1: '#000000', lt1: '#FFFFFF', dk2: '#44546A', lt2: '#E7E6E6', accent1: '#4472C4', accent2: '#ED7D31', accent3: '#A5A5A5',
    accent4: '#FFC000', accent5: '#5B9BD5', accent6: '#70AD47', hlink: '#0563C1', folHlink: '#954F72',
  },
  fonts: { major: 'Calibri Light', minor: 'Calibri' },
};

const face = (font: XmlElement | undefined, script: string): string | undefined => {
  const t = (kid(font, script)?.attrs.typeface ?? '').trim();
  return t === '' ? undefined : t;
};

/**
 * `a:theme`: the 12 colours of `a:clrScheme` (a `srgbClr` or a `sysClr` with its last value) and the latin, east Asian and
 * complex families of `a:fontScheme`. A missing slot or font keeps the Office default; `warn` is told once per gap.
 */
export function parseTheme(root: XmlElement | undefined, warn: (message: string) => void = () => undefined): PptxTheme {
  const elements = kid(root, 'themeElements');
  const scheme = kid(elements, 'clrScheme');
  const colors = { ...FALLBACK_THEME.colors };
  const missing: string[] = [];
  for (const slot of THEME_SLOTS) {
    const c = parseColor(colorElement(kid(scheme, slot)));
    if (c?.kind === 'rgb') colors[slot as PptxThemeSlot] = c.hex;
    else missing.push(slot);
  }
  if (missing.length > 0) warn(`The theme does not define ${missing.join(', ')}; the Office default colours are used for them.`);
  const fonts = kid(elements, 'fontScheme');
  const major = kid(fonts, 'majorFont');
  const minor = kid(fonts, 'minorFont');
  const theme: PptxTheme = {
    name: (root?.attrs.name ?? '').trim() || FALLBACK_THEME.name,
    colors,
    fonts: { major: face(major, 'latin') ?? FALLBACK_THEME.fonts.major, minor: face(minor, 'latin') ?? FALLBACK_THEME.fonts.minor },
  };
  const extras: [keyof PptxTheme['fonts'], string | undefined][] = [
    ['majorEastAsian', face(major, 'ea')], ['minorEastAsian', face(minor, 'ea')],
    ['majorComplex', face(major, 'cs')], ['minorComplex', face(minor, 'cs')],
  ];
  for (const [key, value] of extras) if (value !== undefined) theme.fonts[key] = value;
  return theme;
}
