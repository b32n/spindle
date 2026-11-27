import { kid, kids, type XmlElement } from './xml';

/** Number format codes behind the built-in ids (ECMA-376 18.8.30, en-US). Ids not listed are locale- or script-specific. */
const BUILTIN_NUMBER_FORMATS: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '"$"#,##0_);\\("$"#,##0\\)',
  6: '"$"#,##0_);[Red]\\("$"#,##0\\)',
  7: '"$"#,##0.00_);\\("$"#,##0.00\\)',
  8: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};

export function builtinNumberFormat(id: number): string | undefined {
  return BUILTIN_NUMBER_FORMATS[id];
}

export interface UsedStyleFacts {
  patternFill: number;
  gradientFill: number;
  diagonalBorder: number;
  /** Names of the named cell styles (other than Normal) the used formats are built on. */
  namedStyles: string[];
}

export interface StyleFacts {
  /** The distinct number formats the used cell formats carry (General excluded): the code, or a label when the id has none. */
  numberFormats(used: ReadonlySet<number>): Array<{ code?: string; label: string }>;
  usedFacts(used: ReadonlySet<number>): UsedStyleFacts;
}

const list = (root: XmlElement | undefined, group: string, item: string): XmlElement[] => kids(kid(root, group), item);

export function parseStyles(root: XmlElement | undefined): StyleFacts {
  const codes = new Map<number, string>();
  for (const n of list(root, 'numFmts', 'numFmt')) codes.set(Number(n.attrs.numFmtId), n.attrs.formatCode ?? '');
  const fills = list(root, 'fills', 'fill');
  const borders = list(root, 'borders', 'border');
  const xfs = list(root, 'cellXfs', 'xf');
  const named = list(root, 'cellStyles', 'cellStyle');

  const usedXfs = (used: ReadonlySet<number>) => [...used].filter((i) => i < xfs.length).map((i) => xfs[i]);

  return {
    numberFormats(used) {
      const seen = new Map<string, { code?: string; label: string }>();
      for (const xf of usedXfs(used)) {
        const id = Number(xf.attrs.numFmtId ?? 0);
        if (id === 0) continue;
        const code = codes.get(id) ?? builtinNumberFormat(id);
        const label = code ?? `built-in format ${id}`;
        if (!seen.has(label)) seen.set(label, { code, label });
      }
      return [...seen.values()];
    },
    usedFacts(used) {
      const out: UsedStyleFacts = { patternFill: 0, gradientFill: 0, diagonalBorder: 0, namedStyles: [] };
      for (const xf of usedXfs(used)) {
        const fill = fills[Number(xf.attrs.fillId)];
        const border = borders[Number(xf.attrs.borderId)];
        const pattern = kid(fill, 'patternFill');
        const gradient = kid(fill, 'gradientFill');
        const type = pattern?.attrs.patternType;
        if (type !== undefined && type !== 'none' && type !== 'solid' && type !== 'gray125') out.patternFill++;
        if (gradient) out.gradientFill++;
        if (kid(border, 'diagonal')?.attrs.style !== undefined) out.diagonalBorder++;
        const style = named.find((s) => s.attrs.xfId === xf.attrs.xfId && xf.attrs.xfId !== '0');
        if (style?.attrs.name && style.attrs.name !== 'Normal' && !out.namedStyles.includes(style.attrs.name)) out.namedStyles.push(style.attrs.name);
      }
      return out;
    },
  };
}
