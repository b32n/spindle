/**
 * Counts for one loss id. `add(label, n)` adds `n` objects found at `label` (a slide, a font, a geometry name...);
 * `where` lists the distinct labels in the order they were first met, so a golden is stable.
 */
export class Tally {
  private readonly labels = new Map<string, number>();
  total = 0;
  add(label: string, n = 1): void {
    if (n <= 0) return;
    this.total += n;
    this.labels.set(label, (this.labels.get(label) ?? 0) + n);
  }
  get where(): string[] {
    return [...this.labels.keys()];
  }
}

export const TALLY_KEYS = [
  // summary ids
  'charts', 'smartArt', 'media', 'embedded', 'transitions', 'animations', 'comments', 'hidden', 'notes',
  'geometry', 'unsupportedImages',
  // full ids
  'layoutInheritance', 'colorMap', 'fontScheme', 'effectScheme',
  'autofit', 'bulletsCustom', 'bulletsPicture', 'tabStops', 'columns', 'vertical', 'fields', 'wordArt', 'rtl',
  'spacing', 'baseline', 'lang',
  'custGeom', 'effects', 'gradientFill', 'patternFill', 'elbowCurved', 'nestedGroups', 'scene3d',
  'crop', 'recolor', 'linked', 'svg',
  'merge', 'tableStyles', 'banding',
  'zoom', 'ink', 'equations', 'externalLinks', 'actions',
] as const;

export type Tallies = Record<(typeof TALLY_KEYS)[number], Tally>;

export const newTallies = (): Tallies => Object.fromEntries(TALLY_KEYS.map((k) => [k, new Tally()])) as Tallies;
