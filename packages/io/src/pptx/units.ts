/** 1 px = 9525 EMU (96 DPI), so 12192000 EMU = 1280 px. */
export const EMU_PER_PX = 9525;

/** A number from an attribute, or `undefined` for anything that is not a finite number. */
export function num(v: string | undefined): number | undefined {
  if (v === undefined || v.trim() === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** EMU (an attribute value) to px, or `undefined` when it is not a number. */
export function emuToPx(v: string | undefined): number | undefined {
  const n = num(v);
  return n === undefined ? undefined : n / EMU_PER_PX;
}

/** A `ST_Percentage` attribute (`75000` is 75%) as a fraction (0.75). Also reads the `75%` spelling of the strict schema. */
export function pct(v: string | undefined): number | undefined {
  if (v === undefined) return undefined;
  const text = v.trim();
  if (text.endsWith('%')) {
    const n = num(text.slice(0, -1));
    return n === undefined ? undefined : n / 100;
  }
  const n = num(text);
  return n === undefined ? undefined : n / 100000;
}
