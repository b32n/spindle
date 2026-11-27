import { EMU_PER_PX, emuToPx, num, pct } from './units';

describe('units', () => {
  it('reads finite numbers only', () => {
    expect(num('12')).toBe(12);
    expect(num(' 1.5 ')).toBe(1.5);
    for (const bad of [undefined, '', '  ', 'abc', 'NaN', 'Infinity']) expect(num(bad)).toBeUndefined();
  });

  it('converts EMU to px at 9525 per px', () => {
    expect(EMU_PER_PX).toBe(9525);
    expect(emuToPx('12192000')).toBe(1280);
    expect(emuToPx('6858000')).toBe(720);
    expect(emuToPx('x')).toBeUndefined();
    expect(emuToPx(undefined)).toBeUndefined();
  });

  it('reads percentages in 1000ths of a percent and the strict "%" spelling', () => {
    expect(pct('75000')).toBe(0.75);
    expect(pct('100000')).toBe(1);
    expect(pct('-5000')).toBe(-0.05);
    expect(pct('75%')).toBe(0.75);
    expect(pct('x%')).toBeUndefined();
    expect(pct('x')).toBeUndefined();
    expect(pct(undefined)).toBeUndefined();
  });
});
