import { DEFAULT_COLOR_MAP, THEME_SLOTS, colorElement, parseColor, parseColorMap, toHex } from './color';
import { parseXml, type XmlElement } from '../xlsx/xml';

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const el = (inner: string): XmlElement => parseXml(`<a:solidFill ${A}>${inner}</a:solidFill>`) as XmlElement;
const color = (inner: string, map = DEFAULT_COLOR_MAP) => parseColor(colorElement(el(inner)), map);

describe('toHex', () => {
  it('normalises six hex digits and refuses anything else', () => {
    expect(toHex('4472c4')).toBe('#4472C4');
    expect(toHex(' 4472C4 ')).toBe('#4472C4');
    for (const bad of [undefined, '', '44', '4472C4FF', 'GGGGGG', '#4472C4']) expect(toHex(bad)).toBeUndefined();
  });
});

describe('colour map', () => {
  it('starts from the Office map and takes every valid slot from the master', () => {
    expect(parseColorMap(undefined)).toEqual(DEFAULT_COLOR_MAP);
    expect(DEFAULT_COLOR_MAP).toMatchObject({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' });
    const dark = parseColorMap(parseXml(`<p:clrMap ${A} bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent2" hlink="bogus"/>`) as XmlElement);
    expect(dark).toMatchObject({ bg1: 'dk1', tx1: 'lt1', bg2: 'dk2', tx2: 'lt2', accent1: 'accent2', hlink: 'hlink' });
  });

  it('lists the twelve theme slots', () => {
    expect(THEME_SLOTS).toHaveLength(12);
  });
});

describe('parseColor', () => {
  it('reads a literal, with alpha only when it is not opaque', () => {
    expect(color('<a:srgbClr val="ff0000"/>')).toEqual({ kind: 'rgb', hex: '#FF0000' });
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')).toEqual({ kind: 'rgb', hex: '#FF0000', alpha: 0.5 });
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="100000"/></a:srgbClr>')).toEqual({ kind: 'rgb', hex: '#FF0000' });
    expect(color('<a:srgbClr val="FF0000"><a:alpha val="-10"/></a:srgbClr>')).toEqual({ kind: 'rgb', hex: '#FF0000', alpha: 0 });
    expect(color('<a:srgbClr val="nope"/>')).toBeUndefined();
  });

  it('reads system colours by their last value, with the two window defaults', () => {
    expect(color('<a:sysClr val="windowText" lastClr="111111"/>')).toEqual({ kind: 'rgb', hex: '#111111' });
    expect(color('<a:sysClr val="windowText"/>')).toEqual({ kind: 'rgb', hex: '#000000' });
    expect(color('<a:sysClr val="window"/>')).toEqual({ kind: 'rgb', hex: '#FFFFFF' });
    expect(color('<a:sysClr val="other"/>')).toBeUndefined();
  });

  it('reads preset names, black for one it does not know', () => {
    expect(color('<a:prstClr val="red"/>')).toEqual({ kind: 'rgb', hex: '#FF0000' });
    expect(color('<a:prstClr/>')).toEqual({ kind: 'rgb', hex: '#000000' });
    expect(color('<a:prstClr val="mediumOrchid"/>')).toEqual({ kind: 'rgb', hex: '#000000' });
  });

  it('turns a scheme colour into a slot through the colour map, with luminance transforms', () => {
    expect(color('<a:schemeClr val="accent1"/>')).toEqual({ kind: 'theme', slot: 'accent1' });
    expect(color('<a:schemeClr val="bg1"/>')).toEqual({ kind: 'theme', slot: 'lt1' });
    expect(color('<a:schemeClr val="tx1"/>', parseColorMap(parseXml(`<p:clrMap ${A} tx1="lt1"/>`) as XmlElement))).toEqual({ kind: 'theme', slot: 'lt1' });
    expect(color('<a:schemeClr val="dk2"><a:lumMod val="75000"/><a:lumOff val="25000"/><a:alpha val="40000"/></a:schemeClr>')).toEqual({
      kind: 'theme',
      slot: 'dk2',
      lumMod: 0.75,
      lumOff: 0.25,
      alpha: 0.4,
    });
    expect(color('<a:schemeClr val="phClr"/>')).toBeUndefined();
    expect(color('<a:schemeClr/>')).toBeUndefined();
  });

  it('gives nothing for a colour model it does not read, a missing element or a parent with no colour', () => {
    expect(color('<a:scrgbClr r="1" g="2" b="3"/>')).toBeUndefined();
    expect(parseColor(undefined)).toBeUndefined();
    expect(colorElement(el('<a:gradFill/>'))).toBeUndefined();
    expect(colorElement(undefined)).toBeUndefined();
  });
});
