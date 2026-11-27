import { readFileSync } from 'fs';
import { join } from 'path';
import { importPptx } from '@b32nio/spindle-io/pptx';
import { pruneTiming, timingTargets } from './timing';

const ids = (xml: string | undefined): string[] => (xml === undefined ? [] : [...timingTargets(xml)].sort());
const par = (id: number, spid: string, cls = 'entr'): string =>
  `<p:par><p:cTn id="${id}" presetID="10" presetClass="${cls}" nodeType="clickEffect"><p:childTnLst><p:set><p:cBhvr><p:cTn id="${id + 1}"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par>`;
const click = (id: number, ...effects: string[]): string => `<p:par><p:cTn id="${id}"><p:childTnLst>${effects.join('')}</p:childTnLst></p:cTn></p:par>`;
const timing = (main: string, bld = ''): string => `<p:timing xmlns:p="urn:p"><p:tnLst><p:par><p:cTn id="1" nodeType="tmRoot"><p:childTnLst><p:seq><p:cTn id="2" nodeType="mainSeq"><p:childTnLst>${main}</p:childTnLst></p:cTn><p:nextCondLst><p:cond evt="onNext"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>${bld}</p:timing>`;

describe('pruneTiming', () => {
  it('leaves the text alone when nothing refers to a removed shape', () => {
    const xml = timing(click(3, par(4, '7')), '<p:bldLst><p:bldP spid="7"/></p:bldLst>');
    expect(pruneTiming(xml, new Set(['9']))).toBe(xml);
    expect(pruneTiming(xml, new Set())).toBe(xml);
  });

  it('removes the effect of a removed shape, its click step and its build entry, and keeps the other shapes\' effects as written', () => {
    const keep = click(10, par(11, '8'));
    const xml = timing(click(3, par(4, '7')) + keep, '<p:bldLst><p:bldP spid="7" grpId="0"/><p:bldP spid="8"/></p:bldLst>');
    const out = pruneTiming(xml, new Set(['7'])) as string;
    expect(ids(out)).toEqual(['8']);
    expect(out).toContain(keep);
    expect(out).toContain('<p:bldLst><p:bldP spid="8"/></p:bldLst>');
    expect(out).toContain('<p:sldTgt/>');
  });

  it('removes one of two effects in a step and keeps the step', () => {
    const out = pruneTiming(timing(click(3, par(4, '7'), par(6, '8', 'exit'))), new Set(['8'])) as string;
    expect(ids(out)).toEqual(['7']);
    expect(out).toContain('id="3"');
  });

  it('removes the whole timing when the last effect goes, including a build list left empty', () => {
    expect(pruneTiming(timing(click(3, par(4, '7')), '<p:bldLst><p:bldP spid="7"/></p:bldLst>'), new Set(['7']))).toBeUndefined();
  });

  it('removes only the build entry of a shape no effect names', () => {
    const xml = timing(click(3, par(4, '7')), '<p:bldLst><p:bldP spid="7"/><p:bldGraphic spid="8"><p:bldAsOne/></p:bldGraphic></p:bldLst>');
    const out = pruneTiming(xml, new Set(['8'])) as string;
    expect(out).toContain('<p:bldP spid="7"/>');
    expect(out).not.toContain('bldGraphic');
  });

  it('removes an interactive sequence the shape triggers', () => {
    const trigger = `<p:par><p:cTn id="20" nodeType="interactiveSeq"><p:stCondLst><p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cond></p:stCondLst><p:childTnLst>${par(21, '9')}</p:childTnLst></p:cTn></p:par>`;
    const out = pruneTiming(timing(click(3, par(4, '8')) + trigger), new Set(['7'])) as string;
    expect(ids(out)).toEqual(['8']);
  });

  it('removes just the behaviour when its shape is not inside an effect', () => {
    const loose = `<p:par><p:cTn id="3"><p:childTnLst><p:cmd type="call" cmd="playFrom(0)"><p:cBhvr><p:cTn id="4"/><p:tgtEl><p:spTgt spid="7"/></p:tgtEl></p:cBhvr></p:cmd>${'<p:set><p:cBhvr><p:cTn id="5"/><p:tgtEl><p:spTgt spid="8"/></p:tgtEl></p:cBhvr></p:set>'}</p:childTnLst></p:cTn></p:par>`;
    const out = pruneTiming(timing(loose), new Set(['7'])) as string;
    expect(ids(out)).toEqual(['8']);
    expect(out).not.toContain('<p:cmd');
  });

  it('leaves a reference it cannot place (no effect, behaviour or trigger around it)', () => {
    const odd = '<p:timing><p:tnLst><p:par><p:spTgt spid="7"/></p:par></p:tnLst></p:timing>';
    expect(pruneTiming(odd, new Set(['7']))).toBe(odd);
  });

  it('keeps the text between tags, comments, CDATA and quoted > characters', () => {
    const xml = `<p:timing><!-- a --><p:tnLst><p:par><p:cTn id="1"><p:childTnLst>${click(3, par(4, '7'))}<p:note a=">"/>text<![CDATA[<x>]]><?pi?></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
    const two = xml.replace('</p:childTnLst></p:cTn></p:par></p:tnLst>', `${click(5, par(6, '8'))}</p:childTnLst></p:cTn></p:par></p:tnLst>`);
    const kept = pruneTiming(two, new Set(['7'])) as string;
    expect(kept).toContain('<!-- a -->');
    expect(kept).toContain('<p:note a=">"/>text<![CDATA[<x>]]><?pi?>');
  });

  it('does not touch a container without a child list, or XML that does not nest', () => {
    const noList = `<p:timing><p:tnLst><p:par><p:cTn id="1"/></p:par><p:par>${par(4, '7')}</p:par></p:tnLst></p:timing>`;
    expect(pruneTiming(noList, new Set(['7']))).toBe('<p:timing><p:tnLst><p:par><p:cTn id="1"/></p:par><p:par></p:par></p:tnLst></p:timing>');
    for (const bad of ['<p:timing><p:a></p:b></p:timing>', '<p:timing><p:a>', '</p:timing>']) {
      expect(pruneTiming(bad, new Set(['7']))).toBe(bad);
      expect(timingTargets(bad).size).toBe(0);
    }
    expect(timingTargets('plain text, no tags').size).toBe(0);
  });

  it('reads single-quoted ids and ignores a reference that names no shape', () => {
    const xml = timing(click(3, par(4, '7').replace('spid="7"', "spid='7'"), par(6, '8').replace(' spid="8"', '')), '<p:bldLst><p:bldP/></p:bldLst>');
    expect(ids(xml)).toEqual(['7']);
    const out = pruneTiming(xml, new Set(['7'])) as string;
    expect(ids(out)).toEqual([]);
    expect(out).toContain('<p:spTgt/>');
    expect(out).toContain('<p:bldP/>');
  });

  it('prunes the animations of a real deck', async () => {
    const deck = await importPptx(new Uint8Array(readFileSync(join(__dirname, '../../../../fixtures/pptx/transitions-animations.pptx'))));
    const five = deck.slides[4].opaque?.timing as string;
    expect(ids(five)).toEqual(['3', '4', '5', '6']);
    const without4 = pruneTiming(five, new Set(['4'])) as string;
    expect(ids(without4)).toEqual(['3', '5', '6']);
    expect(without4).not.toContain('spid="4"');
    expect(pruneTiming(without4, new Set(['3', '5', '6']))).toBeUndefined();
    // every effect of the text-range deck targets one shape
    expect(pruneTiming(deck.slides[5].opaque?.timing as string, new Set(['3']))).toBeUndefined();
  });
});
