import {
  MAX_GEOMETRY_PATHS,
  MAX_PATH_COMMANDS,
  arcEndPoint,
  normalizeCustomGeometry,
  parseCustGeom,
  serializeCustGeom,
  type GeomNode,
} from './custom-geometry';
import type { CustomGeometry } from './types';

const n = (name: string, attrs: Record<string, string> = {}, children: GeomNode[] = []): GeomNode => ({ name, attrs, children });
const pt = (x: string | number, y: string | number): GeomNode => n('pt', { x: String(x), y: String(y) });
const geom = (...paths: GeomNode[]): GeomNode => n('custGeom', {}, [n('avLst'), n('pathLst', {}, paths)]);
const BOX = { w: 1000, h: 500 };

describe('parseCustGeom', () => {
  it('reads every command kind, with angles converted to degrees', () => {
    const g = parseCustGeom(
      geom(
        n('path', { w: '100', h: '50' }, [
          n('moveTo', {}, [pt(0, 0)]),
          n('lnTo', {}, [pt(100, 0)]),
          n('cubicBezTo', {}, [pt(100, 10), pt(90, 20), pt(80, 30)]),
          n('quadBezTo', {}, [pt(40, 50), pt(0, 50)]),
          n('arcTo', { wR: '10', hR: '20', stAng: '5400000', swAng: '-10800000' }),
          n('close'),
        ]),
      ),
    );
    expect(g).toEqual({
      paths: [
        {
          w: 100,
          h: 50,
          commands: [
            { op: 'move', x: 0, y: 0 },
            { op: 'line', x: 100, y: 0 },
            { op: 'cubic', x1: 100, y1: 10, x2: 90, y2: 20, x: 80, y: 30 },
            { op: 'quad', x1: 40, y1: 50, x: 0, y: 50 },
            { op: 'arc', wR: 10, hR: 20, stAng: 90, swAng: -180 },
            { op: 'close' },
          ],
        },
      ],
    });
  });

  it('keeps several paths with fill modes, stroke and extrusion flags', () => {
    const g = parseCustGeom(
      geom(
        n('path', { w: '10', h: '10', fill: 'none', stroke: '0' }),
        n('path', { w: '10', h: '10', fill: 'lightenLess', stroke: 'false', extrusionOk: 'true' }),
        n('path', { w: '10', h: '10', fill: 'norm', stroke: '1', extrusionOk: '0' }),
        n('path', { w: '10', h: '10', fill: 'darken', stroke: 'true' }),
      ),
    )!;
    expect(g.paths.map((p) => [p.fill, p.stroke, p.extrusionOk])).toEqual([
      ['none', false, undefined],
      ['lightenLess', false, true],
      [undefined, undefined, undefined],
      ['darken', undefined, undefined],
    ]);
  });

  it('takes the box for a path without a size and resolves the built-in guides', () => {
    const g = parseCustGeom(
      geom(n('path', {}, [n('moveTo', {}, [pt('l', 't')]), n('lnTo', {}, [pt('r', 'b')]), n('lnTo', {}, [pt('hc', 'vc')]), n('lnTo', {}, [pt('wd2', 'hd2')])])),
      BOX,
    )!;
    expect(g.paths[0]).toMatchObject({ w: 1000, h: 500 });
    expect(g.paths[0].commands).toEqual([
      { op: 'move', x: 0, y: 0 },
      { op: 'line', x: 1000, y: 500 },
      { op: 'line', x: 500, y: 250 },
      { op: 'line', x: 500, y: 250 },
    ]);
  });

  it.each([
    ['w', 'h', 1000, 500],
    ['wd4', 'hd4', 250, 125],
    ['wd8', 'hd8', 125, 62.5],
    ['wd3', 'hd3', 1000 / 3, 500 / 3],
    ['wd6', 'hd6', 1000 / 6, 500 / 6],
  ])('resolves guides %s %s', (gx, gy, x, y) => {
    const g = parseCustGeom(geom(n('path', {}, [n('moveTo', {}, [pt(gx, gy)])])), BOX)!;
    expect(g.paths[0].commands[0]).toEqual({ op: 'move', x, y });
  });

  it('resolves guides in arcTo attributes', () => {
    const g = parseCustGeom(geom(n('path', {}, [n('arcTo', { wR: 'wd2', hR: 'hd2', stAng: '0', swAng: '21600000' })])), BOX)!;
    expect(g.paths[0].commands[0]).toEqual({ op: 'arc', wR: 500, hR: 250, stAng: 0, swAng: 360 });
  });

  it.each([
    ['no pathLst', n('custGeom'), undefined],
    ['an empty pathLst', geom(), undefined],
    ['a guide formula as a coordinate', geom(n('path', { w: '1', h: '1' }, [n('moveTo', {}, [pt('x1', 0)])])), undefined],
    ['a missing y', geom(n('path', { w: '1', h: '1' }, [n('moveTo', {}, [n('pt', { x: '1' })])])), undefined],
    ['a missing point', geom(n('path', { w: '1', h: '1' }, [n('moveTo')])), undefined],
    ['too many points', geom(n('path', { w: '1', h: '1' }, [n('lnTo', {}, [pt(0, 0), pt(1, 1)])])), undefined],
    ['a short cubic', geom(n('path', { w: '1', h: '1' }, [n('cubicBezTo', {}, [pt(0, 0)])])), undefined],
    ['a short quad', geom(n('path', { w: '1', h: '1' }, [n('quadBezTo', {}, [pt(0, 0)])])), undefined],
    ['an arc with a guide angle', geom(n('path', { w: '1', h: '1' }, [n('arcTo', { wR: '1', hR: '1', stAng: 'cd4', swAng: '0' })])), undefined],
    ['an arc with a missing radius', geom(n('path', { w: '1', h: '1' }, [n('arcTo', { hR: '1', stAng: '0', swAng: '0' })])), undefined],
    ['an unknown command', geom(n('path', { w: '1', h: '1' }, [n('wiggle')])), undefined],
    ['an unknown fill mode', geom(n('path', { w: '1', h: '1', fill: 'shiny' })), undefined],
    ['a path with no size and no box', geom(n('path')), undefined],
    ['a path with a zero size and no box', geom(n('path', { w: '0', h: '0' })), undefined],
    ['a zero-size box', geom(n('path')), { w: 0, h: 5 }],
    ['a zero-height box', geom(n('path', { w: '4' })), { w: 5, h: 0 }],
  ] as Array<[string, GeomNode, { w: number; h: number } | undefined]>)('returns undefined for %s', (_label, node, box) => {
    expect(parseCustGeom(node, box)).toBeUndefined();
  });

  it('returns undefined past the path and command caps', () => {
    const many = Array.from({ length: MAX_GEOMETRY_PATHS + 1 }, () => n('path', { w: '1', h: '1' }));
    expect(parseCustGeom(geom(...many))).toBeUndefined();
    const cmds = Array.from({ length: MAX_PATH_COMMANDS + 1 }, () => n('close'));
    expect(parseCustGeom(geom(n('path', { w: '1', h: '1' }, cmds)))).toBeUndefined();
    const ok = Array.from({ length: MAX_PATH_COMMANDS }, () => n('close'));
    expect(parseCustGeom(geom(n('path', { w: '1', h: '1' }, ok)))!.paths[0].commands).toHaveLength(MAX_PATH_COMMANDS);
  });
});

describe('serializeCustGeom', () => {
  const sample: CustomGeometry = {
    paths: [
      {
        w: 100,
        h: 50,
        commands: [
          { op: 'move', x: 0, y: 0 },
          { op: 'line', x: 100, y: 0 },
          { op: 'cubic', x1: 100, y1: 10, x2: 90, y2: 20, x: 80, y: 30 },
          { op: 'quad', x1: 40, y1: 50, x: 0, y: 50 },
          { op: 'arc', wR: 10, hR: 20, stAng: 90, swAng: -180 },
          { op: 'close' },
        ],
      },
      { w: 8, h: 8, fill: 'darkenLess', stroke: false, extrusionOk: true, commands: [{ op: 'move', x: 1, y: 2 }] },
      { w: 8, h: 8, fill: 'norm', commands: [] },
    ],
  };

  // the writer emits a tree the reader can consume: a tiny XML tokenizer for the exact subset it writes
  function toTree(xml: string): GeomNode {
    const root = n('#root');
    const stack = [root];
    for (const m of xml.matchAll(/<(\/?)a:(\w+)([^>]*?)(\/?)>/g)) {
      const [, closing, name, attrText, selfClose] = m;
      if (closing) {
        stack.pop();
        continue;
      }
      const attrs: Record<string, string> = {};
      for (const a of attrText.matchAll(/(\w+)="([^"]*)"/g)) attrs[a[1]] = a[2];
      const node = n(name, attrs);
      (stack[stack.length - 1].children as GeomNode[]).push(node);
      if (!selfClose) stack.push(node);
    }
    return root.children[0];
  }

  it('round-trips through parseCustGeom', () => {
    const xml = serializeCustGeom(sample);
    const back = parseCustGeom(toTree(xml));
    // an empty `norm` path survives; `norm` itself is the omitted default
    expect(back).toEqual({
      paths: [sample.paths[0], sample.paths[1], { w: 8, h: 8, commands: [] }],
    });
  });

  it('writes the fixed frame, integer coordinates and degrees as 60000ths', () => {
    const xml = serializeCustGeom({
      paths: [{ w: 10.4, h: 9.6, commands: [{ op: 'move', x: 1.5, y: 2.4 }, { op: 'arc', wR: 3, hR: 4, stAng: 22.5, swAng: 45 }] }],
    });
    expect(xml.startsWith('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>')).toBe(true);
    expect(xml).toContain('<a:path w="10" h="10">');
    expect(xml).toContain('<a:moveTo><a:pt x="2" y="2"/></a:moveTo>');
    expect(xml).toContain('<a:arcTo wR="3" hR="4" stAng="1350000" swAng="2700000"/>');
    expect(xml.endsWith('</a:pathLst></a:custGeom>')).toBe(true);
  });

  it('writes the fill, stroke and extrusion attributes', () => {
    const xml = serializeCustGeom(sample);
    expect(xml).toContain('<a:path w="8" h="8" fill="darkenLess" stroke="0" extrusionOk="1">');
    expect(xml).toContain('<a:path w="8" h="8"></a:path>');
  });
});

describe('normalizeCustomGeometry', () => {
  const ok = { w: 10, h: 10, commands: [{ op: 'move', x: 0, y: 0 }, { op: 'close' }] };

  it('passes valid geometry through and fills in defaults', () => {
    expect(normalizeCustomGeometry({ paths: [ok] })).toEqual({ paths: [ok] });
    expect(
      normalizeCustomGeometry({ paths: [{ ...ok, fill: 'lighten', stroke: false, extrusionOk: true }] }),
    ).toEqual({ paths: [{ ...ok, fill: 'lighten', stroke: false, extrusionOk: true }] });
  });

  it('drops norm, true flags and unknown fill modes', () => {
    expect(normalizeCustomGeometry({ paths: [{ ...ok, fill: 'norm', stroke: true, extrusionOk: false }] })).toEqual({ paths: [ok] });
    expect(normalizeCustomGeometry({ paths: [{ ...ok, fill: 'sparkle' }] })).toEqual({ paths: [ok] });
    expect(normalizeCustomGeometry({ paths: [{ ...ok, fill: 3 }] })).toEqual({ paths: [ok] });
  });

  it('accepts every command kind and drops malformed commands', () => {
    const cmds = [
      { op: 'move', x: 1, y: 2 },
      { op: 'line', x: 1, y: 2 },
      { op: 'cubic', x1: 1, y1: 2, x2: 3, y2: 4, x: 5, y: 6 },
      { op: 'quad', x1: 1, y1: 2, x: 3, y: 4 },
      { op: 'arc', wR: 1, hR: 2, stAng: 3, swAng: 4 },
      { op: 'close', extra: 1 },
    ];
    const bad = [
      null, 5, [], { op: 'teleport' }, { op: 'move', x: 1 }, { op: 'line', x: NaN, y: 1 }, { op: 'cubic', x1: 1 },
      { op: 'quad', x1: 1, y1: 1, x: 'a', y: 1 }, { op: 'arc', wR: 1, hR: 1, stAng: 1 }, { x: 1, y: 1 },
    ];
    const g = normalizeCustomGeometry({ paths: [{ ...ok, commands: [...bad, ...cmds] }] })!;
    expect(g.paths[0].commands).toEqual([...cmds.slice(0, 5), { op: 'close' }]);
  });

  it.each([
    ['null', null],
    ['a string', 'x'],
    ['an array', []],
    ['no paths key', {}],
    ['paths not an array', { paths: {} }],
    ['an empty list', { paths: [] }],
    ['only bad paths', { paths: [null, { w: 0, h: 5, commands: [] }, { w: 5, h: -1, commands: [] }, { w: NaN, h: 1, commands: [] }, { w: 'a', h: 1, commands: [] }, { w: 1, h: 1 }, { w: 1, h: 1, commands: 'x' }] }],
  ])('returns undefined for %s', (_l, v) => {
    expect(normalizeCustomGeometry(v)).toBeUndefined();
  });

  it('caps the paths and the total commands', () => {
    const many = normalizeCustomGeometry({ paths: Array.from({ length: MAX_GEOMETRY_PATHS + 5 }, () => ok) })!;
    expect(many.paths).toHaveLength(MAX_GEOMETRY_PATHS);
    const big = { w: 1, h: 1, commands: Array.from({ length: MAX_PATH_COMMANDS - 1 }, () => ({ op: 'close' })) };
    const capped = normalizeCustomGeometry({ paths: [big, big, big] })!;
    expect(capped.paths.reduce((s, p) => s + p.commands.length, 0)).toBe(MAX_PATH_COMMANDS);
  });

  it('does not alias or carry unknown keys from its input', () => {
    const input = { paths: [{ ...ok, junk: true }], extra: 1 };
    const g = normalizeCustomGeometry(input)!;
    expect(g).toEqual({ paths: [ok] });
    expect(g.paths[0].commands).not.toBe(input.paths[0].commands);
  });
});

describe('arcEndPoint', () => {
  const close = (a: { x: number; y: number }, x: number, y: number): void => {
    expect(a.x).toBeCloseTo(x, 6);
    expect(a.y).toBeCloseTo(y, 6);
  };

  it('quarter of a circle from the top moves to the right', () => {
    // start at the top (270 degrees visually, y down); sweeping +90 clockwise ends at the right of the circle
    close(arcEndPoint(10, 0, 10, 10, 270, 90), 20, 10);
  });

  it('half a circle ends opposite the start', () => {
    close(arcEndPoint(0, 10, 10, 10, 180, 180), 20, 10);
    close(arcEndPoint(0, 10, 10, 10, 180, -180), 20, 10);
    close(arcEndPoint(0, 10, 10, 10, 180, 90), 10, 0);
    close(arcEndPoint(0, 10, 10, 10, 180, -90), 10, 20);
  });

  it('uses the visual angle on an ellipse', () => {
    // a 45 degree visual angle on a wide ellipse is the point where x = y offset from centre along the diagonal
    const e = arcEndPoint(0, 0, 20, 10, 45, 45);
    // centre is at start - (r cos t0, r sin t0) with t0 = atan2(20 sin45, 10 cos45) = atan(2)
    const t0 = Math.atan2(2, 1);
    const t1 = Math.atan2(20, 0);
    close(e, 20 * (Math.cos(t1) - Math.cos(t0)), 10 * (Math.sin(t1) - Math.sin(t0)));
  });

  it('a full turn returns to the start', () => {
    close(arcEndPoint(5, 5, 3, 7, 30, 360), 5, 5);
  });
});
