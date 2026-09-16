// Deck-level controls: the new-slide layout gallery, theme picker, and slide
// background color. Rendered at the start of the toolbar. Inline it uses portal
// Popovers; inside the ResponsiveToolbar overflow surface (useToolbarMenu) it
// switches to menu rows that open a modal for their options.

import React, { useRef, useState } from 'react';
import { Plus, Palette, ChevronDown, SquareDashed } from 'lucide-react';
import { BUILTIN_THEMES, getBuiltinTheme } from '@b32nio/spindle-slides-core';
import { useToolbarMenu, MenuRow } from '@b32nio/spindle-shared/react';
import { useDeck, useActiveSlideId, useTheme, useSlide } from '../hooks';
import { Popover } from './Popover';
import { LayoutThumb } from './LayoutThumb';
import { TB } from './toolbarUI';

const btn: React.CSSProperties = TB.dropdownButton;
const item: React.CSSProperties = { padding: '7px 10px', borderRadius: 5, cursor: 'pointer', fontSize: 13 };

const BG_SWATCHES = [
  '#ffffff', '#f8fafc', '#f1f5f9', '#e2e8f0', '#94a3b8', '#475569', '#1e293b', '#0f172a',
  '#fee2e2', '#fef3c7', '#dcfce7', '#dbeafe', '#ede9fe', '#fce7f3', '#f59e0b', '#3b82f6',
];

export function DeckControls(): React.ReactElement {
  const deck = useDeck();
  const activeSlideId = useActiveSlideId();
  const theme = useTheme();
  const slide = useSlide(activeSlideId);
  const menu = useToolbarMenu();
  const [open, setOpen] = useState<null | 'layout' | 'theme'>(null);
  const layoutRef = useRef<HTMLButtonElement>(null);
  const themeRef = useRef<HTMLButtonElement>(null);

  const layouts = deck.getLayouts();
  const slideSize = deck.getSlideSize();
  const bgHex = slide?.background?.kind === 'solid' && slide.background.color.kind === 'rgb' ? slide.background.color.hex : '#ffffff';

  const hover = (e: React.MouseEvent, on: boolean) => ((e.currentTarget as HTMLElement).style.background = on ? '#f1f4f8' : 'transparent');
  const addSlide = (layoutId: string) => { const s = deck.addSlide({ afterSlideId: activeSlideId, layoutId }); deck.setActiveSlide(s.id); };
  const setBg = (hex: string) => deck.setSlideBackground(activeSlideId, { kind: 'solid', color: { kind: 'rgb', hex } });

  const layoutGrid = (onPick: (id: string) => void) => (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }}>
      {layouts.map((l) => (
        <div key={l.id} onClick={() => onPick(l.id)} style={{ cursor: 'pointer', borderRadius: 6, padding: 4 }} onMouseEnter={(e) => hover(e, true)} onMouseLeave={(e) => hover(e, false)}>
          <div style={{ border: '1px solid #d5d9e0', borderRadius: 4, overflow: 'hidden', aspectRatio: `${slideSize.w} / ${slideSize.h}`, background: '#fff' }}>
            <LayoutThumb layout={l} size={slideSize} />
          </div>
          <div style={{ fontSize: 12, color: '#3e4c59', textAlign: 'center', marginTop: 4 }}>{l.name}</div>
        </div>
      ))}
    </div>
  );

  const themeList = (onPick: (name: string) => void) => (
    <div>
      {BUILTIN_THEMES.map((t) => (
        <div key={t.name} onClick={() => onPick(t.name)} style={{ ...item, display: 'flex', alignItems: 'center', gap: 10, padding: '9px 8px' }} onMouseEnter={(e) => hover(e, true)} onMouseLeave={(e) => hover(e, false)}>
          <span style={{ display: 'flex', gap: 3 }}>
            {(['accent1', 'accent2', 'accent3', 'accent4'] as const).map((s) => (
              <span key={s} style={{ width: 16, height: 16, borderRadius: 4, background: t.colors[s] }} />
            ))}
          </span>
          <span style={{ fontSize: 14 }}>{t.name}</span>
        </div>
      ))}
    </div>
  );

  const bgPanel = (onDone: () => void) => (
    <div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, 1fr)', gap: 8 }}>
        {BG_SWATCHES.map((c) => (
          <button key={c} title={c} onClick={() => { setBg(c); onDone(); }} style={{ width: '100%', aspectRatio: '1', borderRadius: 8, border: c === bgHex ? '2px solid #2d7ff9' : '1px solid #d5d9e0', background: c, cursor: 'pointer' }} />
        ))}
      </div>
      <label style={{ display: 'flex', alignItems: 'center', gap: 10, marginTop: 14, fontSize: 14, color: '#334155', cursor: 'pointer' }}>
        <input type="color" value={bgHex} onChange={(e) => setBg(e.target.value)} style={{ width: 40, height: 32, border: '1px solid #d5d9e0', borderRadius: 6, background: 'none', padding: 2, cursor: 'pointer' }} />
        Custom color
      </label>
    </div>
  );

  // ── Overflow-menu mode: rows + modal (anchored popovers don't work here). ──
  if (menu.inMenu) {
    return (
      <>
        <MenuRow icon={<Plus size={17} />} label="New slide" onClick={() => menu.openModal('New slide', layoutGrid((id) => { addSlide(id); menu.closeMenu(); }))} />
        <MenuRow icon={<Palette size={17} />} label="Theme" hint={theme.name} onClick={() => menu.openModal('Theme', themeList((name) => { deck.setTheme(getBuiltinTheme(name)); menu.closeMenu(); }))} />
        <MenuRow icon={<SquareDashed size={17} />} label="Slide background" hint={<span style={{ width: 16, height: 16, borderRadius: 4, border: '1px solid #d5d9e0', background: bgHex, display: 'inline-block' }} />} onClick={() => menu.openModal('Slide background', bgPanel(() => menu.closeMenu()))} />
      </>
    );
  }

  // ── Inline mode: dropdown buttons with portal Popovers. ──
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
      <button ref={layoutRef} style={btn} title="New slide" onClick={() => setOpen(open === 'layout' ? null : 'layout')}>
        <Plus size={15} /> Slide <ChevronDown size={13} />
      </button>
      {open === 'layout' && (
        <Popover anchor={layoutRef.current} onClose={() => setOpen(null)}>
          <div style={{ fontSize: 11, color: '#8a93a2', padding: '2px 6px 8px' }}>Add slide with layout</div>
          <div style={{ width: 300 }}>{layoutGrid((id) => { addSlide(id); setOpen(null); })}</div>
        </Popover>
      )}

      <button ref={themeRef} style={btn} title="Theme" onClick={() => setOpen(open === 'theme' ? null : 'theme')}>
        <Palette size={15} /> {theme.name} <ChevronDown size={13} />
      </button>
      {open === 'theme' && (
        <Popover anchor={themeRef.current} onClose={() => setOpen(null)}>
          <div style={{ minWidth: 200 }}>{themeList((name) => { deck.setTheme(getBuiltinTheme(name)); setOpen(null); })}</div>
        </Popover>
      )}

      <label style={{ ...btn, gap: 6 }} title="Slide background">
        Background
        <input type="color" value={bgHex} onChange={(e) => setBg(e.target.value)} style={{ width: 22, height: 22, border: 'none', background: 'none', padding: 0, cursor: 'pointer' }} />
      </label>
    </div>
  );
}
