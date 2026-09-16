// PDF export via the browser print pipeline — CLIENT-owned, on purpose. Export
// (print/rasterization/file IO) is kept out of the @b32nio/spindle-*
// packages; apps wire it up themselves from the public rendering API. Renders
// every slide at native size into a hidden, slide-sized off-screen iframe with
// `@page { size: WxH; margin: 0 }` + a page break per slide, then print(); the
// user picks "Save as PDF".

import React from 'react';
import { createRoot } from 'react-dom/client';
import { DeckProvider, SlideView } from '@b32nio/spindle-slides-react';
import type { DeckImpl } from '@b32nio/spindle-slides-core';

function PrintDeck({ deck }: { deck: DeckImpl }): React.ReactElement {
  const ids = deck.getSlideIds();
  const { w, h } = deck.getSlideSize();
  return (
    <>
      {ids.map((id) => (
        <div key={id} className="print-page" style={{ width: w, height: h, overflow: 'hidden', position: 'relative' }}>
          <SlideView slideId={id} />
        </div>
      ))}
    </>
  );
}

export async function exportDeckToPdf(deck: DeckImpl): Promise<void> {
  const { w, h } = deck.getSlideSize();
  const iframe = document.createElement('iframe');
  iframe.setAttribute('aria-hidden', 'true');
  // A slide-sized off-screen frame gives the print engine a proper containing
  // block; a 1px frame makes Chrome ignore @page size and letterbox the slide.
  iframe.style.cssText = `position:fixed;left:-10000px;top:0;width:${w}px;height:${h}px;opacity:0;border:0;`;
  document.body.appendChild(iframe);

  const doc = iframe.contentDocument!;
  doc.open();
  doc.write(
    `<!doctype html><html><head><style>` +
      `@page { size: ${w}px ${h}px; margin: 0; }` +
      `html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact;}` +
      `.print-page{ width:${w}px; height:${h}px; overflow:hidden; page-break-after: always; break-after: page; -webkit-print-color-adjust:exact; print-color-adjust:exact; }` +
      `.print-page:last-child{ page-break-after: auto; break-after: auto; }` +
      `</style></head><body></body></html>`
  );
  doc.close();

  const root = createRoot(doc.body);
  await new Promise<void>((resolve) => {
    root.render(
      <DeckProvider deck={deck}>
        <PrintDeck deck={deck} />
      </DeckProvider>
    );
    setTimeout(resolve, 60);
  });

  try {
    await (doc as Document & { fonts?: FontFaceSet }).fonts?.ready;
  } catch {
    /* fonts API unavailable */
  }
  await Promise.all(Array.from(doc.images).map((img) => img.decode?.().catch(() => {})));

  iframe.contentWindow?.focus();
  iframe.contentWindow?.print();

  setTimeout(() => {
    root.unmount();
    iframe.remove();
  }, 1500);
}
