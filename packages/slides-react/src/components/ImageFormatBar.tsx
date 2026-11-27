// ImageFormatBar — inline image controls, rendered in the main Toolbar row (like
// the other format bars) when a single image is selected: replace the source
// (file or URL) and flip horizontally / vertically. The model already carries
// src / naturalW / naturalH / flipH / flipV / crop. Crop edges are edited in a small
// popover (percent trimmed from each edge) so the toolbar itself gains one button.

import React, { useRef, useState } from 'react';
import { ImageUp, Link2, FlipHorizontal2, FlipVertical2, Crop } from 'lucide-react';
import { normalizeCrop, type ImageElement } from '@b32nio/spindle-slides-core';
import { useImageHooks, readImageFile, storeImageFile, resolveImageUrl } from '@b32nio/spindle-shared/react';
import { useDeck, useSelection, useElement } from '../hooks';
import { ToolbarButton, ToolbarDivider } from './toolbarUI';
import { Popover } from './Popover';
import { ImageCropFields } from './ImageCropFields';

const FITS: Array<{ v: 'fill' | 'contain' | 'cover'; label: string }> = [
  { v: 'fill', label: 'Stretch' },
  { v: 'contain', label: 'Fit' },
  { v: 'cover', label: 'Crop to fill' },
];

/** Load an image and resolve its natural pixel size (or reject). */
function loadImage(src: string): Promise<{ naturalW: number; naturalH: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ naturalW: img.naturalWidth, naturalH: img.naturalHeight });
    img.onerror = reject;
    img.src = src;
  });
}

export function ImageFormatBar(): React.ReactElement | null {
  const deck = useDeck();
  const imageHooks = useImageHooks();
  const selection = useSelection();
  const fileRef = useRef<HTMLInputElement>(null);
  const [cropAnchor, setCropAnchor] = useState<HTMLElement | null>(null);
  const [cropOpen, setCropOpen] = useState(false);
  const id = selection.elementIds.length === 1 ? selection.elementIds[0] : null;
  const el = useElement(id ?? '') as ImageElement | undefined;
  if (!id || !el || el.type !== 'image') return null;

  // `measureSrc` is the local data URL used to read the size; `src` is what is stored
  // (the same data URL, or the URL returned by the host's uploadImage).
  const replace = async (measureSrc: string, src: string = measureSrc) => {
    try {
      const { naturalW, naturalH } = await loadImage(measureSrc);
      deck.updateElement(id, { src, naturalW, naturalH });
    } catch {
      /* couldn't load — leave the current image untouched */
    }
  };

  const onFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    try {
      const dataUrl = await readImageFile(imageHooks, file);
      await replace(dataUrl, await storeImageFile(imageHooks, file, dataUrl));
    } catch (err) {
      imageHooks.onImageError?.(err as Error);
    }
  };

  const replaceByUrl = async () => {
    const url = window.prompt('Image URL', el.src.startsWith('data:') ? '' : el.src);
    if (!url || !url.trim()) return;
    try {
      await replace(await resolveImageUrl(imageHooks, url.trim()));
    } catch (err) {
      imageHooks.onImageError?.(err as Error);
    }
  };

  return (
    <>
      <ToolbarDivider />
      <ToolbarButton title="Replace image (upload)" onClick={() => fileRef.current?.click()}>
        <ImageUp size={15} />
      </ToolbarButton>
      {imageHooks.resolveImage && <ToolbarButton title="Replace with image URL" onClick={replaceByUrl}>
        <Link2 size={15} />
      </ToolbarButton>}
      <select
        title="Image fit"
        value={el.fit ?? 'fill'}
        onMouseDown={(e) => e.stopPropagation()}
        onChange={(e) => deck.updateElement(id, { fit: e.target.value as ImageElement['fit'] })}
        style={{ height: 30, border: '1px solid rgba(0,0,0,0.1)', borderRadius: 8, background: 'rgba(255,255,255,0.8)', color: '#334155', padding: '0 6px', fontSize: 13, fontFamily: '"Inter", sans-serif', cursor: 'pointer' }}
      >
        {FITS.map((f) => (
          <option key={f.v} value={f.v}>{f.label}</option>
        ))}
      </select>
      <span ref={setCropAnchor} style={{ display: 'inline-flex' }}>
        <ToolbarButton title="Crop image" active={!!normalizeCrop(el.crop)} onClick={() => setCropOpen((o) => !o)}>
          <Crop size={15} />
        </ToolbarButton>
      </span>
      {cropOpen && (
        <Popover anchor={cropAnchor} onClose={() => setCropOpen(false)}>
          <ImageCropFields crop={el.crop} onChange={(crop) => deck.updateElement(id, { crop })} />
        </Popover>
      )}
      <ToolbarButton title="Flip horizontal" active={!!el.flipH} onClick={() => deck.updateElement(id, { flipH: !el.flipH })}>
        <FlipHorizontal2 size={15} />
      </ToolbarButton>
      <ToolbarButton title="Flip vertical" active={!!el.flipV} onClick={() => deck.updateElement(id, { flipV: !el.flipV })}>
        <FlipVertical2 size={15} />
      </ToolbarButton>
      <input ref={fileRef} type="file" accept="image/*" onChange={onFile} style={{ display: 'none' }} />
    </>
  );
}
