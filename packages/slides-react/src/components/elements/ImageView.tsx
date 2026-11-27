import React from 'react';
import { cropLayout, normalizeCrop, type ImageElement, type PctRect } from '@b32nio/spindle-slides-core';

const pct = (r: PctRect): React.CSSProperties => ({
  position: 'absolute',
  left: `${r.x}%`,
  top: `${r.y}%`,
  width: `${r.w}%`,
  height: `${r.h}%`,
});

export function ImageView({ el }: { el: ImageElement }): React.ReactElement {
  const flip = el.flipH || el.flipV ? `scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})` : undefined;
  const crop = normalizeCrop(el.crop);
  if (crop) {
    // A cropped picture: the box clips the region, the region clips the whole image positioned inside it.
    const { region, image } = cropLayout(crop, el.fit, el.naturalW, el.naturalH, el.w, el.h);
    return (
      <div
        data-image-crop={`${crop.l},${crop.t},${crop.r},${crop.b}`}
        style={{ position: 'relative', width: '100%', height: '100%', overflow: 'hidden', transform: flip, transformOrigin: 'center center' }}
      >
        <div data-image-crop-region style={{ ...pct(region), overflow: 'hidden' }}>
          <img
            src={el.src}
            alt=""
            draggable={false}
            style={{ ...pct(image), maxWidth: 'none', display: 'block', userSelect: 'none' }}
          />
        </div>
      </div>
    );
  }
  return (
    <img
      src={el.src}
      alt=""
      draggable={false}
      style={{
        width: '100%',
        height: '100%',
        objectFit: el.fit ?? 'fill',
        display: 'block',
        transform: flip,
        transformOrigin: 'center center',
        userSelect: 'none',
      }}
    />
  );
}
