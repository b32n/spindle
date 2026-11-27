import type { OpaqueElement } from '../scene/types';

const KIND_LABEL: Record<OpaqueElement['kind'], string> = {
  chart: 'Chart',
  smartArt: 'SmartArt diagram',
  ole: 'Embedded object',
  media: 'Media clip',
  ink: 'Ink drawing',
};

/** What an audio or video clip is called, from the clip element the importer recorded (`videoFile`, `audioFile`, `quickTimeFile`). */
const CLIP_LABEL: Record<string, string> = { videoFile: 'Video', audioFile: 'Audio', quickTimeFile: 'Video' };

/** The placeholder's label: `Chart: Sales 2024 (kept, not editable)`; the title is left out when the source states none. */
export function opaqueLabel(el: Pick<OpaqueElement, 'kind' | 'title' | 'props'>): string {
  const kind = el.kind === 'media' ? (CLIP_LABEL[el.props?.clip ?? ''] ?? KIND_LABEL.media) : KIND_LABEL[el.kind];
  return `${kind}${el.title ? `: ${el.title}` : ''} (kept, not editable)`;
}
