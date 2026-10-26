import { createContext, useContext } from 'react';

/**
 * How the editors get images. Spindle never fetches an arbitrary image URL on
 * its own: without `resolveImage` the "insert by URL" paths are off and only
 * local files can be inserted (as bounded data URLs, or uploaded via the host).
 */
export interface ImageHooks {
  /**
   * Turns a URL the user typed into image bytes (a Blob) or a `data:image/…`
   * URL. The host decides what may be fetched (allow-list, proxy, auth…).
   * Providing it enables the "insert by URL" UI.
   */
  resolveImage?: (url: string) => Promise<Blob | string>;
  /**
   * Stores a local file and returns the URL to put in the document instead of
   * embedding the bytes as a data URL.
   */
  uploadImage?: (file: File) => Promise<string>;
  /** Called when inserting an image fails (wrong type, too large, hook rejected). Slides has no inline error UI. */
  onImageError?: (error: Error) => void;
  /** Largest image (bytes) accepted from a file or from `resolveImage`. Default 10 MB. */
  maxImageBytes?: number;
}

export const DEFAULT_MAX_IMAGE_BYTES = 10 * 1024 * 1024;

export const ImageHooksContext = createContext<ImageHooks>({});

/** The image hooks provided by the surrounding editor. */
export function useImageHooks(): ImageHooks {
  return useContext(ImageHooksContext);
}

/** A user-facing image failure (wrong type, too large, hook failed). */
export class ImageError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ImageError';
  }
}

function maxBytes(hooks: ImageHooks): number {
  return hooks.maxImageBytes ?? DEFAULT_MAX_IMAGE_BYTES;
}

function tooLarge(hooks: ImageHooks): ImageError {
  const max = maxBytes(hooks);
  const label = max >= 1024 * 1024 ? `${Math.round(max / (1024 * 1024))}MB` : `${Math.max(1, Math.round(max / 1024))}KB`;
  return new ImageError(`Image file is too large (max ${label})`);
}

export function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      if (typeof reader.result === 'string') resolve(reader.result);
      else reject(new ImageError('Failed to read image'));
    };
    reader.onerror = () => reject(new ImageError('Failed to read image'));
    reader.readAsDataURL(blob);
  });
}

/** Validate a local image file (type, size) and read it as a data URL. */
export async function readImageFile(hooks: ImageHooks, file: File): Promise<string> {
  if (!file.type.startsWith('image/')) throw new ImageError('Please select an image file');
  if (file.size > maxBytes(hooks)) throw tooLarge(hooks);
  return blobToDataUrl(file);
}

/**
 * The URL to store for a local image: the host's upload result when
 * `uploadImage` is set, otherwise the (already size-checked) data URL.
 */
export async function storeImageFile(hooks: ImageHooks, file: File, dataUrl: string): Promise<string> {
  if (!hooks.uploadImage) return dataUrl;
  let url: unknown;
  try {
    url = await hooks.uploadImage(file);
  } catch {
    throw new ImageError('Image upload failed');
  }
  if (typeof url !== 'string' || url.trim() === '') throw new ImageError('Image upload failed');
  return url.trim();
}

/** Fetch an image through the host's `resolveImage` and return it as a bounded data URL. */
export async function resolveImageUrl(hooks: ImageHooks, url: string): Promise<string> {
  if (!hooks.resolveImage) {
    throw new ImageError('Inserting images by URL is not enabled. Upload a file instead.');
  }
  let result: Blob | string;
  try {
    result = await hooks.resolveImage(url);
  } catch {
    throw new ImageError('Could not load image. Please check the URL.');
  }
  let dataUrl: string;
  if (typeof result === 'string') {
    if (!/^data:image\//i.test(result)) throw new ImageError('Could not load image. Please check the URL.');
    dataUrl = result;
    // base64 payload ≈ 3/4 of the text
    if ((result.length * 3) / 4 > maxBytes(hooks)) throw tooLarge(hooks);
  } else {
    if (!result || !(result.type === '' || result.type.startsWith('image/'))) {
      throw new ImageError('Could not load image. Please check the URL.');
    }
    if (result.size > maxBytes(hooks)) throw tooLarge(hooks);
    dataUrl = await blobToDataUrl(result);
  }
  return dataUrl;
}
