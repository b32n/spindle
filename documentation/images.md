# Images and privacy

Spindle never fetches an arbitrary image URL on its own. The docs and slides editors take three props that decide how images get in:

```tsx
<DocumentEditor
  resolveImage={async (url) => (await fetch('/api/image-proxy?u=' + encodeURIComponent(url))).blob()}
  uploadImage={async (file) => (await myUpload(file)).publicUrl}
  maxImageBytes={5 * 1024 * 1024}
/>
<SlidesEditor resolveImage={…} uploadImage={…} maxImageBytes={…} onImageError={(e) => toast(e.message)} />
```

| Prop | Behaviour |
|---|---|
| *(none)* | Default. Only local files can be inserted (file picker, drag & drop). The "By URL" tab / "Image from URL" button are hidden and the dialog says images from a web address are turned off. Files are embedded as bounded `data:` URLs. |
| `resolveImage(url)` | Enables insert-by-URL. Your function fetches the image (allow-list, proxy, auth…) and returns a `Blob` or a `data:image/…` URL; Spindle embeds the result as a data URL. A rejection or non-image result shows "Could not load image". |
| `uploadImage(file)` | Local files are handed to you and the **returned URL** is stored in the document instead of the bytes. A rejection or empty URL shows "Image upload failed". |
| `maxImageBytes` | Largest accepted file / resolved image (default 10 MB). |
| `onImageError` (slides) | Slides has no inline error UI; called when an insert fails (wrong type, too large, hook error). |

The headless pieces (`readImageFile`, `storeImageFile`, `resolveImageUrl`, `ImageHooksContext`) are exported from `@b32nio/spindle-shared/react`.

Not covered by these props: an `<img>` whose `src` is already an `http(s)` URL in the document content (loaded from a file, pasted HTML or a collaborator) is rendered by the browser like any image. Use `uploadImage` URLs you control, or sanitise on load, if that matters for your deployment.
