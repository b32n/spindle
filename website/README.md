# Spindle website

The marketing site for [Spindle](https://github.com/b32n/spindle), served at
**spindle.b32n.io**. A small React SPA built with Vite.

## Develop

```bash
cd website
npm install
npm run dev        # http://localhost:5173
```

## Build

```bash
npm run build      # outputs to website/dist
npm run preview    # preview the production build locally
```

## Deploy — Cloudflare Pages

Connect the `b32n/spindle` repo in the Cloudflare Pages dashboard and set:

| Setting | Value |
| --- | --- |
| Framework preset | Vite |
| Build command | `npm run build` |
| Build output directory | `dist` |
| Root directory | `website` |

`public/_redirects` sends every path to `index.html` (SPA fallback). Add
`spindle.b32n.io` as a custom domain in the Pages project once the first
deploy is green.
