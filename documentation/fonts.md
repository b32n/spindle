# Fonts and privacy

Spindle never loads fonts from Google (or any third party) on its own. The three editors take a
`fontSource` prop that says where web fonts come from:

```tsx
<WorkbookCanvas fontSource={{ kind: 'local' }} />      // sheets
<DocumentEditor fontSource={{ kind: 'local' }} />      // docs
<SlidesEditor   fontSource={{ kind: 'local' }} />      // slides
```

| `fontSource` | Behaviour |
|---|---|
| *(omitted)* | Default. System fonts first; web fonts (Roboto, Lato, Inter, …) are requested from the Spindle font mirror, `https://fonts.b32n.io/css2`. If the mirror is unreachable the stylesheet is dropped and text renders in the next font in the stack (every stack ends in a generic family). |
| `{ kind: 'local' }` | **Zero font requests.** Only fonts installed on the user's machine render. Use this for air-gapped or strict-CSP deployments. |
| `{ kind: 'css', url }` | Load from your own stylesheet endpoint. It must accept the Google Fonts `css2` query format: `<url>?family=Roboto:wght@400&family=Lato:wght@400&display=swap`. |
| `{ kind: 'custom', load }` | `load(families: string[])` is called once per new web family; inject `@font-face` rules, preload files, or call your own service. Errors are swallowed and fall back to the font stack. |

Fonts that ship with common operating systems (Arial, Georgia, Times New Roman, Verdana, Courier New, …)
are never requested.

Where fonts are used:

- **Sheets**: the toolbar font picker requests the families it lists (first 50 matches) when it is opened.
- **Docs**: the toolbar font menu's web families are requested when the editor mounts.
- **Slides**: the deck theme's heading/body fonts are requested when the editor mounts.

If you use the editors without these components, the same behaviour is available as
`createFontLoader(source)` from `@b32nio/spindle-shared/react`.

## Alternatives to the default mirror

- **Self-host**: serve `css2`-compatible CSS plus the woff2 files from your own origin (for example built from
  the OFL-licensed [Fontsource](https://fontsource.org) packages) and pass `{ kind: 'css', url: 'https://your.host/css2' }`.
- **Bunny Fonts**: a privacy-friendlier, GDPR-oriented third party with a Google-compatible API:
  `{ kind: 'css', url: 'https://fonts.bunny.net/css2' }`.
- **Bundle the fonts** with your app (`@fontsource/*`) and use `{ kind: 'local' }` so Spindle makes no requests.

## Content Security Policy

With the default source you need `style-src https://fonts.b32n.io` and `font-src` for the mirror's font host.
With `{ kind: 'local' }` no CSP exceptions are needed.
