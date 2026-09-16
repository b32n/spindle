# Contributing to Spindle

Thanks for your interest in Spindle — the open-source editing layer for the
web (spreadsheets, documents, and slides as React libraries). Contributions
of all kinds are welcome: bug reports, docs, and code.

## Project layout

Spindle is an npm-workspaces monorepo:

```
packages/
  spindle-shared               Framework-agnostic utilities & collaboration primitives
  spindle-sheets-core          Spreadsheet engine (sparse store, formulas) — zero React
  spindle-sheets-react         React canvas grid, formula bar, toolbar & dialogs
  spindle-docs-core            Document engine (True Layout) — zero React
  spindle-docs-react           React document editor components
  spindle-slides-core          Presentation engine (scene layer, fractional index) — zero React
  spindle-slides-react         React slide editor: stage, gestures, present mode
  spindle-transport-websocket  WebSocket CollabProvider for real-time collaboration
website/                       The spindle.b32n.io marketing + docs site (standalone)
documentation/                 Markdown docs (rendered by the website)
examples/                      Runnable demos
```

The `-core` packages carry **zero React** — keep framework code out of them.
Cross-package types resolve through each package's built `dist/*.d.ts`, so
after changing a public API, rebuild before trusting a workspace type-check.

## Getting started

Requires Node `>=18`.

```bash
npm install            # install root + workspace deps
npm run build          # build packages in dependency order
npm test               # run the test suites
npm run type-check     # type-check every package
npm run lint           # eslint over packages
```

## Development workflow

1. Branch off an up-to-date `master`.
2. Make your change with tests. Core logic (formulas, storage, collaboration,
   layout) should have unit tests in the relevant `-core` package.
3. Ensure `npm run build`, `npm test`, `npm run type-check`, and
   `npm run lint` all pass.
4. Open a pull request against `master` describing the change and why.

## Coding conventions

- **TypeScript strict mode** throughout — no `any` escape hatches without a
  clear reason.
- Match the style of the surrounding code; keep the `-core`/`-react` boundary
  clean.
- Update the docs under `documentation/` when you change public behavior — the
  website renders them directly.

## Reporting bugs

Open a GitHub issue with a minimal reproduction (a code snippet or a failing
test is ideal), what you expected, and what happened instead.

## License

By contributing, you agree that your contributions are licensed under the
project's [MIT License](LICENSE).
