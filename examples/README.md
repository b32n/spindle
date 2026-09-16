# Examples

Standalone demos for the Spindle libraries — each runs without a backend or
authentication.

| Demo | Directory | Showcases |
|------|-----------|-----------|
| Docs | [`docs-demo/`](docs-demo/) | The document editor (`@b32nio/spindle-docs-react`) |
| Sheets | [`sheets-demo/`](sheets-demo/) | The spreadsheet (`@b32nio/spindle-sheets-react`) |
| Slides | [`slides-demo/`](slides-demo/) | The presentation editor (`@b32nio/spindle-slides-react`) |
| Collab server | [`collab-server/`](collab-server/) | A local WebSocket relay for trying real-time collaboration across the three editors |

## Run a demo

```bash
cd examples/sheets-demo   # or docs-demo / slides-demo
npm install
npm run dev
```

Each demo is pre-loaded with sample content and exercises the library's core
features — handy for testing and development. The collaboration demos can point
at the local relay in `collab-server/`.
