# Spindle Docs — Overview

Spindle Docs is a high-performance document editor library for React, providing a Google Docs-like editing experience with true page-based layout and real-time pagination.

## Architecture

The library is structured as two packages with a clear separation of concerns:

```mermaid
graph TD
    A[Spindle Docs] --> B[@b32nio/spindle-docs-core]
    A --> C[@b32nio/spindle-docs-react]
    B --> D[Document Model]
    B --> E[ProseMirror Integration]
    B --> F[Block Types]
    C --> G[TrueLayoutEditor]
    C --> H[Selection Overlay]
    C --> I[Input Bridge]
    C --> J[DOM Painter]
```

### Core Package (`@b32nio/spindle-docs-core`)

The framework-agnostic document engine:

- **Document model** — the root container with title, sections, and page configuration
- **ProseMirror integration** — schema, plugins, and editing commands
- **Block types** — paragraphs, headings, lists, tables, images, and more
- **True Layout paginator** — line-level pagination that mirrors print output

### React Package (`@b32nio/spindle-docs-react`)

The React editor built on the core engine:

- **TrueLayoutEditor** — the paginated editing surface
- **Toolbar, ruler, and dialogs** — the document UI chrome
- **Selection overlay & input bridge** — native-feeling cursor and text input

## Design Principles

### Separation of Concerns

The core package is completely framework-agnostic and holds all business logic; the React package supplies the interface. This keeps UI and engine boundaries clean, makes the core easy to test, and leaves room to port the editor to other frameworks later.

### Print-True Pagination

Layout decisions prioritize fidelity to the printed page:

- **Line-level pagination** — content flows across real page boundaries
- **Live repagination** — edits reflow pages incrementally, not all at once
- **Headers and footers** — per-section configuration that survives layout

### Type Safety

TypeScript strict mode is enforced throughout, with comprehensive type definitions and interface segregation for clean, predictable APIs.

## Next Steps

- **[Architecture](architecture.md)** — core and React package internals, extension points, and performance
- **[Component Reference](components.md)** — every React component and its props
- **[Data Structures](data-structures.md)** — the document model and type definitions
