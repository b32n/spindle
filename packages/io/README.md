# @b32nio/spindle-io

Optional XLSX and DOCX import/export for Spindle. It holds the format-level
readers and writers (and their heavy dependencies: `exceljs`, `docx`, `jszip`,
`fast-xml-parser`) so `@b32nio/spindle-shared` stays light. Everything runs
locally; there are no network calls.

`@b32nio/spindle-sheets-core` and `@b32nio/spindle-docs-core` depend on this
package and load it on first Open/Download, so editor hosts need no extra setup.
Install it directly for headless conversion.

```bash
npm install @b32nio/spindle-io
```

```typescript
import { exportXlsx, importXlsx, detectXlsxLoss } from '@b32nio/spindle-io/xlsx';
import { exportDocx, importDocx, detectDocxLoss } from '@b32nio/spindle-io/docx';
```

- `@b32nio/spindle-io/xlsx` — `exportXlsx`, `importXlsx`, `detectXlsxLoss` (needs `exceljs`)
- `@b32nio/spindle-io/docx` — `exportDocx`, `importDocx`, `detectDocxLoss` (needs `docx`, `jszip`, `fast-xml-parser`)
- `@b32nio/spindle-io` — both, plus the CSV helpers and `ImportLossReport` utilities re-exported from `@b32nio/spindle-shared`

Prefer the subpaths so a consumer that needs one format never loads the other's libraries.

## Migrating from `@b32nio/spindle-shared`

`@b32nio/spindle-shared/export/xlsx` and `/export/docx` still work as deprecated async shims
that load this package on demand (it must be installed). Switch the import path to
`@b32nio/spindle-io/xlsx` or `/docx`; the function signatures are unchanged. The input/output
types stay exported from `@b32nio/spindle-shared/export/{xlsx,docx}` and are re-exported here.
