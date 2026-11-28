# PPTX real-app checks (D8)

Filled by the maintainer for Spindle exports opened in PowerPoint and Google Slides. This is manual testing, not library behaviour.

| Fixture | Producer | App + version | Opens cleanly | Repair prompt | Notes |
| --- | --- | --- | --- | --- | --- |
| any `exportPptx` output with text and shapes (K12b) | Spindle | _pending_ | _pending_ | _pending_ | Check that text boxes, placeholders, fills, outlines, bullets, fields and hyperlinks look as in the source deck, and that the list style `a:lstStyle` written on each shape gives the right default size and colour. |
| any `exportPptx` output (K12a) | Spindle | _pending_ | _pending_ | _pending_ | Check that the master written without `p:bg` and with a single empty `a:lvl1pPr` in `p:txStyles` opens without a repair prompt and shows the layouts' placeholders. |
