"""Manifest data for the PPTX corpus: what each file is, what it claims to contain and how it is made.

`constructs` are claims: the Jest suite checks each one against the package XML with the probes in
scripts/pptx-fidelity/constructs.ts, and every construct in that registry must be claimed by some file.
"""
LICENSE = "Original content authored for this repository (MIT, same as the repo)"
OS = "Debian GNU/Linux 13 (trixie), Linux x86_64 container"


def f(id, title, constructs, notes, generator, tool="programmatic", ext="pptx", **extra):
    return {"id": id, "filename": f"{id}.{ext}", "title": title, "tool": tool, "constructs": constructs.split(), "notes": notes, "generator": f"generators/{generator}", **extra}


TEXT_COMMON = "slides.multiple text.basic"

CATALOG = [
    f("text-basic", "Basic text", TEXT_COMMON + " text.runs text.alignment text.levels text.bullets text.numbering text.hyperlinks text.spacing text.highlight text.insets text.fields layouts.footerPlaceholders layouts.inheritance shapes.textBox fills.themeColors",
      "A title slide, five paragraph levels with bullet characters, auto-numbered paragraphs (Arabic and lower-case Roman starting at 3), an external and a mailto hyperlink, bold / italic / underline / strike / size / colour / theme colour / font / highlight runs, the four alignments, paragraph spacing (before, after, percent and exact line spacing), custom text insets, slide-number and date fields, and footer and slide-number placeholders on a slide.", "build_text.py"),
    f("text-rich", "Rich text", "slides.multiple text.basic text.tabStops text.columns text.vertical text.autofit text.autofitShape text.bulletsCustom text.bulletsPicture text.bullets text.numbering text.baseline text.caps text.charSpacing shapes.textBox",
      "Tab stops (left, centre, right), a two-column text box, four vertical text modes (vert, vert270, eaVert, wordArtVert), shrink-on-overflow autofit with fontScale and lnSpcReduction, resize-shape-to-fit-text, superscript and subscript, all caps, small caps and character spacing, bullets with colour, size, a symbol font and a picture, and a no-wrap text box.", "build_text.py"),
    f("layouts-masters", "Masters and layouts", TEXT_COMMON + " masters.multiple masters.handoutMaster masters.txStyles masters.background masters.shapes layouts.placeholders layouts.placeholderKinds layouts.background layouts.hideMasterShapes layouts.inheritance layouts.footerPlaceholders theme.multiple text.fields",
      "Two slide masters (the second with its own theme and background), a handout master, a layout with every placeholder type (chart, table, diagram, media, clip art, picture, object), a layout with a gradient background, a layout that hides master shapes, a decorative shape on the first master, and slides that fill the placeholders and carry date, footer and slide-number placeholders.", "build_layouts.py"),
    f("themes", "Themes", TEXT_COMMON + " theme.multiple theme.colorMap theme.fontScheme theme.fontsEastAsian theme.effectScheme fills.themeColors fills.colorTransforms fills.solid",
      "Two masters with different themes: the default one and a second one with its own colour scheme, Georgia / Verdana fonts, East Asian fonts and effect styles with shadow, glow and soft edge, mapped dark through the master colour map. One slide overrides the colour map. Slides show all twelve theme colour slots, colour transforms, theme fonts and a shape whose effect comes from its style reference.", "build_layouts.py"),
    f("shapes-presets", "Every preset geometry", TEXT_COMMON + " shapes.presets shapes.presetKinds shapes.adjustments shapes.flip shapes.rotation shapes.textBox fills.solid fills.themeColors fills.colorTransforms",
      "All 187 ST_ShapeType preset geometries once (generated from the list in build_shapes.py, 24 per slide, each labelled), then shapes with adjustment values, rotated, flipped and rotated-and-flipped shapes, a text box and text inside a shape.", "build_shapes.py"),
    f("shapes-custgeom", "Custom geometry", "text.basic shapes.custGeom fills.solid fills.themeColors",
      "Freeform shapes: a straight-line triangle, cubic and quadratic Bezier curves, arcs in two paths (one unfilled, one unstroked), a shape with adjust values, guides, connection sites and a text rectangle, and a hand-drawn star.", "build_shapes.py"),
    f("fills-lines", "Fills and lines", TEXT_COMMON + " fills.solid fills.noFill fills.gradient fills.gradientPath fills.pattern fills.picture fills.alpha fills.background fills.themeColors fills.colorTransforms lines.dash lines.arrowheads lines.joinCap lines.compound",
      "Solid, gradient (linear, three stops, circle and rectangle paths), pattern, picture (stretched and tiled), transparent and empty fills; eleven dash styles, five arrowhead types, caps, joins, widths, compound lines, a gradient line and a connector line; and four slide backgrounds (gradient, tiled picture, solid, default).", "build_shapes.py"),
    f("effects", "Effects", TEXT_COMMON.split()[0] + " effects.shadow effects.glow effects.reflection effects.softEdge effects.blur effects.scene3d effects.sp3d effects.bevel effects.text fills.solid fills.themeColors fills.alpha",
      "Outer, inner and preset shadows, glow, reflection, soft edges, blur and a shadow-and-glow combination, then 3-D formatting: an extruded bevelled shape, a perspective camera, shadowed text and an effect taken from the theme through a style reference.", "build_shapes.py"),
    f("sizes-16x9", "Slide size 16:9", "sizes.16x9 slides.multiple text.basic", "A 12192000 x 6858000 EMU deck (widescreen 16:9, no type attribute), the size every other deck here uses.", "build_layouts.py"),
    f("sizes-4x3", "Slide size 4:3", "sizes.4x3 slides.multiple text.basic", "A 9144000 x 6858000 EMU deck (the classic 4:3 size, type screen4x3) with a title slide, a content slide and fixed shapes.", "build_layouts.py"),
    f("sizes-a4", "Slide size A4", "sizes.a4 slides.multiple text.basic", "An A4 landscape deck (10692000 x 7560000 EMU, type A4).", "build_layouts.py"),
    f("sizes-custom", "Custom slide size", "sizes.custom slides.multiple text.basic", "A custom 11 x 7 inch deck (type custom).", "build_layouts.py"),
    f("sizes-portrait", "Portrait slides", "sizes.portrait slides.multiple text.basic", "A portrait 3:4 deck (6858000 x 9144000 EMU).", "build_layouts.py"),
    f("edge-truncated-zip", "Edge: truncated zip", "edge.truncatedZip", "A deck cut off at 60 percent of its bytes: it starts as a zip and has no central directory.", "build_edge.py", ext="pptx"),
    f("edge-no-presentation", "Edge: no presentation part", "edge.noPresentation slides.multiple", "A package with slides, masters and a theme but no ppt/presentation.xml (and no relationship to it).", "build_edge.py"),
    f("edge-slide-missing", "Edge: listed slide without its part", "edge.slideMissing edge.brokenRels slides.multiple", "presentation.xml lists three slides and relates to all of them, but ppt/slides/slide2.xml is not in the package.", "build_edge.py"),
    f("edge-bad-rels", "Edge: relationships that point nowhere", "edge.brokenRels slides.multiple", "Slide relationships that point at a missing image, a missing layout and a part outside the package root, next to an external hyperlink (which is fine).", "build_edge.py"),
    f("edge-invalid-xml", "Edge: malformed slide XML", "edge.invalidXml slides.multiple", "A slide with an unclosed element and a stray closing tag; the other slides are fine.", "build_edge.py"),
    f("edge-deep-groups", "Edge: deeply nested groups", "edge.deepGroups", "One slide whose innermost shape sits in 60 nested groups.", "build_edge.py"),
    f("edge-relationship-cycles", "Edge: relationship cycles", "edge.relCycle slides.multiple", "Two slides whose relationship parts point at each other, and one that points at itself.", "build_edge.py"),
    f("edge-encrypted", "Edge: encrypted presentation (OLE stub)", "edge.encrypted", "An OLE compound-file stub with EncryptionInfo and EncryptedPackage directory entries (a password-protected file), header and directory only.", "build_edge.py"),
    f("edge-legacy-ppt", "Edge: legacy .ppt (OLE stub)", "edge.legacyBinary", "An OLE compound-file stub with a PowerPoint Document stream entry (the binary format before 2007), header and directory only.", "build_edge.py", ext="ppt"),
]

RUNTIME = [
    f("edge-1000-slides", "Edge: 1000 slides (generated at run time)", "edge.manySlides slides.multiple text.basic", "A deck with 1000 slides (about 900 KB, so never committed): write it with generators/build_runtime.py.", "build_runtime.py", generatedAtRuntime=True),
]
