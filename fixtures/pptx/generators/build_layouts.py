"""Masters, layouts, themes and slide sizes."""
import pptxbuilder as X
from dml import body_ph, grad, line, para, rgb, run, scheme, sp, solid, textbox, title
from pptxbuilder import Layout, Master


def _placeholders_layout(cx, cy):
    g = X._geo(cx, cy)
    kinds = [("chart", 2), ("tbl", 3), ("dgm", 4), ("media", 5), ("clipArt", 6), ("pic", 7), (None, 8)]
    shapes = X.ph_shape(2, "Title 1", ("title", None, ""), None, X._run("Click to edit the title"))
    w, h = (cx - 2 * int(cx * 0.07)) // 4, int(cy * 0.28)
    for n, (typ, idx) in enumerate(kinds):
        col, row = n % 4, n // 4
        shapes += X.ph_shape(3 + n, f"{(typ or 'object').title()} Placeholder {idx}", (typ, idx, ' sz="quarter"'), (int(cx * 0.07) + col * w, int(cy * 0.26) + row * (h + 100000), w - 100000, h), X._run(f"{typ or 'object'} placeholder"))
    shapes += X.ph_shape(20, "Date Placeholder", ("dt", 10, ' sz="half"'), None, X.FLD["dt"]) + X.ph_shape(21, "Slide Number Placeholder", ("sldNum", 12, ' sz="quarter"'), None, X.FLD["sldNum"])
    del g
    return Layout("All placeholder types", "cust", shapes)


def layouts_masters(path):
    cx, cy = X.SLIDE_W, X.SLIDE_H
    first = X.standard_layouts(cx, cy, extra=[_placeholders_layout(cx, cy),
                                              Layout("Gradient background", "titleOnly", X.ph_shape(2, "Title 1", ("title", None, ""), None, X._run("Click to edit the title")),
                                                     bg="<p:bgPr>" + grad([(0, rgb("DDEBF7")), (100000, rgb("9DC3E6"))]) + "<a:effectLst/></p:bgPr>"),
                                              Layout("Hides master shapes", "blank", "", show_master_sp=False)])
    deco = sp(30, "Master accent bar", 0, 0, 12192000, 180000, fill=solid(scheme("accent1")))
    second = X.standard_layouts(cx, cy)
    for lay in second:
        lay.name = "Alt " + lay.name
    masters = [Master("First master", layouts=first, shapes=deco),
               Master("Second master", theme=X.theme_xml(name="Second theme", colors=[(n, v) for n, v in X.COLORS[:2]] + [("accent1", "7030A0")] + X.COLORS[3:]),
                      bg="<p:bgPr>" + solid(rgb("FFF2CC")) + "<a:effectLst/></p:bgPr>", layouts=second)]
    d = X.Deck(title="Masters and layouts", masters=masters, handout_master=True)
    s = d.slide(0, 0)
    s.add(X.ph_shape(2, "Title 1", ("ctrTitle", None, ""), None, run("Masters and layouts")))
    s.add(X.ph_shape(3, "Subtitle 2", ("subTitle", 1, ""), None, run("Two masters, placeholder types, footers")))
    s = d.slide(5, 0)
    s.add(title(2, "Every placeholder type"))
    s.add(X.ph_shape(3, "Object Placeholder 8", (None, 8, ' sz="quarter"'), None, run("Content in the generic object placeholder")))
    s.add(X.ph_shape(4, "Slide Number Placeholder 3", ("sldNum", 12, ' sz="quarter"'), None, X.FLD["sldNum"].replace("‹#›", "2")))
    s = d.slide(1, 1)
    s.add(title(2, "Second master"))
    s.add(body_ph(3, para("Text from the second master's styles", ppr_children="") + para("A second paragraph")))
    s.add(X.ph_shape(4, "Date Placeholder 3", ("dt", 10, ' sz="half"'), None, X.FLD["dt"]))
    s.add(X.ph_shape(5, "Footer Placeholder 4", ("ftr", 11, ' sz="quarter"'), None, run("Footer on the second master")))
    s.add(X.ph_shape(6, "Slide Number Placeholder 5", ("sldNum", 12, ' sz="quarter"'), None, X.FLD["sldNum"].replace("‹#›", "3")))
    s = d.slide(6, 0)
    s.add(title(2, "Gradient background layout"))
    s = d.slide(7, 0)
    s.add(textbox(2, "Plain text box", 914400, 914400, 6000000, 600000, para("This layout hides the master's accent bar.")))
    d.write(path)


def themes(path):
    dark_map = {"bg1": "dk1", "tx1": "lt1", "bg2": "dk2", "tx2": "lt2"}
    second_theme = X.theme_xml(name="Studio", clr_scheme="Studio colours", font_scheme="Studio fonts", major="Georgia", minor="Verdana", ea=("MS Mincho", "MS Gothic"), effects="shadow",
                               colors=[("dk2", "1F2A44"), ("lt2", "EDEDED"), ("accent1", "E4572E"), ("accent2", "17BEBB"), ("accent3", "FFC914"), ("accent4", "76B041"),
                                       ("accent5", "2E282A"), ("accent6", "9B5DE5"), ("hlink", "00A6FB"), ("folHlink", "8F2D56")])
    masters = [Master("Light master"), Master("Dark master", theme=second_theme, clr_map=dark_map, bg="<p:bgRef idx=\"1001\"><a:schemeClr val=\"bg1\"/></p:bgRef>")]
    d = X.Deck(title="Themes", masters=masters)
    s = d.slide(3, 0)
    s.add(title(2, "Theme colours with transforms"))
    slots = ["tx1", "bg1", "tx2", "bg2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]
    for n, slot in enumerate(slots):
        s.add(sp(3 + n, slot, 600000 + (n % 6) * 1800000, 1500000 + (n // 6) * 1300000, 1600000, 1100000, fill=solid(scheme(slot)),
                 paras=para(slot, algn="ctr"), ln=line(9525)))
    for n, mods in enumerate((("lumMod", 75000), ("lumMod", 40000), ("lumOff", 60000), ("tint", 40000), ("shade", 50000), ("satMod", 150000))):
        s.add(sp(20 + n, f"accent1 {mods[0]}", 600000 + n * 1800000, 4300000, 1600000, 900000, fill=solid(scheme("accent1", ("lumMod", mods[1]) if mods[0] == "lumMod" else mods, *((("lumOff", 60000),) if mods[0] == "lumMod" and mods[1] == 40000 else ()))),
                 paras=para(mods[0], algn="ctr")))
    s = d.slide(3, 0)
    s.add(title(2, "Colour map override on this slide"))
    s.clr_map_ovr = ('<p:clrMapOvr><a:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent3" accent2="accent4" accent3="accent5" accent4="accent6" '
                     'accent5="accent1" accent6="accent2" hlink="hlink" folHlink="folHlink"/></p:clrMapOvr>')
    s.add(sp(3, "Accent 1 remapped", 914400, 1700000, 3000000, 1200000, fill=solid(scheme("accent1")), paras=para("accent1 maps to accent3 here", algn="ctr")))
    s = d.slide(3, 1)
    s.add(title(2, "Dark master with theme fonts"))
    s.add(textbox(3, "Major font", 914400, 1700000, 9000000, 700000, para(run("Major font (Georgia)", font="+mj-lt"))))
    s.add(textbox(4, "Minor font", 914400, 2500000, 9000000, 700000, para(run("Minor font (Verdana)", font="+mn-lt"))))
    s.add(textbox(5, "East Asian font", 914400, 3300000, 9000000, 700000, para(run("日本語のテキスト", ea="+mn-ea", font="+mn-lt", lang="ja-JP"))))
    s.add(sp(6, "Effect style", 914400, 4200000, 3000000, 1000000, fill=solid(scheme("accent1")), paras=para("Shape style effects", algn="ctr"),
             style='<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>'))
    d.write(path)


def slide_size(path, size, size_type, label):
    d = X.Deck(title=f"Slide size: {label}", size=size, size_type=size_type)
    cx, cy = size
    s = d.slide(0)
    s.add(X.ph_shape(2, "Title 1", ("ctrTitle", None, ""), None, run(f"Slide size: {label}")))
    s.add(X.ph_shape(3, "Subtitle 2", ("subTitle", 1, ""), None, run(f"{cx} by {cy} EMU")))
    s = d.slide(1)
    s.add(title(2, "Content scales with the page"))
    s.add(body_ph(3, para("Placeholders inherit their position from the layout") + para("so they follow the page size")))
    s = d.slide(3)
    s.add(title(2, "Fixed shapes"))
    s.add(sp(3, "Page frame", 0, 0, cx, cy, fill="<a:noFill/>", ln=line(38100, solid(rgb("C00000")))))
    s.add(sp(4, "Centre ellipse", cx // 2 - cx // 8, cy // 2 - cy // 8, cx // 4, cy // 4, prst="ellipse", fill=solid(scheme("accent1")), paras=para("centre", algn="ctr")))
    d.write(path)


def sizes_4x3(path):
    slide_size(path, (9144000, 6858000), "screen4x3", "4:3")


def sizes_a4(path):
    slide_size(path, (10692000, 7560000), "A4", "A4")


def sizes_custom(path):
    slide_size(path, (10058400, 6400800), "custom", "custom 11 x 7 in")


def sizes_portrait(path):
    slide_size(path, (6858000, 9144000), None, "portrait 3:4")


def sizes_16x9(path):
    slide_size(path, (X.SLIDE_W, X.SLIDE_H), None, "16:9")
