"""Pictures in every format, crops and transforms, linked pictures, and audio and video."""
import assets as A
import pptxbuilder as X
from dml import line, para, pic, rgb, run, scheme, solid, sp, textbox, title

SVG_EXT = ('<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="%s"/></a:ext></a:extLst>')
W, H = 2400000, 1800000


def _label(s, sid, x, y, text):
    s.add(textbox(sid, f"Label {sid}", x, y, W, 320000, para(run(text, sz=1200), algn="ctr")))


def images(path):
    d = X.Deck(title="Images")
    s = d.slide(3)
    s.add(title(2, "PNG, JPEG and GIF"))
    for n, (name, data, label) in enumerate((("picture.png", X.png(32, 32), "PNG"), ("picture.jpg", A.jpeg(), "JPEG"), ("picture.gif", A.gif(), "GIF"))):
        s.add(pic(3 + n, f"{label} picture", s.media(name, data), 700000 + n * 3700000, 1500000, W, H, descr=f"A small {label} picture used as a fixture"))
        _label(s, 10 + n, 700000 + n * 3700000, 3400000, label)
    s = d.slide(3)
    s.add(title(2, "SVG with fallback, EMF and WMF"))
    png_rid, svg_rid = s.media("vector-fallback.png", X.png(32, 32, fg=(68, 114, 196))), s.media("vector.svg", A.svg("Spindle fixture vector"))
    s.add(pic(3, "SVG with PNG fallback", png_rid, 700000, 1500000, W, H, descr="Vector picture with a PNG fallback", blip_children=SVG_EXT % svg_rid))
    s.add(pic(4, "EMF picture", s.media("picture.emf", A.emf()), 4400000, 1500000, W, H, descr="An empty enhanced metafile"))
    s.add(pic(5, "WMF picture", s.media("picture.wmf", A.wmf()), 8100000, 1500000, W, H, descr="An empty Windows metafile"))
    for n, label in enumerate(("SVG + PNG", "EMF", "WMF")):
        _label(s, 10 + n, 700000 + n * 3700000, 3400000, label)
    s = d.slide(3)
    s.add(title(2, "Crop, rotation, flips and recolour"))
    rid = s.media("picture.png", X.png(32, 32))
    s.add(pic(3, "Cropped", rid, 500000, 1500000, 1800000, 1400000, crop='<a:srcRect l="25000" t="10000" r="25000" b="40000"/>'))
    s.add(pic(4, "Rotated", rid, 2700000, 1500000, 1800000, 1400000, rot=1800000))
    s.add(pic(5, "Flipped horizontally", rid, 4900000, 1500000, 1800000, 1400000, flipH=True))
    s.add(pic(6, "Flipped vertically", rid, 7100000, 1500000, 1800000, 1400000, flipV=True))
    s.add(pic(7, "Bordered, rounded", rid, 9300000, 1500000, 1800000, 1400000, ln=line(38100, solid(rgb("C00000"))), geometry='<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'))
    s.add(pic(8, "Duotone", rid, 500000, 3600000, 1800000, 1400000, blip_children='<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>'))
    s.add(pic(9, "Brightness and contrast", rid, 2700000, 3600000, 1800000, 1400000, blip_children='<a:lum bright="30000" contrast="-20000"/>'))
    s.add(pic(10, "Greyscale", rid, 4900000, 3600000, 1800000, 1400000, blip_children="<a:grayscl/>"))
    s.add(pic(11, "Colour change", rid, 7100000, 3600000, 1800000, 1400000, blip_children='<a:clrChange><a:clrFrom><a:srgbClr val="F0F0F0"/></a:clrFrom><a:clrTo><a:srgbClr val="F0F0F0"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>'))
    s.add(pic(12, "Transparency", rid, 9300000, 3600000, 1800000, 1400000, blip_children='<a:alphaModFix amt="50000"/>'))
    s = d.slide(3)
    s.add(title(2, "Linked picture"))
    link = s.rels.add("image", "file:///nonexistent/spindle-fixture-linked.png", external=True)
    s.add(pic(3, "Linked picture", None, 700000, 1500000, W, H, link_rid=link, descr="A picture that is linked, not embedded: the target is never fetched"))
    s.add(textbox(4, "Note", 3400000, 1500000, 6000000, 800000, para("The picture above is stored as an external link to a local path that does not exist.")))
    d.write(path)


def media(path):
    d = X.Deck(title="Audio and video")
    s = d.slide(3)
    s.add(title(2, "Audio and video"))
    poster = s.media("poster.png", X.png(32, 32, fg=(60, 60, 60), bg=(200, 200, 200)))
    ms = "http://schemas.microsoft.com/office/2007/relationships/media"
    for n, (kind, tag, name, data) in enumerate((("video", "videoFile", "clip.mp4", A.mp4()), ("audio", "audioFile", "sound.wav", A.wav()))):
        d.media[name] = data
        link = s.rels.add(kind, f"../media/{name}")
        embed = s.rels.add(ms, f"../media/{name}")
        nvpr = (f'<a:{tag} r:link="{link}"/><p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}">'
                f'<p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" r:embed="{embed}"/></p:ext></p:extLst>')
        s.add(pic(3 + n, f"{kind.title()} {n + 1}", poster, 1500000 + n * 5000000, 1700000, 3000000, 2200000, nvpr=nvpr, descr=f"An inert {kind} file (not playable)"))
        s.add(textbox(10 + n, f"Label {kind}", 1500000 + n * 5000000, 4000000, 3000000, 400000, para(run(f"{kind.title()} ({name})", sz=1200), algn="ctr")))
    d.write(path)
