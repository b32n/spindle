"""Files that are never committed (too large): written on demand, stdlib only.

    python3 build_runtime.py --out DIR [edge-1000-slides ...]

edge-1000-slides   a package with 1000 slides (each one placeholder with a line of text), about 900 KB
"""
import argparse
import os

import pptxbuilder as X
from dml import run, title


def many_slides(path, count=1000):
    d = X.Deck(title="One thousand slides")
    for n in range(1, count + 1):
        d.slide(1).add(title(2, f"Slide {n}")).add(
            f'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
            f"<p:txBody><a:bodyPr/><a:lstStyle/><a:p>{run(f'Body text of slide {n}')}</a:p></p:txBody></p:sp>")
    d.write(path)


BUILDERS = {"edge-1000-slides": many_slides}

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("only", nargs="*")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    for name in a.only or sorted(BUILDERS):
        p = os.path.join(a.out, name + ".pptx")
        BUILDERS[name](p)
        print(f"wrote {p} ({os.path.getsize(p)} bytes)")
