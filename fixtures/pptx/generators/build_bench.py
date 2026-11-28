"""Benchmark decks for the local review pack (plan K17); never committed, written on demand, stdlib only, deterministic.

    python3 build_bench.py --out DIR [bench-100-slides bench-200-slides ...]

Every slide holds a title, a four-line body, a filled rectangle and an ellipse with text; every fifth slide also holds a
small picture (one shared media part), every tenth a table-like grid of text boxes. No randomness, no real data.
"""
import argparse
import os

import pptxbuilder as X
from dml import para, run, sp, solid, rgb, pic, title

PALETTE = ["4472C4", "ED7D31", "70AD47", "FFC000", "5B9BD5"]


def bench_deck(path, count):
    d = X.Deck(title=f"Benchmark deck of {count} slides")
    for n in range(1, count + 1):
        s = d.slide(1)
        s.add(title(2, f"Slide {n} of {count}"))
        lines = "".join(para(run(f"Line {k} of slide {n}: the quick brown fox jumps over the lazy dog")) for k in range(1, 5))
        s.add(f'<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>'
              f"<p:txBody><a:bodyPr/><a:lstStyle/>{lines}</p:txBody></p:sp>")
        colour = PALETTE[n % len(PALETTE)]
        s.add(sp(4, "Block", 8000000, 1600000, 3000000, 1400000, fill=solid(rgb(colour)), paras=para(run(f"Block {n}"))))
        s.add(sp(5, "Badge", 8000000, 3400000, 2000000, 2000000, prst="ellipse", fill=solid(rgb(PALETTE[(n + 2) % len(PALETTE)])), paras=para(run(str(n)))))
        sid = 6
        if n % 5 == 0:
            s.add(pic(sid, "Picture", s.media("bench.png", X.png(32, 32)), 600000, 5200000, 1200000, 1200000))
            sid += 1
        if n % 10 == 0:
            for row in range(3):
                for col in range(3):
                    s.add(sp(sid, f"Cell {row}{col}", 2000000 + col * 1500000, 4800000 + row * 600000, 1500000, 600000,
                             fill=solid(rgb("E7E6E6")), paras=para(run(f"r{row}c{col}"))))
                    sid += 1
    d.write(path)


BUILDERS = {"bench-100-slides": lambda p: bench_deck(p, 100), "bench-200-slides": lambda p: bench_deck(p, 200)}

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
