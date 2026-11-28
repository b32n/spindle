#!/usr/bin/env python3
"""Writes the PPTX fixtures of the slides file-I/O E2E specs into e2e/fixtures (plan K15).

    python3 scripts/testing/make-pptx-fixtures.py [--out DIR]

Deterministic (fixed zip timestamps, no randomness): rebuilding gives the same bytes. Built with the stdlib builder of the PPTX
fidelity corpus (fixtures/pptx/generators), so no tool beyond python3 is needed. Original content authored for this repository (MIT).

- six-slides.pptx: six text slides (a title and a bullet list each); the first slide reads "Quarterly review".
- transitions-charts.pptx: three slides, a fade transition on the first, a bar chart with its workbook on the second and a push
  transition on the third: what the import dialog lists under "Kept, not editable".
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.join(ROOT, "fixtures", "pptx", "generators"))

import build_charts as C  # noqa: E402
import pptxbuilder as X  # noqa: E402
from dml import body_ph, para, title  # noqa: E402

SIX = [
    ("Quarterly review", ["Revenue grew in every region", "Costs stayed flat"]),
    ("Highlights", ["Three launches shipped", "Support backlog halved"]),
    ("Customers", ["Two large renewals", "Net retention above plan"]),
    ("Risks", ["Hiring is behind", "One supplier is late"]),
    ("Roadmap", ["Q3: collaboration", "Q4: offline mode"]),
    ("Thank you", ["Questions are welcome"]),
]


def six_slides(path):
    d = X.Deck(title="Six slides")
    for heading, points in SIX:
        s = d.slide(1)
        s.add(title(2, heading))
        s.add(body_ph(3, "".join(para(p) for p in points)))
    d.write(path)


def transitions_charts(path):
    d = X.Deck(title="Transitions and charts")
    s = d.slide(3)
    s.add(title(2, "Fade in"))
    s.transition = '<p:transition spd="med"><p:fade/></p:transition>'
    s = d.slide(3)
    s.add(title(2, "Bar chart"))
    C.add_chart(d, s, "bar", 3, 1500000, 1500000, 8000000, 4300000)
    s = d.slide(3)
    s.add(title(2, "Push out"))
    s.transition = '<p:transition spd="slow" advClick="0" advTm="5000"><p:push dir="u"/></p:transition>'
    d.write(path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "e2e", "fixtures"))
    out = ap.parse_args().out
    os.makedirs(out, exist_ok=True)
    six_slides(os.path.join(out, "six-slides.pptx"))
    transitions_charts(os.path.join(out, "transitions-charts.pptx"))
