"""Files that are never committed (too large or hostile): written on demand, stdlib only.

    python3 build_runtime.py --out DIR [zip-bomb|huge-dimension|perf-100k-x-10 ...]

edge-zip-bomb      a ~300 KB package whose sheet part holds a 300 MB XML comment (compression ratio ~1000:1)
edge-huge-dimension a tiny package declaring a sheet dimension of A1:XFD1048576
perf-100k-x-10     100,000 rows x 10 columns of numbers and short strings (about 5 MB)
"""
import argparse
import os
import zipfile
import zlib

import xlsxbuilder as X


def _skeleton(rows_xml_stream=None, dimension="A1"):
    """A one-sheet package whose worksheet XML is replaced by the caller; returns (parts, sheet_head, sheet_tail)."""
    b = X.Book(minimal=True, title="Runtime fixture")
    b.sheet("Sheet1")
    parts = b.build()
    head = f'{X.HEAD}<worksheet xmlns="{X.M}"><dimension ref="{dimension}"/><sheetData>'
    return parts, head, "</sheetData></worksheet>"


def _write(path, parts, sheet_chunks):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in sorted(parts, key=lambda n: (n != "[Content_Types].xml", n)):
            zi = zipfile.ZipInfo(n, date_time=X.STAMP)
            zi.compress_type = zipfile.ZIP_DEFLATED
            if n == "xl/worksheets/sheet1.xml":
                with z.open(zi, "w", force_zip64=True) as f:
                    for chunk in sheet_chunks:
                        f.write(chunk)
            else:
                z.writestr(zi, parts[n])


def zip_bomb(path):
    parts, head, tail = _skeleton()
    block = b" " * (1 << 20)
    chunks = [head.encode(), b'<row r="1"><c r="A1"><v>1</v></c></row></sheetData><!--']
    chunks += [block] * 300 + [b"--></worksheet>"]
    _write(path, parts, chunks)


def huge_dimension(path):
    parts, head, tail = _skeleton(dimension="A1:XFD1048576")
    _write(path, parts, [head.encode(), b'<row r="1"><c r="A1"><v>1</v></c></row>', tail.encode()])


def perf(path, rows=100_000, cols=10):
    parts, head, tail = _skeleton(dimension=f"A1:{X.col(cols)}{rows}")

    def gen():
        yield head.encode()
        for r in range(1, rows + 1):
            cells = "".join(f'<c r="{X.ref(r, c)}"><v>{(r * c) % 9973 + c / 8}</v></c>' if c % 3 else f'<c r="{X.ref(r, c)}" t="inlineStr"><is><t>item {r}-{c}</t></is></c>' for c in range(1, cols + 1))
            yield f'<row r="{r}">{cells}</row>'.encode()
        yield tail.encode()

    _write(path, parts, gen())


BUILDERS = {"edge-zip-bomb": zip_bomb, "edge-huge-dimension": huge_dimension, "perf-100k-x-10": perf}

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", required=True)
    ap.add_argument("only", nargs="*")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    for name in a.only or sorted(BUILDERS):
        p = os.path.join(a.out, name + ".xlsx")
        BUILDERS[name](p)
        print(f"wrote {p} ({os.path.getsize(p)} bytes)")
