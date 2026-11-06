#!/usr/bin/env python3
"""Render a report PDF so the pages can be looked at.
usage: python -X utf8 render_pages.py report.pdf outdir            contact sheets of every page (sheet0.png, ...)
       python -X utf8 render_pages.py report.pdf outdir 1,2,3 90   the listed physical pages at the given dpi (p1.png, ...)
"""
import os, sys
import fitz
from PIL import Image

if len(sys.argv) < 3:
    sys.exit(__doc__)
pdf, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
d = fitz.open(pdf)
if len(sys.argv) > 3:
    dpi = int(sys.argv[4]) if len(sys.argv) > 4 else 90
    for n in (int(x) for x in sys.argv[3].split(",")):
        d[n - 1].get_pixmap(dpi=dpi).save(os.path.join(out, f"p{n}.png"))
    print("wrote", sys.argv[3], "to", out)
else:
    ims = []
    for p in d:
        px = p.get_pixmap(dpi=45)
        ims.append(Image.frombytes("RGB", (px.width, px.height), px.samples))
    w, h = ims[0].size
    cols = 10
    for k in range(0, len(ims), 30):
        sub = ims[k:k + 30]
        rows = (len(sub) + cols - 1) // cols
        c = Image.new("RGB", (w * cols, h * rows), "white")
        for j, im in enumerate(sub):
            c.paste(im, ((j % cols) * w, (j // cols) * h))
        c.save(os.path.join(out, f"sheet{k // 30}.png"))
    print(len(ims), "pages,", (len(ims) + 29) // 30, "sheets in", out)
