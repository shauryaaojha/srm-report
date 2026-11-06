#!/usr/bin/env python3
"""Read an SRM report template (or a finished report) and write down everything a builder must know.

usage: python -X utf8 inspect_template.py <template.docx> <outdir>

Writes to <outdir>:
  structure.txt  every body element: index, type, style, size, bold, align, markers, text
  rules.txt      sentences that state a rule, inline notes like <Font Size 18>, and placeholders to fill
  sections.txt   each section's page size, margins, numbering and footer text
Prints a short summary. Read all three files before planning the report.
"""
import os, re, sys
from collections import Counter
import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

RULE_WORDS = re.compile(
    r"\b(must|should|shall|minimum|maximum|mandatory|at least|not more than|font|margin|spacing|"
    r"justify|similarity|turnitin|plagiarism|caption|binding|pages?\b.*\b(excluding|minimum|maximum)|"
    r"every chapter|page numbers?|references? (must|to be)|hard ?bound|spiral)\b", re.I)
NOTE = re.compile(r"<(?!<)[^<>\n]{2,80}>(?!>)")
PLACEHOLDER = re.compile(
    r"<<[^<>]+>>|\[[A-Z][A-Za-z .]*(NAME|NUM|NO)[^\]]*\]|\.{5,}|_{5,}|\(GUIDE NAME\s*\)|\(Designation[^)]*\)|"
    r"\(SPECIALIZATION NAME\)|STUDENT ?\d ?NAME|PROJECT TITLE|TITLE OF THE [A-Z]+|Name of the [Ss]tudent|Dr\. ?ABCD")


def ptext(el):
    out = []
    for x in el.iter():
        if x.tag == qn("w:t") and x.text: out.append(x.text)
        elif x.tag == qn("w:tab"): out.append("\t")
        elif x.tag == qn("w:br"): out.append(" / ")
    return "".join(out)


def describe(p):
    size = bold = None
    for r in p.runs:
        if r.text.strip():
            size = r.font.size.pt if r.font.size else None
            bold = r.bold
            break
    if size is None and p.style is not None and p.style.font.size:
        size = p.style.font.size.pt
    if bold is None and p.style is not None:
        bold = p.style.font.bold
    al = p.alignment
    return size, bold, (al.name if al is not None else "")


def main(path, out):
    os.makedirs(out, exist_ok=True)
    d = docx.Document(path)
    body = d.element.body
    lines, rules, notes, holders = [], [], Counter(), Counter()

    def scan(t, where):
        for m in NOTE.finditer(t): notes[m.group(0)] += 1
        for m in PLACEHOLDER.finditer(t): holders[m.group(0).strip()] += 1
        if RULE_WORDS.search(t) and len(t) > 25: rules.append(f"[{where}] {t.strip()}")

    for i, el in enumerate(body.iterchildren()):
        if el.tag == qn("w:sectPr"):
            lines.append(f"{i:4d} FINAL-SECTPR")
            continue
        if el.tag == qn("w:tbl"):
            t = Table(el, d)
            n_rows, n_cols = len(t.rows), len(t.columns)
            first = " | ".join(c.text.strip().replace("\n", " / ")[:70] for c in t.rows[0].cells)
            lines.append(f"{i:4d} TABLE {n_rows}x{n_cols}: {first}")
            for tr in el.iter(qn("w:p")):
                scan(ptext(tr), f"{i} table")
            continue
        if el.tag != qn("w:p"):
            lines.append(f"{i:4d} {el.tag.split('}')[1]}")
            continue
        p = Paragraph(el, d)
        t = ptext(el)
        size, bold, al = describe(p)
        marks = []
        if el.find(".//" + qn("w:sectPr")) is not None: marks.append("SECT")
        if el.find(".//" + qn("w:pageBreakBefore")) is not None: marks.append("PBB")
        if any(b.get(qn("w:type")) == "page" for b in el.iter(qn("w:br"))): marks.append("PB")
        n_img = len(list(el.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip")))
        if n_img: marks.append(f"IMG{n_img}")
        if el.find(".//" + qn("w:txbxContent")) is not None: marks.append("TEXTBOX")
        if el.find(".//" + qn("w:numPr")) is not None: marks.append("LIST")
        if el.find(".//" + qn("w:instrText")) is not None: marks.append("FIELD")
        style = p.style.name if p.style is not None else "-"
        if t.strip() or marks:
            lines.append(f"{i:4d} P [{style}|{size or ''}|{'B' if bold else ''}|{al}]{' {' + ','.join(marks) + '}' if marks else ''} {t[:170]}")
        scan(t, str(i))

    secs = []
    cm = lambda v: round(v / 360000, 2) if v is not None else None
    for i, s in enumerate(d.sections):
        pg = s._sectPr.find(qn("w:pgNumType"))
        num = dict((k.split("}")[1], v) for k, v in pg.attrib.items()) if pg is not None else {}
        foot = " ".join(p.text for p in s.footer.paragraphs).strip() if not s.footer.is_linked_to_previous else "(linked)"
        fld = bool(s.footer._element.findall(".//" + qn("w:instrText")) or s.footer._element.findall(".//" + qn("w:fldSimple"))) if not s.footer.is_linked_to_previous else None
        secs.append(f"{i}: page {cm(s.page_width)}x{cm(s.page_height)} cm  margins L {cm(s.left_margin)} R {cm(s.right_margin)} "
                    f"T {cm(s.top_margin)} B {cm(s.bottom_margin)} cm  numbering {num}  footer {foot!r} pagefield={fld}")

    with open(os.path.join(out, "structure.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    with open(os.path.join(out, "sections.txt"), "w", encoding="utf-8") as f:
        f.write("\n".join(secs))
    with open(os.path.join(out, "rules.txt"), "w", encoding="utf-8") as f:
        f.write("# RULE SENTENCES (template's own words)\n" + "\n".join(rules))
        f.write("\n\n# INLINE NOTES (specimen hints: apply them to the text next to them, then delete them)\n")
        f.write("\n".join(f"{v}x {k}" for k, v in notes.most_common()))
        f.write("\n\n# PLACEHOLDERS TO FILL\n" + "\n".join(f"{v}x {k}" for k, v in holders.most_common()))

    print(f"{os.path.basename(path)}: {len(lines)} body elements, {len(d.sections)} sections, {len(d.tables)} tables, "
          f"{len(rules)} rule sentences, {sum(notes.values())} inline notes, {sum(holders.values())} placeholders")
    print("wrote structure.txt, rules.txt, sections.txt to", out)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    main(sys.argv[1], sys.argv[2])
