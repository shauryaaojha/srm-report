#!/usr/bin/env python3
"""Check a built report against the rules. Exit code 1 if anything FAILs.

usage: python -X utf8 check_report.py report.docx report.pdf [checks.json]

checks.json (all optional; take the values from the template's rules, not from memory):
  {"must_contain": ["names, reg numbers, course code"], "must_not_contain": ["SUPERVISOR", "old spelling"],
   "min_refs": 15, "min_pages": 30, "body_pt": 12, "layout": "layout.json", "font": "Times New Roman",
   "max_para_chars": 1250, "allow_first_person": false}
"""
import json, re, sys
from collections import Counter
import docx, fitz
from docx.oxml.ns import qn

LEFTOVER = re.compile(r"<<[^<>]*>>|<\s*Font[^<>]*>|<\s*[lI1]\.5 line spacing>|<Italic>|\.{6,}(?=\s*$)|STUDENT ?\d|REG NUM|\(GUIDE NAME|\(Designation, Department\)|"
                      r"SPECIALIZATION NAME|PROJECT TITLE|TITLE OF THE (?:REPORT|COMMUNITY)|Name of the [Ss]tudent|Dr\. ?ABCD|Specimen|Update this field|Error! |No table of (?:contents|figures) entries|Subtitle 1")
FIRST = re.compile(r"(?<![\w/])(I|my|My|me|mine|myself)(?![\w/.-])")
ROMAN_LABEL = re.compile(r"(Sprint|SPRINT|Phase|Part|Type|Review|Table|Class|Stage|Level|Tier|Unit|Module)\s+I\b")
results = []


def report(level, name, detail=""):
    results.append(level)
    print(f"{level:4s} {name}{': ' + detail if detail else ''}")


def spans(pdf):
    for p in pdf:
        for b in p.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                for sp in l["spans"]:
                    if sp["text"].strip():
                        yield sp


def main(docx_path, pdf_path, cfg):
    d = docx.Document(docx_path)
    body = d.element.body
    pdf = fitz.open(pdf_path)
    paras = [("".join(t.text or "" for t in p.iter(qn("w:t"))), p) for p in body.iter(qn("w:p"))]
    style = lambda p: (p.xpath("./w:pPr/w:pStyle/@w:val") or [""])[0]
    alltext = "\n".join(t for t, _ in paras)
    pdf_text = [p.get_text(sort=True) for p in pdf]   # sort: Word writes the footer first in the content stream
    pdf_all = "\n".join(pdf_text)

    hits = sorted({m.group(0) for m in LEFTOVER.finditer(pdf_all)})
    report("FAIL" if hits else "ok", "template placeholders and specimen text left", ", ".join(hits[:12]))

    flat = pdf_all.replace("\n", " ").lower()
    for s in cfg.get("must_contain", []):
        if s.lower() not in flat and s.lower() not in alltext.lower():
            report("FAIL", "missing from report", s)
    for s in cfg.get("must_not_contain", []):
        if s.lower() in alltext.lower() or s.lower() in pdf_all.lower():
            report("FAIL", "must not appear", s)

    if not cfg.get("allow_first_person"):
        fp = []
        for t, p in paras:
            if style(p) in ("RptCode", "RptRef") or t.strip() == "I":   # a lone I is a Sprint I table cell
                continue
            for m in FIRST.finditer(t):
                ctx = t[max(0, m.start() - 25):m.end() + 25]
                if m.group(1) == "I" and ROMAN_LABEL.search(ctx):
                    continue
                if re.match(r"\s*As an? ", t):   # user stories are written in the first person on purpose
                    continue
                fp.append(ctx.strip())
        report("WARN" if fp else "ok", "first person ('I', 'my') outside code: team reports say 'We'", " | ".join(fp[:6]))

    want = cfg.get("font", "Times New Roman").replace(" ", "").lower()
    fonts = Counter(); sizes = Counter()
    for sp in spans(pdf):
        fonts[sp["font"].split("+")[-1].split(",")[0].split("-")[0]] += len(sp["text"])
        if len(sp["text"].strip()) > 3:
            sizes[round(sp["size"] * 2) / 2] += len(sp["text"])
    other = {k: v for k, v in fonts.items() if want not in k.lower().replace("psmt", "").replace("ps", "") and not re.search("courier|symbol|wingdings", k, re.I)}
    total = sum(fonts.values()) or 1
    report("WARN" if sum(other.values()) > total * 0.002 else "ok", f"font is {cfg.get('font', 'Times New Roman')} (code may use Courier)", str(other) if other else "")
    top = sizes.most_common(1)[0][0] if sizes else 0
    if "body_pt" in cfg:
        report("ok" if top == cfg["body_pt"] else "FAIL", f"body text size is {cfg['body_pt']}pt", f"most common size {top}pt")

    # figures and tables
    fig, tab, prob = [], [], []
    kids = list(body.iterchildren())
    has_pic = lambda e: e.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip") is not None
    for i, e in enumerate(kids):
        if e.tag != qn("w:p"):
            continue
        t = "".join(x.text or "" for x in e.iter(qn("w:t")))
        m = re.match(r"(?:Fig|Figure)\.?\s*([A-Z0-9]+)\.(\d+)\s*:", t)
        if m and style(e) == "FigCaption":
            fig.append((m.group(1), int(m.group(2))))
            if i == 0 or not has_pic(kids[i - 1]):
                prob.append(f"{t[:30]} is not directly below a figure")
        m = re.match(r"Table\s*([A-Z0-9]+)\.(\d+)\s*:", t)
        if m and style(e) == "TabCaption":
            tab.append((m.group(1), int(m.group(2))))
            if i + 1 >= len(kids) or kids[i + 1].tag != qn("w:tbl"):
                prob.append(f"{t[:30]} is not directly above a table")
    for name, seq in (("Fig", fig), ("Table", tab)):
        last = {}
        for ch, n in seq:
            if n != last.get(ch, 0) + 1:
                prob.append(f"{name} {ch}.{n} breaks the numbering")
            last[ch] = n
    report("FAIL" if prob else "ok", f"captions: {len(fig)} figures below, {len(tab)} tables above, numbered per chapter", "; ".join(prob[:6]))

    # citations
    ref_at = max((i for i, (t, _) in enumerate(paras) if t.strip().upper() == "REFERENCES"), default=len(paras))
    ref_paras = [t for t, _ in paras[ref_at:] if re.match(r"\[\d+\]", t)]
    ref_nums = [int(re.match(r"\[(\d+)\]", t).group(1)) for t in ref_paras]
    order, seen = [], set()
    for t, p in paras[:ref_at]:
        if style(p) == "RptCode":
            continue
        for m in re.finditer(r"\[(\d+(?:\s*[,-]\s*\d+)*)\]", t):
            for part in m.group(1).split(","):
                part = part.strip()
                a, b = part.split("-") if "-" in part else (part, part)
                for n in range(int(a), int(b) + 1):
                    if n not in seen:
                        seen.add(n); order.append(n)
    prob = []
    if ref_nums != list(range(1, len(ref_nums) + 1)):
        prob.append("reference list is not numbered 1..N")
    if order != sorted(order):
        prob.append("citations do not first appear in order 1, 2, 3...")
    un = sorted(set(ref_nums) - seen); no = sorted(seen - set(ref_nums))
    if un:
        prob.append(f"never cited: {un[:10]}")
    if no:
        prob.append(f"cited but not listed: {no[:10]}")
    mn = cfg.get("min_refs", 0)
    if len(ref_nums) < mn:
        prob.append(f"only {len(ref_nums)} references, template asks for {mn}")
    report("FAIL" if prob else "ok", f"citations: {len(ref_nums)} references", "; ".join(prob))

    limit = cfg.get("max_para_chars", 1250)
    long_p = [t[:50] for t, p in paras if style(p) in ("RptBody", "") and len(t) > limit]
    report("WARN" if long_p else "ok", f"paragraphs near the 12-line limit (over {limit} characters)", " | ".join(long_p[:5]))

    # layout from the PDF
    def foot(pg):
        lines = [x.strip() for x in pg.get_text(sort=True).splitlines() if x.strip()]
        return lines[-1] if lines else ""

    chap = [i for i, t in enumerate(pdf_text) if re.search(r"^\s*CHAPTER\s+\d+\s*$", t, re.M) and "TABLE OF CONTENTS" not in t and not re.search(r"\.{5,}", t)]
    late = [i + 1 for i in chap if not re.match(r"\s*CHAPTER\s+\d+", pdf_text[i])]
    report("WARN" if late else "ok", f"{len(chap)} chapters each start at the top of a new page", f"physical pages {late}" if late else "")
    if chap:
        first = chap[0]
        seq = [foot(pdf[i]) for i in range(first, len(pdf))]
        bad = [i for i, s in enumerate(seq) if s != str(i + 1)]
        report("FAIL" if bad else "ok", "arabic page numbers start at 1 on Chapter 1 and run to the end", f"at position {bad[0] + 1} the footer reads {seq[bad[0]]!r}" if bad else "")
        rom = [i + 1 for i in range(1, first) if not re.fullmatch(r"[ivxlc]+", foot(pdf[i]))]
        report("WARN" if rom else "ok", "front pages show roman numerals", f"physical pages without one: {rom}" if rom else "")
    blank = [i + 1 for i, p in enumerate(pdf) if len(pdf_text[i].strip()) < 5 and not p.get_images()]
    report("FAIL" if blank else "ok", "no blank pages", str(blank) if blank else "")
    # One stray line on its own page: a fixed page overflowed, as an author list did in testing.
    stub = [i + 1 for i, p in enumerate(pdf) if 5 <= len(re.sub(r"\s+", "", pdf_text[i])) < 60 and not p.get_images()]
    report("FAIL" if stub else "ok", "no page holding only a stray line", str(stub) if stub else "")
    if cfg.get("min_pages"):
        report("ok" if len(pdf) >= cfg["min_pages"] else "FAIL", f"at least {cfg['min_pages']} pages", f"{len(pdf)} pages (the template excludes code and screenshots from this count)")

    if cfg.get("layout"):
        lay = json.load(open(cfg["layout"], encoding="utf-8-sig"))
        tabs = lay.get("tables") or []
        tabs = tabs if isinstance(tabs, list) else [tabs]
        sp = [t for t in tabs if t["start"] != t["end"]]
        report("FAIL" if sp else "ok", f"no table breaks across pages ({len(tabs)} tables)", "; ".join(f"pages {t['start']}-{t['end']}: {t['text'][:30]}" for t in sp[:6]))

    # Main headings open a page. A template that paginates with blank lines loses this silently when text moves.
    misplaced = []
    for h in ("ACKNOWLEDGEMENTS", "ABSTRACT", "TABLE OF CONTENTS", "LIST OF FIGURES", "LIST OF TABLES", "ABBREVIATIONS", "REFERENCES"):
        for i, t in enumerate(pdf_text):
            lines = [x.strip() for x in t.splitlines() if x.strip()]
            if h in lines and lines[0] != h:
                misplaced.append(f"{h} on physical page {i + 1}")
    report("FAIL" if misplaced else "ok", "main headings start at the top of a page", "; ".join(misplaced))

    # Fixed pages that must stay whole. Text can overflow silently, so look at the rendered pages.
    pairs = [["BONAFIDE CERTIFICATE", "EXAMINER 2"]] + cfg.get("same_page", [])
    for a, b in pairs:
        on_a = [i for i, t in enumerate(pdf_text) if a in t]; on_b = [i for i, t in enumerate(pdf_text) if b in t]
        if on_a or on_b:
            ok = bool(set(on_a) & set(on_b))
            report("ok" if ok else "FAIL", f"'{a}' and '{b}' are on one page", "" if ok else f"physical pages {[x + 1 for x in on_a]} and {[x + 1 for x in on_b]}")
    if not pdf[0].get_images():
        report("FAIL", "cover has its logo", "no image on page 1")
    else:
        report("ok", "cover has its logo")
    in_cells = [s for s in body.iter(qn("w:sectPr")) if any(a.tag == qn("w:tc") for a in s.iterancestors())]
    report("FAIL" if in_cells else "ok", "no stray section breaks inside tables", str(len(in_cells)) if in_cells else "")
    print(f"\n{results.count('FAIL')} fail, {results.count('WARN')} warn, {results.count('ok')} ok, {len(pdf)} pages")
    return 1 if "FAIL" in results else 0


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit(__doc__)
    cfg = json.load(open(sys.argv[3], encoding="utf-8")) if len(sys.argv) > 3 else {}
    sys.exit(main(sys.argv[1], sys.argv[2], cfg))
