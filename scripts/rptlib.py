"""Build an SRM report by filling a COPY of the template, so fixed pages stay exactly as issued.

Typical use (see assets/examples/project_21cse734p.py):
    r = Report(template, style={...})
    r.strip_notes(); r.replace(old, new); r.set_para(prefix, text); r.signature_table(...)
    r.cut_from(r.find("ABSTRACT"))          # drops the specimen body, keeps the pages before it
    r.prepare(all_nodes, refs, img_dir)
    r.begin_front(); r.add_nodes(...); r.add_toc(); r.begin_body(); r.add_nodes(chapters); r.add_references()
    r.save(out)
Then run word_pass.ps1 (fills TOC and lists, exports PDF) and check_report.py.

Content nodes are plain dicts with key "t": chapter, appendix, front, heading, h2, h3, p, quote, num, bullets,
fig, table, code, abbrev, pagebreak. Inline text: [@key] citation, {fig:id} / {tab:id} cross reference, **bold**.
"""
import copy, io, os, re
import docx
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt
from PIL import Image

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XS = "{http://www.w3.org/XML/1998/namespace}space"
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
WP_DOCPR = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr"
NOTE_RE = re.compile(r"(?<!<)<(?!<)[^<>\n]{2,80}>(?!>)")

DEFAULT_STYLE = dict(font="Times New Roman", body=12, line=1.5, after=12, h1=16, h2=16, h3=12, h3_caps=True,
                     front_h=16, caption=12, table=12, table_small=10.5, code=8, left=1.0, right=0.5, top=1.0,
                     bottom=0.9, footer=0.42, page_w=8.27, page_h=11.69, fig_label="Fig", tab_label="Table")


def ptext(el):
    return "".join(x.text or "" for x in el.iter(qn("w:t")))


def _mk_t(text):
    t = OxmlElement("w:t"); t.text = text; t.set(XS, "preserve"); return t


def set_text(p_el, text):
    """Replace a paragraph's text and keep its first run's formatting. Removes stale runs, which is what
    caused the 'DepartmentAssistant Professor,' bug when only some runs were edited."""
    runs = [r for r in p_el.iter(qn("w:r")) if r.find(qn("w:t")) is not None or r.find(qn("w:tab")) is not None]
    if not runs:
        r = OxmlElement("w:r"); r.append(_mk_t(text)); p_el.append(r); return
    first = runs[0]
    for r in runs[1:]:
        r.getparent().remove(r)
    for x in list(first):
        if x.tag != qn("w:rPr"):
            first.remove(x)
    for li, line in enumerate(text.split("\n")):
        if li:
            first.append(OxmlElement("w:br"))
        for i, part in enumerate(line.split("\t")):
            if i:
                first.append(OxmlElement("w:tab"))
            if part:
                first.append(_mk_t(part))


def add_run_like(p_el, text, like, bold=None):
    r = OxmlElement("w:r")
    rpr = like.find(qn("w:rPr")) if like is not None else None
    rpr = copy.deepcopy(rpr) if rpr is not None else OxmlElement("w:rPr")
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            for x in rpr.findall(qn(tag)):
                rpr.remove(x)
        b = OxmlElement("w:b"); b.set(qn("w:val"), "1" if bold else "0"); rpr.append(b)
    r.append(rpr)
    for i, part in enumerate(text.split("\t")):
        if i:
            r.append(OxmlElement("w:tab"))
        if part:
            r.append(_mk_t(part))
    p_el.append(r)
    return r


def _replace_across_runs(p_el, a, b):
    """Replace a with b even when Word split a across runs. The new text takes the first touched run's format."""
    nodes = list(p_el.iter(qn("w:t")))
    full = "".join(t.text or "" for t in nodes)
    n = 0
    start = full.find(a)
    while start != -1:
        end = start + len(a); pos = 0; placed = False
        for t in nodes:
            s = t.text or ""; ts, te = pos, pos + len(s); pos = te
            if te <= start or ts >= end:
                continue
            left = s[:max(0, start - ts)]
            right = s[end - ts:] if te > end else ""
            t.text = left + (b if not placed else "") + right
            t.set(XS, "preserve"); placed = True
        n += 1
        full = "".join(t.text or "" for t in nodes)
        start = full.find(a, start + len(b))
    return n


class Report:
    def __init__(self, template, style=None):
        self.path = template
        self.d = docx.Document(template)
        self.body = self.d.element.body
        self.s = {**DEFAULT_STYLE, **(style or {})}
        self.fig_no, self.tab_no, self.cite_no = {}, {}, {}
        self.refs = {}
        self._styles_ready = False
        self._img_base = "."

    # ---------------- reading the template ----------------
    def els(self):
        return [e for e in self.body.iterchildren() if e.tag != qn("w:sectPr")]

    def all_paras(self):
        return list(self.body.iter(qn("w:p")))

    def find(self, prefix, nth=0):
        """First top-level paragraph starting with prefix. Raises if missing, so a different template fails loudly."""
        seen = 0
        for e in self.els():
            if e.tag == qn("w:p") and ptext(e).strip().startswith(prefix):
                if seen == nth:
                    return e
                seen += 1
        raise KeyError(f"not found in template: {prefix!r} (nth={nth})")

    def tables_with(self, text):
        return [t for t in self.body.iter(qn("w:tbl")) if text in ptext(t)]

    # ---------------- editing kept pages ----------------
    def apply_cover_notes(self):
        """Carry out the cover's notes, then delete them. In SRM templates a line such as <Font Size 16> <Italic> is an
        instruction for the text ABOVE it, back to the previous note, and <Font Style ... Bold> covers the whole cover.
        Call this before filling names, so cloned lines inherit the size. Returns the number of notes applied."""
        kids = self.els(); group = []; all_bold = False; applied = 0

        def fmt_of(notes):
            joined = " ".join(notes)
            m = re.search(r"Font Size\s*(\d+)", joined)
            return dict(size=int(m.group(1)) if m else None, italic="italic" in joined.lower(),
                        wide=bool(re.search(r"[1lI]\.5 line", joined)))

        def apply(paras, f):
            for k, p in enumerate(paras):
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
                for tag in ("w:ind", "w:spacing", "w:jc", "w:tabs"):
                    for x in ppr.findall(qn(tag)):
                        ppr.remove(x)
                after = 120 if k == len(paras) - 1 else 0
                line = 360 if f["wide"] else 276
                ppr.append(parse_xml(f'<w:spacing xmlns:w="{W}" w:before="0" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'))
                ppr.append(parse_xml(f'<w:ind xmlns:w="{W}" w:left="0" w:right="0" w:firstLine="0"/>'))
                ppr.append(parse_xml(f'<w:jc xmlns:w="{W}" w:val="center"/>'))
                for r in p.iter(qn("w:r")):
                    rpr = r.find(qn("w:rPr"))
                    if rpr is None:
                        rpr = OxmlElement("w:rPr"); r.insert(0, rpr)
                    # Negative character spacing in the template squeezes words together ("requirementsfor").
                    for tag in ("w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:spacing", "w:position"):
                        for x in rpr.findall(qn(tag)):
                            rpr.remove(x)
                    font = self.s["font"]
                    rpr.append(parse_xml(f'<w:rFonts xmlns:w="{W}" w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'))
                    val = "1" if all_bold else "0"
                    rpr.append(parse_xml(f'<w:b xmlns:w="{W}" w:val="{val}"/>'))
                    rpr.append(parse_xml(f'<w:bCs xmlns:w="{W}" w:val="{val}"/>'))
                    if f["italic"]:
                        rpr.append(OxmlElement("w:i"))
                    if f["size"]:
                        for tag in ("w:sz", "w:szCs"):
                            for x in rpr.findall(qn(tag)):
                                rpr.remove(x)
                        rpr.append(parse_xml(f'<w:sz xmlns:w="{W}" w:val="{f["size"] * 2}"/>'))
                        rpr.append(parse_xml(f'<w:szCs xmlns:w="{W}" w:val="{f["size"] * 2}"/>'))

        for e in kids:
            if e.tag != qn("w:p"):
                break
            has_sect = e.find(".//" + qn("w:sectPr")) is not None
            has_pic = e.find(".//" + A_BLIP) is not None
            t = ptext(e); notes = NOTE_RE.findall(t); clean = NOTE_RE.sub("", t).strip()
            if clean.startswith("(A typical Specimen"):
                self.body.remove(e)
            elif notes and not clean and not has_pic:   # the logo can sit inside a note's paragraph
                if "bold" in " ".join(notes).lower() and not group:
                    all_bold = True
                elif group:
                    apply(group, fmt_of(notes)); applied += 1
                group = []
                self.body.remove(e)
            elif notes:
                for m in set(notes):
                    _replace_across_runs(e, m, "")
                apply(group + [e], fmt_of(notes)); applied += 1
                group = []
            elif not clean and not has_pic and not has_sect:
                self.body.remove(e)
            elif clean or has_pic:
                group.append(e)
            if has_sect:
                break
        return applied

    def strip_notes(self, keep_lines=False):
        """Delete remaining hints like <Font Size 18>, along with paragraphs that hold nothing else. Keeping those as
        empty lines (keep_lines=True) pushes a full page such as the bonafide onto a second page."""
        removed = 0
        for p in self.all_paras():
            t = ptext(p)
            if not NOTE_RE.search(t):
                continue
            only_note = not NOTE_RE.sub("", t).strip()
            if only_note and not keep_lines and p.find(".//" + A_BLIP) is None and p.find(".//" + qn("w:sectPr")) is None:
                p.getparent().remove(p); removed += 1; continue
            for m in set(NOTE_RE.findall(t)):
                _replace_across_runs(p, m, "")
            removed += 1
        for p in self.all_paras():
            if ptext(p).strip().startswith("(A typical Specimen"):
                p.getparent().remove(p); removed += 1
        return removed

    def replace(self, a, b, expect=None):
        n = sum(_replace_across_runs(p, a, b) for p in self.all_paras())
        if expect is not None and n != expect:
            raise ValueError(f"replace {a!r}: expected {expect}, got {n}")
        if n == 0:
            raise KeyError(f"replace: {a!r} not found")
        return n

    def set_para(self, prefix, text, nth=0):
        seen = 0
        for p in self.all_paras():
            if ptext(p).strip().startswith(prefix):
                if seen == nth:
                    set_text(p, text); return p
                seen += 1
        raise KeyError(f"set_para: {prefix!r} not found")

    def fill_field(self, label, value, tab_pos=3780, left=3960, hanging=3060):
        """Fill a 'Label<tab>:' row, wrapping long values under the value column."""
        for p in self.all_paras():
            if ptext(p).strip().startswith(label):
                like = next(r for r in p.iter(qn("w:r")) if r.find(qn("w:t")) is not None)
                like = copy.deepcopy(like)
                set_text(p, label)
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
                for x in ppr.findall(qn("w:tabs")) + ppr.findall(qn("w:ind")):
                    ppr.remove(x)
                ppr.append(parse_xml(f'<w:tabs xmlns:w="{W}"><w:tab w:val="left" w:pos="{tab_pos}"/></w:tabs>'))
                ppr.append(parse_xml(f'<w:ind xmlns:w="{W}" w:left="{left}" w:hanging="{hanging}"/>'))
                add_run_like(p, "\t:", like, bold=True)
                add_run_like(p, " " + value, like, bold=False)
                return p
        raise KeyError(f"fill_field: {label!r} not found")

    def trim_blank_before(self, el, keep=0):
        """Remove empty lines directly above el, leaving `keep` of them. Stops at an image, section break or page break."""
        blanks = []
        prev = el.getprevious()
        while prev is not None and prev.tag == qn("w:p") and not ptext(prev).strip() \
                and prev.find(".//" + qn("w:drawing")) is None and prev.find(".//" + qn("w:sectPr")) is None \
                and not any(b.get(qn("w:type")) == "page" for b in prev.iter(qn("w:br"))):
            blanks.append(prev); prev = prev.getprevious()
        for b in blanks[keep:]:
            b.getparent().remove(b)

    def start_new_page(self, el):
        """Give a heading a real page break. Templates often reach the next page with blank lines alone, so anything
        that removes those lines or lengthens the page above needs this."""
        self.trim_blank_before(el)
        ppr = el.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr"); el.insert(0, ppr)
        if ppr.find(qn("w:pageBreakBefore")) is None:
            ppr.append(OxmlElement("w:pageBreakBefore"))

    def tidy_paragraphs(self, start_prefix, end_prefix):
        """Clean template padding between two markers: repeated and leading spaces, and first-line offsets.
        Left and right indents stay, because they set the width the page's text lines up on."""
        a = self.find(start_prefix); b = self.find(end_prefix); on = False
        for e in self.els():
            if e is a:
                on = True; continue
            if e is b:
                break
            if not on or e.tag != qn("w:p"):
                continue
            ind = e.find(qn("w:pPr") + "/" + qn("w:ind")) if e.find(qn("w:pPr")) is not None else None
            if ind is not None:
                for att in ("w:firstLine", "w:hanging"):
                    ind.attrib.pop(qn(att), None)
            space = True   # true at the start, so leading spaces go too
            for t in e.iter(qn("w:t")):
                out = []
                for ch in t.text or "":
                    is_sp = ch.isspace()
                    if not (is_sp and space):
                        out.append(" " if is_sp else ch)
                    space = is_sp
                t.text = "".join(out); t.set(XS, "preserve")

    def clone_after(self, model, text, after=None):
        """Copy paragraph `model`, set its text, place it after `after` (default: after model)."""
        c = copy.deepcopy(model)
        for s in list(c.iter(qn("w:sectPr"))):
            s.getparent().remove(s)   # a cloned sectPr silently starts a new page
        set_text(c, text)
        (after if after is not None else model).addnext(c)
        return c

    def bold_names(self, start_prefix, end_prefix, names):
        """Bold each name in the paragraphs between two markers (acknowledgements). Raises if a name is absent."""
        a = self.find(start_prefix); b = self.find(end_prefix)
        rx = re.compile("(" + "|".join(re.escape(n) for n in sorted(names, key=len, reverse=True)) + ")")
        found = set(); on = False
        for e in self.els():
            if e is a:
                on = True; continue
            if e is b:
                break
            if not on or e.tag != qn("w:p"):
                continue
            t = ptext(e)
            if not rx.search(t):
                continue
            t = re.sub(r"\s{2,}", " ", t).strip()   # the template pads some lines with spaces
            runs = e.findall(qn("w:r")); like = runs[0] if runs else None
            for r in runs:
                e.remove(r)
            for part in rx.split(t):
                if part:
                    is_name = bool(rx.fullmatch(part))
                    if is_name:
                        found.add(part)
                    add_run_like(e, part, like, bold=is_name)
        missing = set(names) - found
        if missing:
            raise KeyError(f"names not found between markers: {sorted(missing)}")

    def signature_table(self, table_el, blocks):
        """Rebuild a bonafide signature table with one equal column per signer.
        blocks = [[('SIGNATURE', True), ('Dr. X', False), ('GUIDE', True), ...], ...]
        Pad blocks to the same line count so the columns line up."""
        tr = table_el.find(qn("w:tr")); tcs = tr.findall(qn("w:tc"))
        model = next(c for c in tcs if ptext(c).strip())
        para_model = next(p for p in model.iter(qn("w:p")) if ptext(p).strip())
        run_model = copy.deepcopy(next(r for r in para_model.iter(qn("w:r")) if r.find(qn("w:t")) is not None))
        grid = table_el.find(qn("w:tblGrid")); gcs = grid.findall(qn("w:gridCol"))
        total = sum(int(g.get(qn("w:w"))) for g in gcs)
        col = total // len(blocks)
        for g in gcs:
            grid.remove(g)
        for _ in blocks:
            grid.append(parse_xml(f'<w:gridCol xmlns:w="{W}" w:w="{col}"/>'))
        new = []
        for lines in blocks:
            tc = OxmlElement("w:tc")
            tc.append(parse_xml(f'<w:tcPr xmlns:w="{W}"><w:tcW w:w="{col}" w:type="dxa"/><w:tcMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>'))
            for text, bold in lines:
                p = OxmlElement("w:p")
                p.append(parse_xml(f'<w:pPr xmlns:w="{W}"><w:spacing w:before="0" w:after="0"/><w:jc w:val="center"/></w:pPr>'))
                add_run_like(p, text, run_model, bold=bold)
                tc.append(p)
            new.append(tc)
        for c in tcs:
            tr.remove(c)
        for c in new:
            tr.append(c)

    def copy_block(self, src_path, start_prefix, end_prefix, anchor, style_prefix="Tpl"):
        """Copy template elements [start..end] in front of `anchor`, for a report NOT built from the template (for example
        one the user hand-edited). Brings images and list numbering along, and renames styles so they keep the template's look."""
        t = docx.Document(src_path)
        tb = [e for e in t.element.body.iterchildren() if e.tag != qn("w:sectPr")]
        i = next(k for k, e in enumerate(tb) if ptext(e).strip().startswith(start_prefix))
        j = next(k for k in range(i, len(tb)) if ptext(tb[k]).strip().startswith(end_prefix))
        have = {s.style_id for s in self.d.styles}
        tstyles = {s.style_id: s for s in t.styles}
        smap = {}

        def bring(sid):
            if sid in smap:
                return smap[sid]
            new_id = style_prefix + sid; smap[sid] = new_id
            if new_id in have or sid not in tstyles:
                return new_id
            el = copy.deepcopy(tstyles[sid].element); el.set(qn("w:styleId"), new_id)
            el.find(qn("w:name")).set(qn("w:val"), style_prefix + " " + tstyles[sid].name)
            el.attrib.pop(qn("w:default"), None)
            for tag in ("w:link", "w:next"):
                for x in el.findall(qn(tag)):
                    el.remove(x)
            based = el.find(qn("w:basedOn"))
            if based is not None:
                based.set(qn("w:val"), bring(based.get(qn("w:val"))))
            self.d.styles.element.append(el); have.add(new_id)
            return new_id

        normal = next((s.style_id for s in t.styles if s.name == "Normal"), "Normal")
        out = []; dp = 9000; num_map = {}
        tnum = t.part.numbering_part.element
        for e in tb[i:j + 1]:
            c = copy.deepcopy(e)
            for s in list(c.iter(qn("w:sectPr"))):
                s.getparent().remove(s)
            for ps in c.iter(qn("w:pStyle")):
                ps.set(qn("w:val"), bring(ps.get(qn("w:val"))))
            for p in c.iter(qn("w:p")):
                ppr = p.find(qn("w:pPr"))
                if ppr is None:
                    ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
                if ppr.find(qn("w:pStyle")) is None:
                    ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), bring(normal)); ppr.insert(0, ps)
            for n in c.iter(qn("w:numId")):
                old = n.get(qn("w:val"))
                if old not in num_map:
                    num_map[old] = self._bring_numbering(tnum, old)
                n.set(qn("w:val"), num_map[old])
            for b in c.iter(A_BLIP):
                blob = t.part.rels[b.get(qn("r:embed"))].target_part.blob
                rid, _ = self.d.part.get_or_add_image(io.BytesIO(blob)); b.set(qn("r:embed"), rid)
            for x in c.iter(WP_DOCPR):
                dp += 1; x.set("id", str(dp))
            anchor.addprevious(c); out.append(c)
        return out

    def _bring_numbering(self, tnum, old):
        dnum = self.d.part.numbering_part.element
        tn = next(n for n in tnum.findall(qn("w:num")) if n.get(qn("w:numId")) == old)
        aid = tn.find(qn("w:abstractNumId")).get(qn("w:val"))
        ta = next(a for a in tnum.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == aid)
        new_a = str(max([int(a.get(qn("w:abstractNumId"))) for a in dnum.findall(qn("w:abstractNum"))] + [0]) + 1)
        new_n = str(max([int(n.get(qn("w:numId"))) for n in dnum.findall(qn("w:num"))] + [0]) + 1)
        na = copy.deepcopy(ta); na.set(qn("w:abstractNumId"), new_a)
        for x in na.findall(qn("w:nsid")) + na.findall(qn("w:tmpl")):
            na.remove(x)
        first = dnum.find(qn("w:num"))
        if first is not None:
            first.addprevious(na)
        else:
            dnum.append(na)
        dnum.append(parse_xml(f'<w:num xmlns:w="{W}" w:numId="{new_n}"><w:abstractNumId w:val="{new_a}"/></w:num>'))
        return new_n

    # ---------------- sections ----------------
    def cut_from(self, el):
        """Delete el and everything after it (the specimen). The kept pages keep their own section settings."""
        kids = self.els()
        i = kids.index(el)
        # Empty lines and page breaks in front of the cut would spill onto a blank page.
        while i > 0 and kids[i - 1].tag == qn("w:p") and not ptext(kids[i - 1]).strip() \
                and kids[i - 1].find(".//" + A_BLIP) is None and kids[i - 1].find(".//" + qn("w:sectPr")) is None:
            i -= 1
        prev_has = i > 0 and kids[i - 1].find(".//" + qn("w:sectPr")) is not None
        next_sect = None
        for e in kids[i:]:
            next_sect = e.find(".//" + qn("w:sectPr"))
            if next_sect is not None:
                next_sect = copy.deepcopy(next_sect)
                break
        for e in kids[i:]:
            self.body.remove(e)
        if not prev_has:
            # The kept pages before the cut need a section of their own, or the final sectPr would restyle them.
            carrier = parse_xml(f'<w:p xmlns:w="{W}"><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>')
            sect = next_sect if next_sect is not None else copy.deepcopy(self.body.find(qn("w:sectPr")))
            carrier.find(qn("w:pPr")).append(sect)
            self.body.find(qn("w:sectPr")).addprevious(carrier)
        return self

    def _footer(self, sec, fmt, start=None, number=True):
        s = self.s
        sec.footer.is_linked_to_previous = False
        ft = sec.footer
        for p in list(ft.paragraphs)[1:]:
            p._p.getparent().remove(p._p)
        p = ft.paragraphs[0] if ft.paragraphs else ft.add_paragraph()
        for c in list(p._p):
            if c.tag != qn("w:pPr"):
                p._p.remove(c)
        ppr = p._p.get_or_add_pPr()
        for x in ppr.findall(qn("w:pBdr")) + ppr.findall(qn("w:jc")):
            ppr.remove(x)   # number only, never a rule line
        ppr.append(parse_xml(f'<w:jc xmlns:w="{W}" w:val="center"/>'))
        if number:
            for part in ('<w:fldChar w:fldCharType="begin"/>', '<w:instrText xml:space="preserve"> PAGE </w:instrText>', '<w:fldChar w:fldCharType="end"/>'):
                p._p.append(parse_xml(f'<w:r xmlns:w="{W}"><w:rPr><w:rFonts w:ascii="{s["font"]}" w:hAnsi="{s["font"]}" w:cs="{s["font"]}"/><w:sz w:val="{int(s["body"] * 2)}"/></w:rPr>{part}</w:r>'))
        sp = sec._sectPr
        for x in sp.findall(qn("w:pgNumType")):
            sp.remove(x)
        attrs = f'w:fmt="{fmt}"' + (f' w:start="{start}"' if start is not None else "")
        sp.find(qn("w:pgMar")).addnext(parse_xml(f'<w:pgNumType xmlns:w="{W}" {attrs}/>'))

    def number_kept_sections(self, cover_sections=1, fmt="lowerRoman"):
        """Footer numbers on the kept front pages. The cover counts as page i but shows no number.
        Call after begin_front, which claims the last section."""
        secs = list(self.d.sections)[:-1]
        for i, sec in enumerate(secs):
            show = i >= cover_sections
            self._footer(sec, fmt, start=1 if i == 0 else None, number=show)
            if show:
                sec.footer_distance = Inches(self.s["footer"])
                sec.bottom_margin = max(sec.bottom_margin, Inches(0.62))

    def _apply_page(self, sec):
        s = self.s
        sec.page_width, sec.page_height = Inches(s["page_w"]), Inches(s["page_h"])
        sec.left_margin, sec.right_margin = Inches(s["left"]), Inches(s["right"])
        sec.top_margin, sec.bottom_margin = Inches(s["top"]), Inches(s["bottom"])
        sec.footer_distance = Inches(s["footer"]); sec.header_distance = Inches(0.3)
        for x in sec._sectPr.findall(qn("w:cols")):
            x.attrib.pop(qn("w:num"), None)

    def begin_front(self, fmt="lowerRoman"):
        """Start the generated front section (abstract, contents, lists). Uses the document's last section."""
        self._ensure_styles()
        sec = self.d.sections[-1]
        self._apply_page(sec); self._footer(sec, fmt)

    def begin_body(self, fmt="decimal", start=1):
        """Start the arabic-numbered section. Chapter 1 is page 1."""
        self._ensure_styles()
        sec = self.d.add_section(WD_SECTION.NEW_PAGE)
        self._apply_page(sec); self._footer(sec, fmt, start=start)

    # ---------------- styles ----------------
    def _para_style(self, sid, size, bold=False, italic=False, jc="both", line=None, before=0, after=None, outline=None, keep_next=False, page_break=False):
        s = self.s
        line = int((line or s["line"]) * 240)
        after = int((s["after"] if after is None else after) * 20)
        for st in list(self.d.styles):
            if st.style_id == sid:
                self.d.styles.element.remove(st.element)
        flags = ("<w:keepNext/>" if keep_next else "") + ("<w:pageBreakBefore/>" if page_break else "")
        ol = f'<w:outlineLvl w:val="{outline}"/>' if outline is not None else ""
        bi = lambda on: "1" if on else "0"
        self.d.styles.element.append(parse_xml(
            f'<w:style xmlns:w="{W}" w:type="paragraph" w:customStyle="1" w:styleId="{sid}"><w:name w:val="{sid}"/><w:qFormat/>'
            f'<w:pPr>{flags}<w:widowControl/><w:spacing w:before="{int(before * 20)}" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'
            f'<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="{jc}"/>{ol}</w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{s["font"]}" w:eastAsia="{s["font"]}" w:hAnsi="{s["font"]}" w:cs="{s["font"]}"/>'
            f'<w:b w:val="{bi(bold)}"/><w:bCs w:val="{bi(bold)}"/><w:i w:val="{bi(italic)}"/><w:color w:val="000000"/>'
            f'<w:sz w:val="{int(size * 2)}"/><w:szCs w:val="{int(size * 2)}"/></w:rPr></w:style>'))

    def _ensure_styles(self):
        if self._styles_ready:
            return
        s = self.s
        # Own style names: template pages use Heading 1-5 on the cover, so a plain \o TOC would list cover lines.
        self._para_style("RptChapter", s["h1"], bold=True, jc="center", after=24, outline=0, keep_next=True, page_break=True)
        self._para_style("RptFront", s["front_h"], bold=True, jc="center", after=18, outline=0, keep_next=True, page_break=True)
        self._para_style("RptH2", s["h2"], bold=True, jc="left", before=12, outline=1, keep_next=True, after=12)
        self._para_style("RptH3", s["h3"], bold=True, jc="left", before=6, outline=2, keep_next=True, after=10)
        self._para_style("RptBody", s["body"])
        self._para_style("RptCell", s["table"], jc="left", line=1.05, after=0)
        self._para_style("RptRef", s["body"], jc="left", line=1.25, after=8)
        self._para_style("RptCode", s["code"], jc="left", line=1.0, after=0)
        self._para_style("RptFigure", s["body"], jc="center", line=1.0, before=6, after=6, keep_next=True)
        self._para_style("FigCaption", s["caption"], jc="center", line=1.15, before=3, after=12)
        self._para_style("TabCaption", s["caption"], jc="center", line=1.15, before=3, after=12, keep_next=True)
        self._para_style("CodeCaption", s["caption"] - 1, italic=True, jc="center", line=1.15, before=3, after=12)
        self._toc_styles()
        self._styles_ready = True

    def _toc_styles(self):
        """Chapter lines bold, sub-topics regular. Set in the styles so it survives 'Update field' in Word."""
        s = self.s
        tw = int((s["page_w"] - s["left"] - s["right"]) * 1440)
        for level, name, bold, left in ((1, "toc 1", True, 0), (2, "toc 2", False, 240), (3, "toc 3", False, 480), (4, "toc 4", False, 0)):
            sid = f"TOC{level}"
            for st in list(self.d.styles):
                if st.style_id == sid or st.name.lower() == name:
                    self.d.styles.element.remove(st.element)
            b = "1" if bold else "0"
            self.d.styles.element.append(parse_xml(
                f'<w:style xmlns:w="{W}" w:type="paragraph" w:styleId="{sid}"><w:name w:val="{name}"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/>'
                f'<w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{tw}"/></w:tabs><w:spacing w:before="0" w:after="80" w:line="276" w:lineRule="auto"/><w:ind w:left="{left}" w:right="360"/></w:pPr>'
                f'<w:rPr><w:rFonts w:ascii="{s["font"]}" w:eastAsia="{s["font"]}" w:hAnsi="{s["font"]}" w:cs="{s["font"]}"/><w:b w:val="{b}"/><w:bCs w:val="{b}"/><w:sz w:val="{int(s["body"] * 2)}"/></w:rPr></w:style>'))

    # ---------------- numbering ----------------
    def prepare(self, nodes, refs, img_dir="."):
        """First pass over all nodes: number figures and tables per chapter, citations in order of first use."""
        self.refs = refs; self._img_base = img_dir
        ch, f, t = "", 0, 0
        for n in nodes:
            k = n.get("t")
            if k == "chapter":
                ch, f, t = str(n["num"]), 0, 0
            if k == "appendix":
                ch, f, t = str(n["letter"]), 0, 0
            if k == "fig":
                f += 1; self.fig_no[n["id"]] = f"{ch}.{f}"
            if k == "table":
                t += 1; self.tab_no[n["id"]] = f"{ch}.{t}"
            texts = [n.get("text")] + list(n.get("items") or []) + list(n.get("head") or []) + [c for r in (n.get("rows") or []) for c in r]
            for s in texts:
                if s:
                    self._scan(str(s))
        unused = [k for k in refs if k not in self.cite_no]
        if unused:
            raise ValueError(f"references never cited: {unused}")

    def _scan(self, s):
        for m in re.finditer(r"\[@([^\]]+)\]", s):
            for k in m.group(1).split(","):
                k = k.strip().lstrip("@")
                if k not in self.refs:
                    raise KeyError(f"citation with no reference: {k}")
                self.cite_no.setdefault(k, len(self.cite_no) + 1)

    def _fmt_cite(self, keys):
        nums = sorted({self.cite_no[k.strip().lstrip("@")] for k in keys}); out = []; i = 0
        while i < len(nums):
            j = i
            while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
                j += 1
            if j - i >= 2:
                out.append(f"{nums[i]}-{nums[j]}")
            else:
                out.extend(str(x) for x in nums[i:j + 1])
            i = j + 1
        return "[" + ",".join(out) + "]"

    def _resolve(self, s):
        s = re.sub(r"\[@([^\]]+)\]", lambda m: self._fmt_cite(m.group(1).split(",")), str(s))

        def fig(m):
            return f'{self.s["fig_label"]} {self.fig_no[m.group(1)]}'

        def tab(m):
            return f'{self.s["tab_label"]} {self.tab_no[m.group(1)]}'
        return re.sub(r"\{tab:([\w-]+)\}", tab, re.sub(r"\{fig:([\w-]+)\}", fig, s))

    # ---------------- generated content ----------------
    def _runs(self, p, text, bold=False, italic=False, size=None):
        for seg in re.split(r"(\*\*[^*]+\*\*)", self._resolve(text)):
            if not seg:
                continue
            b = seg.startswith("**") and seg.endswith("**")
            r = p.add_run(seg[2:-2] if b else seg)
            if b or bold:
                r.bold = True
            if italic:
                r.italic = True
            if size:
                r.font.size = Pt(size)

    def _p(self, style, text=None, bold=False, italic=False, size=None):
        p = self.d.add_paragraph(style=self.d.styles[style])
        if text is not None:
            self._runs(p, text, bold, italic, size)
        return p

    def _field(self, instr, style="RptBody"):
        p = self.d.add_paragraph(style=self.d.styles[style])
        for part in ('<w:fldChar w:fldCharType="begin" w:dirty="true"/>', f'<w:instrText xml:space="preserve"> {instr} </w:instrText>',
                     '<w:fldChar w:fldCharType="separate"/>', '<w:t>Update this field in Word</w:t>', '<w:fldChar w:fldCharType="end"/>'):
            p._p.append(parse_xml(f'<w:r xmlns:w="{W}">{part}</w:r>'))
        return p

    def add_toc(self, title="TABLE OF CONTENTS"):
        self._p("RptFront", title)
        self._field('TOC \\h \\z \\t "RptFront,1,RptChapter,1,RptH2,2,RptH3,3"')

    def add_list_of_figures(self, title="LIST OF FIGURES"):
        self._p("RptFront", title); self._field('TOC \\h \\z \\t "FigCaption,4"')

    def add_list_of_tables(self, title="LIST OF TABLES"):
        self._p("RptFront", title); self._field('TOC \\h \\z \\t "TabCaption,4"')

    def _borders(self, tbl_el, on=True):
        pr = tbl_el.find(qn("w:tblPr"))
        for x in pr.findall(qn("w:tblBorders")):
            pr.remove(x)
        val = "single" if on else "nil"
        pr.append(parse_xml(f'<w:tblBorders xmlns:w="{W}">' + "".join(f'<w:{e} w:val="{val}" w:sz="4" w:space="0" w:color="808080"/>' for e in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"))

    def _table(self, widths_in, rows, head=True, size=None, borders=True, shade="D9E2F3", cell_style="RptCell", bold_first=False):
        text_w = self.s["page_w"] - self.s["left"] - self.s["right"] - 0.03
        scale = min(1, text_w / sum(widths_in))
        tw = [int(w * 1440 * scale) for w in widths_in]
        t = self.d.add_table(rows=len(rows), cols=len(widths_in))
        el = t._tbl; pr = el.find(qn("w:tblPr"))
        for tag in ("w:tblStyle", "w:tblW", "w:jc", "w:tblLayout"):
            for x in pr.findall(qn(tag)):
                pr.remove(x)
        pr.append(parse_xml(f'<w:tblW xmlns:w="{W}" w:w="{sum(tw)}" w:type="dxa"/>'))
        pr.append(parse_xml(f'<w:jc xmlns:w="{W}" w:val="center"/>'))
        self._borders(el, borders)
        pr.append(parse_xml(f'<w:tblLayout xmlns:w="{W}" w:type="fixed"/>'))
        pr.append(parse_xml(f'<w:tblCellMar xmlns:w="{W}"><w:top w:w="50" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar>'))
        for g, w in zip(el.find(qn("w:tblGrid")).findall(qn("w:gridCol")), tw):
            g.set(qn("w:w"), str(w))
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                c = t.cell(ri, ci); c.width = Emu(tw[ci] * 635)
                p = c.paragraphs[0]; p.style = self.d.styles[cell_style]
                is_head = head and ri == 0
                self._runs(p, val, bold=is_head or (bold_first and ci == 0), size=size)
                if is_head and shade:
                    c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd xmlns:w="{W}" w:val="clear" w:color="auto" w:fill="{shade}"/>'))
        return t

    def add_nodes(self, nodes):
        s = self.s
        for n in nodes:
            k = n["t"]
            if k == "chapter":
                p = self._p("RptChapter", f'CHAPTER {n["num"]}'); p.add_run().add_break(); p.add_run(n["title"])
            elif k == "appendix":
                p = self._p("RptChapter", f'APPENDIX {n["letter"]}'); p.add_run().add_break(); p.add_run(n["title"])
            elif k == "front":
                self._p("RptFront", n["text"])
            elif k == "heading":
                self._p("RptChapter", n["text"])
            elif k == "h2":
                self._p("RptH2", n["text"])
            elif k == "h3":
                self._p("RptH3", n["text"].upper() if s["h3_caps"] else n["text"])
            elif k == "p":
                self._p("RptBody", n["text"], size=n.get("size"))
            elif k == "quote":
                p = self._p("RptBody", n["text"], italic=True)
                p.paragraph_format.left_indent = Inches(0.4); p.paragraph_format.right_indent = Inches(0.4)
            elif k in ("num", "bullets"):
                for i, it in enumerate(n["items"]):
                    p = self._p("RptBody", f"{i + 1}.\t{it}" if k == "num" else "•\t" + it)
                    p.paragraph_format.left_indent = Inches(0.3); p.paragraph_format.first_line_indent = Inches(-0.3)
            elif k == "fig":
                self._add_fig(n)
            elif k == "table":
                self._add_table(n)
            elif k == "code":
                self._add_code(n)
            elif k == "abbrev":
                self._table([1.4, 5.3], [[a, b] for a, b in n["items"]], head=False, borders=False, bold_first=True)
            elif k == "pagebreak":
                self.d.add_page_break()
            else:
                raise ValueError("unknown node " + k)

    def _add_fig(self, n):
        path = n["file"] if os.path.isabs(n["file"]) else os.path.join(self._img_base, n["file"])
        with Image.open(path) as im:
            iw, ih = im.size
        max_w = self.s["page_w"] - self.s["left"] - self.s["right"]
        w = min(float(n.get("w", 6.0)), max_w); h = w * ih / iw
        if h > 8.0:
            w *= 8.0 / h
        p = self._p("RptFigure"); p.add_run().add_picture(path, width=Inches(w))
        d = p._p.find(".//" + WP_DOCPR)
        if d is not None:
            d.set("descr", n["caption"])
        c = self._p("FigCaption")
        c.add_run(f'{self.s["fig_label"]} {self.fig_no[n["id"]]}: ').bold = True
        c.add_run(self._resolve(n["caption"]))

    def _add_table(self, n):
        c = self._p("TabCaption")
        c.add_run(f'{self.s["tab_label"]} {self.tab_no[n["id"]]}: ').bold = True
        c.add_run(self._resolve(n["caption"]))
        size = self.s["table_small"] if n.get("small") else None
        t = self._table(n["widths"], [n["head"]] + n["rows"], size=size)
        t.rows[0]._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        self._p("RptBody")

    def _add_code(self, n):
        """Code box. Long listings are cut into (a), (b), (c) at blank lines, because a box longer than a page would split."""
        lines = n["lines"] if "lines" in n else open(n["file"], encoding="utf-8").read().replace("\r", "").replace("\t", "  ").split("\n")
        lines = list(lines)
        while lines and not lines[-1].strip():
            lines.pop()
        limit = int(n.get("max_lines", 55)); parts = []; cur = []
        for ln in lines:
            cur.append(ln)
            if (len(cur) >= limit - 12 and not ln.strip()) or len(cur) >= limit:
                parts.append(cur); cur = []
        if cur:
            parts.append(cur)
        width = self.s["page_w"] - self.s["left"] - self.s["right"] - 0.05
        for i, chunk in enumerate(parts):
            t = self.d.add_table(rows=1, cols=1); el = t._tbl; pr = el.find(qn("w:tblPr"))
            for x in pr.findall(qn("w:tblStyle")):
                pr.remove(x)
            self._borders(el)
            cell = t.cell(0, 0); cell.width = Inches(width)
            cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd xmlns:w="{W}" w:val="clear" w:color="auto" w:fill="F4F4F4"/>'))
            for j, ln in enumerate(chunk):
                p = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
                p.style = self.d.styles["RptCode"]
                r = p.add_run(ln or " "); r.font.name = "Courier New"
                r._r.get_or_add_rPr().rFonts.set(qn("w:cs"), "Courier New")
            cap = n["caption"]
            if len(parts) > 1:
                suffix = f'({"abcdefghijkl"[i]})'
                if ":" in cap:
                    a, b = cap.split(":", 1); cap = f"{a}{suffix}:{b}"
                else:
                    cap = f"{cap} {suffix}"
            self._p("CodeCaption", cap)

    def add_references(self, title="REFERENCES"):
        self._p("RptChapter", title)
        for k in sorted(self.cite_no, key=self.cite_no.get):
            p = self._p("RptRef"); p.add_run(f"[{self.cite_no[k]}]\t{self.refs[k]}")
            p.paragraph_format.left_indent = Inches(0.42); p.paragraph_format.first_line_indent = Inches(-0.42)
            p.paragraph_format.tab_stops.add_tab_stop(Inches(0.42))

    # ---------------- final passes ----------------
    def no_split_tables(self):
        """Keep each table on one page: rows cannot split and every row stays with the next. A table longer than a page
        still breaks, and word_pass reports it as SPLIT so it can be cut into two numbered tables."""
        def flag(p, tag):
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr"); p.insert(0, ppr)
            if ppr.find(qn(tag)) is None:
                ppr.append(OxmlElement(tag))
        for tbl in self.body.iter(qn("w:tbl")):
            rows = tbl.findall(qn("w:tr"))
            nxt = tbl.getnext()
            caption_below = nxt is not None and nxt.tag == qn("w:p") and "CodeCaption" in (nxt.xpath("./w:pPr/w:pStyle/@w:val") or [""])
            for i, tr in enumerate(rows):
                trpr = tr.find(qn("w:trPr"))
                if trpr is None:
                    trpr = OxmlElement("w:trPr"); pex = tr.find(qn("w:tblPrEx"))
                    if pex is not None:
                        pex.addnext(trpr)
                    else:
                        tr.insert(0, trpr)
                if trpr.find(qn("w:cantSplit")) is None:
                    trpr.insert(0, OxmlElement("w:cantSplit"))
                for p in tr.iter(qn("w:p")):
                    flag(p, "w:keepLines")
                    if i < len(rows) - 1 or caption_below:
                        flag(p, "w:keepNext")

    def schema_order(self):
        """Word calls a file corrupt when pPr/rPr/trPr children are out of schema order. Sort them."""
        order = {
            "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange",
            "rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath",
            "trPr": "cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden ins del trPrChange",
            "tblPr": "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption tblDescription",
            "tcPr": "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark",
            "sectPr": "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings",
        }
        order = {k: {n: i for i, n in enumerate(v.split())} for k, v in order.items()}
        for tag, rank in order.items():
            for el in self.body.iter(qn("w:" + tag)):
                kids = list(el)
                srt = sorted(kids, key=lambda x: rank.get(x.tag.split("}")[1], 999) if isinstance(x.tag, str) else 999)
                if srt != kids:
                    for x in kids:
                        el.remove(x)
                    for x in srt:
                        el.append(x)

    def save(self, out):
        self.no_split_tables(); self.schema_order()
        self.d.save(out)
        return dict(figures=len(self.fig_no), tables=len(self.tab_no), references=len(self.cite_no))
