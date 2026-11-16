#!/usr/bin/env python3
"""Recipe: 21CSC305P Machine Learning integrated project report (PBL 1-3) from 'Machine Learning Project Report_Sample template.docx'.

usage: python -X utf8 ml_project_21csc305p.py template.docx facts.json content.py out.docx
facts.json = {"title", "students": [["Name", "RA..."]], "guide": "Dr. ...", "desig": "Assistant Professor", "dept": "...", "code_root": "../.."}
content.py = a module defining ABSTRACT (str), ABBREV ([[abbr, meaning]]), REFS ({key: text}), CH1..CH6 (node lists),
             APPENDIX_INTRO (nodes) and code_sections(root) -> nodes. Nodes are rptlib's, plus {"t": "eq", "text": ...}.
The template keeps only its cover page; everything from ABSTRACT on is a specimen and is generated here, following the
template's own formatting table (A4, Times New Roman, chapter 16pt bold caps centred, section 14pt, sub-section 12pt,
body 12pt justified 1.5 spacing 6pt after, margins L 1.5" R/T/B 1", captions 11pt bold, table text 11pt,
code Consolas 10pt single, Figure/Table/Equation numbered chapter-wise).
"""
import os, re, sys
HERE = os.path.dirname(os.path.abspath(__file__))
PLACES = [os.environ.get("SRM_REPORT_SKILL"), os.path.join(HERE, "..", "..")] + [
    os.path.expanduser(f"~/{p}/skills/srm-report") for p in (".agents", ".claude", ".gemini/config", ".cline", ".kiro")]
SKILL = next((p for p in PLACES if p and os.path.exists(os.path.join(p, "scripts", "rptlib.py"))), None)
if SKILL is None:
    sys.exit("Can't find the srm-report skill. Set SRM_REPORT_SKILL to its folder.")
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from rptlib import Report, W, ptext, set_text
import importlib.util, json


STYLE = dict(font="Times New Roman", body=12, line=1.5, after=6, h1=16, h2=14, h3=12, h3_caps=False, front_h=16,
             caption=11, table=11, table_small=10, code=10, code_font="Consolas", left=1.5, right=1.0, top=1.0,
             bottom=1.0, page_w=8.27, page_h=11.69, fig_label="Figure", tab_label="Table")


class MLReport(Report):
    """Adds chapter-wise numbered equations: {"t": "eq", "text": "..."} -> centred equation, (c.n) at the right margin."""
    eq_ch, eq_n = "0", 0

    def add_nodes(self, nodes):
        for n in nodes:
            if n["t"] == "chapter":
                self.eq_ch, self.eq_n = str(n["num"]), 0
            if n["t"] == "eq":
                self.eq_n += 1
                p = self._p("RptEq")
                r = p.add_run("\t" + n["text"]); r.italic = False
                r.font.name = "Cambria Math"; r._r.get_or_add_rPr().rFonts.set(qn("w:hAnsi"), "Cambria Math")
                p.add_run(f"\t({self.eq_ch}.{self.eq_n})")
            else:
                super().add_nodes([n])

    def restyle(self):
        s = self.s
        tw = int((s["page_w"] - s["left"] - s["right"]) * 1440)
        self._para_style("FigCaption", s["caption"], bold=True, jc="center", line=1.15, before=3, after=12)
        self._para_style("TabCaption", s["caption"], bold=True, jc="center", line=1.15, before=6, after=6, keep_next=True)
        self._para_style("CodeCaption", s["caption"], bold=True, jc="center", line=1.15, before=3, after=12)
        self._para_style("RptEq", s["body"], jc="left", line=1.2, before=4, after=8)
        st = self.d.styles["RptEq"].element
        st.find(qn("w:pPr")).insert(0, parse_xml(f'<w:tabs xmlns:w="{W}"><w:tab w:val="center" w:pos="{tw // 2}"/><w:tab w:val="right" w:pos="{tw}"/></w:tabs>'))


def cover(r, F):
    r.set_para("Design, Implementation, and Performance", F["title"])
    s1 = r.set_para("Name  [Reg No]", "x")
    s2 = r.find("Name  [Reg No]")
    s2.getparent().remove(s2)
    first = s1
    set_text(first, f'{F["students"][0][0]}  [{F["students"][0][1]}]')
    cur = first
    for n, g in F["students"][1:]:
        cur = r.clone_after(first, f"{n}  [{g}]", after=cur)
    r.set_para("Faculty Name", F["guide"])
    r.set_para("(Designation, Department of Computational Intelligence)",
               f'({F["desig"]}, Department of {F["dept"]})')
    spec = r.find("with specialization in AIML")
    spec.getparent().remove(spec)      # core CSE: no specialization line
    r.set_para("DEPARTMENT OF COMPUTATIONAL INTELLIGENCE", f'DEPARTMENT OF {F["dept"].upper()}')
    # The template pads some cover lines with spaces, offsets others with indents, and squeezes "AND TECHNOLOGY" with
    # negative character spacing. Centre every cover line on the page axis instead.
    end = r.find("ABSTRACT")
    for p in r.els():
        if p is end:
            break
        if p.tag != qn("w:p") or not ptext(p).strip():
            continue
        t = ptext(p)
        if t != t.strip() or "  " in t:
            set_text(p, re.sub(r"\s{2,}", " ", t).strip()) if "[" not in t else None
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            continue
        for x in ppr.findall(qn("w:ind")) + ppr.findall(qn("w:jc")):
            ppr.remove(x)
        ppr.append(parse_xml(f'<w:ind xmlns:w="{W}" w:left="0" w:right="0" w:firstLine="0"/>'))
        ppr.append(parse_xml(f'<w:jc xmlns:w="{W}" w:val="center"/>'))
        for rpr in p.iter(qn("w:rPr")):
            for x in rpr.findall(qn("w:spacing")):
                rpr.remove(x)
    return s1


def keep_code_links_with_listing(r):
    """The 'Code: <link>' line under each appendix heading must stay on the page of its first listing."""
    for p in r.all_paras():
        if ptext(p).startswith("Code: https://"):
            ppr = p.find(qn("w:pPr"))
            if ppr is None:
                ppr = parse_xml(f'<w:pPr xmlns:w="{W}"/>'); p.insert(0, ppr)
            ppr.insert(0, parse_xml(f'<w:keepNext xmlns:w="{W}"/>'))


def build(template, facts_path, content_path, out):
    FACTS = json.load(open(facts_path, encoding="utf-8"))
    spec = importlib.util.spec_from_file_location("content", content_path)
    C = importlib.util.module_from_spec(spec); spec.loader.exec_module(C)
    ROOT = os.path.dirname(os.path.abspath(content_path))
    r = MLReport(template, style=STYLE)
    cover(r, FACTS)
    r.cut_from(r.find("ABSTRACT"))
    chapters = C.CH1 + C.CH2 + C.CH3 + C.CH4 + C.CH5 + C.CH6
    # facts "code_root": the project folder the listings are read from (relative to content.py), default its folder
    appendix = C.APPENDIX_INTRO + C.code_sections(os.path.normpath(os.path.join(ROOT, FACTS.get("code_root", "."))))
    r.prepare(chapters + appendix, C.REFS, ROOT)
    r.begin_front(); r.restyle(); r.number_kept_sections()
    r.add_nodes([{"t": "front", "text": "ABSTRACT"}, {"t": "p", "text": C.ABSTRACT}])
    r.add_toc(); r.add_list_of_figures(); r.add_list_of_tables()
    r.add_nodes([{"t": "front", "text": "ABBREVIATIONS"}, {"t": "abbrev", "items": C.ABBREV}])
    r.begin_body(); r.add_nodes(chapters); r.add_references(); r.add_nodes(appendix)
    keep_code_links_with_listing(r)
    print(r.save(out))
    print("abstract words:", len(C.ABSTRACT.split()))


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    build(*sys.argv[1:5])
