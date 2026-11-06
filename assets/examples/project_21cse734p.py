#!/usr/bin/env python3
"""Recipe: 21CSE734P / 21CSP302L project report from the official 'report format.docx'.
Copy it into the project's report/source/ folder, then write facts.json and content.json.

usage: python -X utf8 build.py template.docx content.json facts.json out.docx

facts.json   = {"title", "course", "month", "dept", "degree", "specialization" (null for core CSE),
                "students": [["Name", "RA..."]], "guides": [{"name", "desig", "label"}] (senior first),
                "hod": {"name", "desig"}, "faculty_advisors": ["Dr. ..."], "bold_names": ["Dr. ..."]}
content.json = {"abstract": ["para"], "abbreviations": [["API", "Application Programming Interface"]],
                "chapters": [nodes], "appendices": [nodes], "refs": {"key": "text"}, "img_dir": "img"}

Paragraph prefixes below match the template this recipe was written against. If the attached template differs, find()
raises KeyError naming the missing line: run inspect_template.py and adjust the prefix. Don't catch the error.
"""
import json, os, re, sys

SKILL = os.environ.get("SRM_REPORT_SKILL", os.path.expanduser("~/.claude/skills/srm-report"))
sys.path.insert(0, os.path.join(SKILL, "scripts"))
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from rptlib import Report, W, ptext, set_text


def cover(r, F):
    names, regs = zip(*F["students"])
    dept = F["dept"]
    r.set_para("PROJECT TITLE", F["title"].upper())
    r.set_para("21CSP302L", f'{F["course"]} - PROJECT')
    first = r.set_para("STUDENT1 NAME", f"{names[0].upper()} [{regs[0]}]")
    second = r.find("STUDENT2 NAME") if any(ptext(e).strip().startswith("STUDENT2 NAME") for e in r.els()) else None
    if second is not None:
        second.getparent().remove(second)
    cur = first
    for n, g in F["students"][1:]:
        cur = r.clone_after(first, f"{n.upper()} [{g}]", after=cur)
    cur.find(qn("w:pPr")).find(qn("w:spacing")).set(qn("w:after"), "120")   # gap before "Under the Guidance of"
    g0 = F["guides"][0]
    name_p = r.set_para("(GUIDE NAME", g0["name"].upper())
    des_p = r.set_para("(Designation, Department)", f'({g0["desig"]}, Department of {dept})')
    cur = des_p
    for g in F["guides"][1:]:
        cur = r.clone_after(name_p, g["name"].upper(), after=cur)
        cur = r.clone_after(des_p, f'({g["desig"]}, Department of {dept})', after=cur)
    r.set_para("COMPUTER SCIENCE ENGINEERING", F["degree"])
    spec = r.find("with specialization in")
    if F.get("specialization"):
        set_text(spec, f'with specialization in {F["specialization"]}')
    else:
        spec.getparent().remove(spec)   # core CSE: leave the line out, never print the placeholder
    r.set_para("DEPARTMENT OF COMPUTATIONAL", f"DEPARTMENT OF {dept.upper()}\nCOLLEGE OF ENGINEERING AND TECHNOLOGY")
    r.set_para("MAY", F["month"])


def declaration(r, F):
    names, regs = zip(*F["students"])
    r.replace("Department of Computational Intelligence", f'Department of {F["dept"]}')
    r.fill_field("Degree/ Course", f'B.Tech, {F["degree"].title()} / {F["course"]} - Project')
    r.fill_field("Student Name", ", ".join(names))
    r.fill_field("Registration Number", ", ".join(regs))
    r.fill_field("Title of Work", F["title"])
    r.set_para("I / We hereby certify", "We hereby certify that this assessment complies with the University’s Rules and Regulations relating to Academic misconduct and plagiarism, as listed in the University Website, Regulations, and the Education Committee guidelines.")
    r.set_para("I / We confirm", "We confirm that all the work contained in this assessment is our own except where indicated, and that we have met the following conditions:")
    r.set_para("Given the sources", "Given the sources of all pictures, data etc. that are not our own")
    r.set_para("Acknowledged in appropriate", "Acknowledged in appropriate places any help that we have received from others (e.g. fellow students, technicians, statisticians, external sources)")
    r.set_para("Compiled with any other", "Complied with any other plagiarism criteria specified in the Course handbook / University website")
    r.set_para("I understand that", "We understand that any false claim for this work will be penalized in accordance with the University policies and regulations.")
    r.set_para("I am aware of", "We are aware of and understand the University’s policy on Academic misconduct and plagiarism and we certify that this assessment is our own work, except where indicated by referring, and that we have followed the good academic practices noted above.")
    sign = r.set_para("<<Student 1", "x")
    ppr = sign.find(qn("w:pPr"))
    for x in ppr.findall(qn("w:tabs")):
        ppr.remove(x)
    ppr.append(parse_xml(f'<w:tabs xmlns:w="{W}"><w:tab w:val="center" w:pos="2382"/><w:tab w:val="center" w:pos="7146"/></w:tabs>'))
    lines = [f"{n} [{g}]" for n, g in F["students"]]
    rows = [lines[i:i + 2] for i in range(0, len(lines), 2)]
    set_text(sign, "\t" + "\t".join(rows[0]))
    cur = sign
    for row in rows[1:]:
        cur = r.clone_after(sign, " ", after=cur)
        cur = r.clone_after(sign, " ", after=cur)
        cur = r.clone_after(sign, "\t" + "\t".join(row), after=cur)


def bonafide(r, F):
    dept = F["dept"].upper()
    all_names = ", ".join(f"{n.upper()} [{g}]" for n, g in F["students"])
    cert = r.find("Certified that")
    text = ptext(cert).replace("21CSP302L", F["course"])
    text = re.sub(r"PROJECT TITLE\s*", F["title"].upper(), text)
    text = text.replace("STUDENT1 NAME [REG NUM], STUDENT2 NAME [REG NUM]", all_names)
    # "our" even with one guide: the guide and the HoD both sign this page.
    text = text.replace("under my supervision", "under our supervision").replace("best of my knowledge", "best of our knowledge")
    set_text(cert, text)
    set_text(r.find("<<Signature"), " ")   # the line stays so the signature block keeps its place
    tbl = r.tables_with("SIGNATURE")[0]
    blocks = [[("SIGNATURE", True), (g["name"], False), (g.get("label", "GUIDE"), True), (g["desig"], False), ("DEPARTMENT OF", False), (dept, False)] for g in F["guides"]]
    blocks.append([("SIGNATURE", True), (F["hod"]["name"], False), ("PROFESSOR & HEAD", True), (" ", False), ("DEPARTMENT OF", False), (dept, False)])
    r.signature_table(tbl, blocks)


def acknowledgements(r, F):
    dept = F["dept"]
    gs = F["guides"]; fa = F["faculty_advisors"]
    r.replace("<<Name , Designation & Department>>", f'{F["hod"]["name"]}, {F["hod"]["desig"]}, Department of {dept}')
    r.replace(f"Panel Members Department of {dept}", f"Panel Members, Department of {dept}")
    fa_txt = fa[0] if len(fa) == 1 else ", ".join(fa[:-1]) + " and " + fa[-1]
    r.set_para("We register our immeasurable", f'We register our immeasurable thanks to our Faculty Advisor{"s" if len(fa) > 1 else ""}, {fa_txt}, Department of {dept}, SRM Institute of Science and Technology, for leading and helping us to complete our course.')
    gtxt = ", and ".join(f'{g["name"]}, {g["desig"]}' for g in gs)
    r.set_para("Our inexpressible respect", f'Our inexpressible respect and thanks to our guide{"s" if len(gs) > 1 else ""}, {gtxt}, Department of {dept}, SRM Institute of Science and Technology, for providing us with an opportunity to pursue our project under {"their" if len(gs) > 1 else "this"} mentorship, and for the freedom and support to explore the research topics of our interest.')
    r.replace("<<Dept. Name>>", dept)
    authors = next(e for e in r.els() if ptext(e).strip() == "Authors")
    set_text(authors, f'{F["students"][0][0]} [{F["students"][0][1]}]')
    r.trim_blank_before(authors, keep=1)   # a longer author list would otherwise spill to a new page
    r.start_new_page(r.find("ACKNOWLEDGEMENTS"))
    r.tidy_paragraphs("ACKNOWLEDGEMENTS", "ABSTRACT")
    last = next((p for p in r.all_paras() if ptext(p).rstrip().endswith("support and encouragement")), None)
    if last is not None:   # the template forgets this full stop
        ts = list(last.iter(qn("w:t"))); ts[-1].text = ts[-1].text.rstrip() + "."
    ppr = authors.find(qn("w:pPr"))
    for x in ppr.findall(qn("w:jc")) + ppr.findall(qn("w:ind")) + ppr.findall(qn("w:tabs")):
        ppr.remove(x)
    ppr.append(parse_xml(f'<w:jc xmlns:w="{W}" w:val="right"/>'))
    cur = authors
    for n, g in F["students"][1:]:
        cur = r.clone_after(authors, f"{n} [{g}]", after=cur)
    r.bold_names("ACKNOWLEDGEMENTS", "ABSTRACT", F.get("bold_names", []) + [F["hod"]["name"]] + fa + [g["name"] for g in gs])


def build(template, content_path, facts_path, out):
    F = json.load(open(facts_path, encoding="utf-8"))
    c = json.load(open(content_path, encoding="utf-8"))
    r = Report(template)
    r.apply_cover_notes()   # before filling, so cloned student and guide lines inherit the right size
    cover(r, F); declaration(r, F); bonafide(r, F); acknowledgements(r, F)
    r.strip_notes()
    r.cut_from(r.find("ABSTRACT"))
    r.prepare(c["chapters"] + c.get("appendices", []), c["refs"], c.get("img_dir", "."))
    r.begin_front(); r.number_kept_sections()
    r.add_nodes([{"t": "front", "text": "ABSTRACT"}] + [{"t": "p", "text": a} for a in c["abstract"]])
    r.add_toc(); r.add_list_of_figures(); r.add_list_of_tables()
    r.add_nodes([{"t": "front", "text": "ABBREVIATIONS"}, {"t": "abbrev", "items": c["abbreviations"]}])
    r.begin_body(); r.add_nodes(c["chapters"]); r.add_references()
    r.add_nodes(c.get("appendices", []))
    print(r.save(out))


if __name__ == "__main__":
    if len(sys.argv) != 5:
        sys.exit(__doc__)
    build(*sys.argv[1:5])
