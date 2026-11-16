# Report families seen so far

Use this for the shape of a report only. Exact wording, page order and sizes come from the attached template.

## Project report (21CSE734P, 21CSP302L)

Fully tested recipe: `assets/examples/project_21cse734p.py`.

1. **Introduction:** introduction, problem statement, motivation, Sustainable Development Goal
2. **Literature Survey:** overview, existing models, gaps, objectives, product backlog, plan of action
3. **Sprint Planning and Execution:** for each sprint: objectives with user stories, functional document, architecture document, outcome and result analysis, retrospective
4. **Results and Discussions**
5. **Conclusion and Future Enhancement**

Then References, Appendix A Coding, B Conference Publication, C Journal Publication, D Plagiarism Report. Drop and re-letter B and C if there are no publications and the user agrees.

## Community Connect (21GNP301L)

The template comes with a separate rubric, so read it as well. No recipe yet: copy the project recipe and rewrite the fill steps.

- Cover
- Bonafide with Faculty Advisor, Panel Member, Academic Advisor and HoD, worded "under NAME OF THE ORGANIZATION"
- Completion certificate from the organisation (the user supplies it)
- Log sheet
- Objectives: one paragraph of at most 300 words covering the services and the reason for choosing the organisation
- Table of Contents
- Ch 1 Introduction of the organisation, at least 2 pages
- Ch 2 Week-wise report
- Ch 3 Tangible learning, at least 2 pages
- Ch 4 Student's contribution, 2 pages per student with their own photos (mandatory)
- Ch 5 Conclusion
- Binding instructions: follow them, and don't print them in the report

Photos need matching dates and sessions. Ask for them rather than reusing whatever is at hand.

## Course case studies and small projects (APP, OS, DBMS, DIP)

These vary the most. One OS case study template uses 14pt body text on a Letter page. Take the page size, margins and body size from `sections.txt` and `structure.txt`, and pass them through `Report(template, style={...})`. These templates are often short, so most of their pages may be kept and filled in place, with no cut at all.

### 21CSC302J Computer Networks course case study

- The template keeps only two fixed pages: the cover (title 26pt, "Student Name [Reg.No]", guide, degree, logo, month) and a bonafide with one Faculty In-Charge block. Everything after "ABSTRACT" is a specimen: a TOC table, a "Report Structure Overview", chapter headings with guidance text, and a formatting page. Cut from ABSTRACT and generate the rest.
- The overview lists a Declaration and Acknowledgements in the front matter, but the template gives no page for them. Generate both, in "we", after the bonafide.
- Formatting comes from the template's own sentences: Times New Roman; chapter headings "Chapter N: Title", 16pt bold, Title Case, left-aligned; sections 14pt bold; sub-sections 12pt bold (not capitals); body 12pt with 1.5 spacing; margins left 1.5 in, the others 1 in; roman front matter and arabic from Chapter 1. Override `RptChapter` to left alignment and emit "Chapter N: Title" yourself, since `add_nodes` prints "CHAPTER N".
- The abstract is a single paragraph of at most 300 words in 14pt, as the template's own rule sentence says. This is a stated rule, not an inline note, so it is followed.
- Chapters follow the specimen TOC: Introduction (background, domain context, problem statement, objectives, scope), Literature Survey (introduction, existing approaches, comparative analysis), Methodology (approach, requirements, architecture, implementation strategy), Implementation (execution, core configurations, challenges), Results and Discussion (findings, discussion, practical implementations), Conclusion and Future Scope, then References and appendices.
- Template typos to fix quietly, and mention: "Batchelor Of Technology" and "SRM INSTITUTION OF SCIENCE AND TECHNOLOGY" on the bonafide.
- With a large team, a long title at 26pt plus six names overflows the cover. 20pt for the title and the spacing fix in rules.md keep it on one page.


### 21CSC305P Machine Learning integrated project report (PBL 1-3)

Recipe: `assets/examples/ml_project_21csc305p.py` (run end to end on the real template, 113 pages, 0 checker fails).

- The only fixed page is the cover: title 20pt bold, "21CSC305P- MACHINE LEARNING PROJECT", "Name  [Reg No]" lines (two in the specimen; clone for more), "Faculty Name" and "(Designation, Department of Computational Intelligence)", degree block with a "with specialization in AIML" line, logo, department block, month. No declaration, bonafide or acknowledgements page exists in the template, so none is generated. Everything from ABSTRACT on is specimen: cut and generate.
- The last page of the template is a formatting table and is the spec: A4, Times New Roman, title 18pt, chapter heading 16pt bold caps centred, section 14pt bold left, sub-section 12pt bold left, body 12pt justified 1.5 spacing 6pt after, margins left 1.5" and 1" elsewhere, captions 11pt bold (the whole caption, not only the label), table text 11pt, code Consolas 10pt single. "Figure 2.1" below, "Table 2.1" above, equations numbered (2.1) chapter-wise. Pass these through `style=` and `code_font`; the specimen headings' own sizes (16/18pt) disagree with the table, and the table wins.
- Abstract: one paragraph, 150-250 words, no citations, equations or tables; every unit represented, actual results.
- Chapters: 1 Introduction (background, integrated problem statement, objectives per unit, unit-wise allocation table), 2 Units 1-2 (topics, dataset, preprocessing, EDA, workflow diagram, theory, results, discussion, conclusion), 3 Units 3-4 with the template's 3.1-3.13 headings (3.11 "Results and Performance Evaluation" holds the Unit 4 results and a separate-units summary), 4 Unit 5 (4.1-4.9, including "Comparison with Previous Models" on the same folds and metric), 5 Integrated Discussion, 6 Conclusion and Future Scope, References, Appendix A Coding.
- Appendix A must hold the complete Python for all five units plus Colab links. A public GitHub link per file is an accepted substitute if the user says so. Group listings by the template's appendix lines (Chapter 2 / Unit 3 / Unit 4 / Chapter 4). Full source makes the appendix ~45 pages; use `wrap` on code nodes (see rules.md).
- Every chapter needs a workflow diagram; draw them with matplotlib in Times New Roman on white if the project has none.
