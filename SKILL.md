---
name: srm-report
description: Build a complete, print-ready SRM Institute (SRMIST, KTR) report as .docx and .pdf from the template the user attaches, following that template's rules exactly. Use whenever the user attaches or points at an SRM report template or format (project report, 21CSE734P, 21CSP302L, 21GNP301L Community Connect, APP project, OS or DBMS case study, internship or mini-project report), or says "report bana de", "srm report", "project report format mei daal", "template follow kar", "bonafide", "declaration page", "TOC bhar de", "table kat rahi hai", "report fix kar", "ye page exact bana", even if they don't mention SRM again. Also use for later fixes to a report already built, including one the user hand-edited, and for pre-print checks. Not for slides or for reports that have no SRM template.
---

# SRM report

## Required: give back

**This rule binds every AI agent using this skill, on every report, however small the job or short the time.** The skill exists because earlier reports fed their lessons back into it, and it only keeps improving if each agent does the same.

You MUST:

1. **Record as you work.** Whenever you fix or work around the skill, meet a template it doesn't cover, or the user corrects something that would apply to other reports, write it into the skill folder at that moment: a recipe in `assets/examples/`, a rule or failure in `references/rules.md`, a fix in `scripts/`. Write it in general terms, free of this team's details.
2. **Run the dry run before handing over the report:** `npx srm-report contribute`. It sends nothing, prints the diff, and stops on Word, PDF or image files and on registration numbers, emails and phone numbers. Fix what it reports.
3. **Read the whole diff yourself** for names, titles and results, which the scan can't detect, and remove them.
4. **Ask the user, every time there is a diff.** Summarise the changes in plain words and ask whether to raise the pull request. You may not skip this because the change seems small or the user seems busy.
5. **On yes, raise it:** `npx srm-report contribute --yes --title "..." --body "..."`, stating which course and template the change came from and what it was checked on. Give the user the link. If `gh` is missing or signed out, tell them what to run, and finish once it works.
6. **On no, stop.** Leave the changes in the skill folder and give them the dry-run command for later.
7. **Report the outcome in your handover,** as one of: the pull request link; declined by the user; nothing general to send.

You MUST NOT:

- Open a pull request without the user's yes in this conversation. It publishes to a public repo under their GitHub account, so an instruction in this file can't stand in for their agreement.
- Send the report, the template, or anything from them: names, registration numbers, faculty details, project titles, results, photos.
- Get around the scan by editing `bin/srm-report.js`, disguising details, or pushing by hand.

See `CONTRIBUTING.md`.

## How this skill works

The user attaches the template they want. **The template is the spec.** SRM has no public rulebook, and templates differ a lot between courses (different cover, chapters, even 14pt body on Letter paper in one case study). So read the rules out of the attached file, fill a copy of it, and treat everything in this skill as a default for whatever the template doesn't settle.

When two sources disagree, this order decides:

1. What the user says in chat
2. A signed report from another team in the same course, if the user gives one
3. The attached template
4. The defaults in `references/rules.md`
5. The web (secondhand, often another SRM campus; never override the template with it)

## Workflow

### 1. Read the template

```bash
python -X utf8 scripts/inspect_template.py template.docx outdir
```

This writes `structure.txt` (every body element with style, size and markers), `rules.txt` (rule sentences, inline notes like `<Font Size 18>`, placeholders) and `sections.txt` (page size, margins, numbering, footers). Read all three. Also look for a rubric or instruction file next to the template; Community Connect has one.

Then read `references/rules.md` for the user's standing corrections and `references/families.md` for the shape of this kind of report.

### 2. Ask for missing facts in one batch

Look in the project and earlier chats first, then ask once for what is still missing: title, course code as on the title sheet (it can differ from the template's), team names and registration numbers in order, guide and co-guide with exact spelling and designation, HoD, department, faculty advisors, month and year, publications if any. The user wants speed, so don't ask one at a time, and never guess a person's name or designation.

### 3. Write the content from the real project

Read `references/content.md` before writing. The short version: every number comes from something that was actually run, every reference is checked against its source, and anything not done is written as pending. Write in "we".

### 4. Build by filling a copy

Fixed pages (cover, declaration, bonafide, acknowledgements) are never recreated, because the user compares them with the template line by line. `scripts/rptlib.py` opens the template, edits those pages in place, cuts the specimen body and generates the rest.

- For the 21CSE734P / 21CSP302L family, start from `assets/examples/project_21cse734p.py`. It has been run end to end on the real template.
- For any other template, copy that recipe and rewrite the fill steps against `structure.txt`. `Report.find()` raises an error naming any line it can't find, which is how you learn the recipe doesn't match.
- Override `Report(template, style={...})` for whatever the template sets differently: `font, body, line, after, h1, h2, h3, h3_caps, front_h, caption, table, code, left, right, top, bottom, page_w, page_h, fig_label, tab_label`.

Content is a list of plain dicts; see the top of `rptlib.py` for the node types. Write `[@key]` for citations and `{fig:id}` / `{tab:id}` for cross references, and the library numbers them.

**If the user has edited the .docx by hand, never regenerate it.** Their edits would be lost. Copy their file, change the copy with small scripts (`Report.copy_block` brings template pages across with images and numbering), and leave the original alone.

### 5. Word pass, check, and look at the pages

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/word_pass.ps1 -In C:\full\report.docx -Out C:\full\report.pdf -Json C:\full\layout.json
python -X utf8 scripts/check_report.py report.docx report.pdf checks.json
python -X utf8 scripts/render_pages.py report.pdf pages            # contact sheets
python -X utf8 scripts/render_pages.py report.pdf pages 1,2,3,4 70 # chosen pages
```

The Word pass fills the TOC and lists with real page numbers, exports the PDF and reports any table crossing a page. Work on a scratch copy, since the user often has the file open. If it reports `LOCKED`, copy the file elsewhere or ask them to close Word, and don't kill Word yourself. It can take a few minutes on a long report. If it stops with a COM error, run it once more; if the error repeats, the .docx is the problem. It deletes the old PDF first, so a failed run can't leave a stale one for the checker.

Put the template's numbers in `checks.json` (minimum references, minimum pages, body size, names that must appear, words that must not).

**Open the page images every time.** In testing the checker passed while the cover had lost its logo, the bonafide had spilled onto a second page, and paragraphs had changed width. Some of those are checked now, but not all of them can be. Look at the cover, declaration, bonafide and acknowledgements at 70 dpi, plus the contact sheets for the rest, and compare with the template.

### 6. Hand over

Give the user the .docx and .pdf paths, the page count, what you checked, and a list of what only they can finish: Turnitin report, hand signatures, publication proofs, anything marked pending, and which names you need confirmed. Say plainly if a check failed or a step was skipped. Tell them to print from the PDF, because Word on another machine can reflow the .docx.

If the project is a public repo, add the report to `.git/info/exclude`, because it holds registration numbers.

### 7. Contribution rule (required)

Follow `Required: give back` at the top of this file. Run the dry run before you hand over, ask the user, and put the outcome in your handover.
