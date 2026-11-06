# Rules and lessons

These come from real report sessions and the signed reports compared against them. The attached template overrides the formatting defaults in section 1. The corrections in section 2 came from the user and apply to every report.

## 1. Formatting defaults (21CSP302L / 21CSE734P template)

| Item | Rule |
|---|---|
| Font | Times New Roman throughout (code may use Courier New) |
| Chapter number and title | 16pt bold caps, centred, every chapter on a new page |
| Sub-heading (1.1) | 16pt bold Title Case |
| Sub-sub-heading (1.1.1) | 12pt bold capitals |
| Body | 12pt, 1.5 spacing, justified, space between paragraphs, at most 12 lines a paragraph |
| Page | A4, left margin 1 inch, right 0.5 inch |
| Figures | "Fig X.Y" centred, caption below |
| Tables | "Table X.Y" centred, caption above |
| Page numbers | Bottom centre. Lower roman on front matter, arabic from Chapter 1 (page 1) to the end |
| Citations | [1], [4,5], [7-10], numbered by first use, listed in the same order, at least 15 |
| Length | At least 30 pages, excluding code and screenshots |
| Plagiarism | Turnitin through the guide, 10% or less; publication proofs where publications are claimed |

Front order: cover, Own Work Declaration, Bonafide, Acknowledgements, Abstract, Table of Contents, List of Figures, List of Tables, Abbreviations.

Older departmental sheets online mention 15% plagiarism and other margins, and SRM NCR uses different margins again. They are not sources for this campus. Trust the template.

### Reading template notes

- A note such as `<Font Size 18> <1.5 line spacing>` is an instruction for the text above it. Apply it, then delete it. `apply_cover_notes()` does this for the cover. `<Font Style ... Bold>` at the top of the cover covers the whole cover.
- `<Font Style Times New Roman, Font Size 14>` near the abstract is not followed in the signed report or in the accepted report: abstract heading 16pt, body 12pt. Keep 12pt unless the user says otherwise.
- A note's paragraph may also hold the logo, so check for an image before deleting it.
- Negative character spacing in template runs squeezes words together on the page ("requirementsfor"). The text is fine, so the fix is removing the spacing.

### Template errors to fix quietly, and mention

Chapter numbering jumps from 3 to 6; "2.5" used twice; "EXECTION"; "compiles with" and "Compiled with" (should be "complies" and "Complied"); no full stop on the last acknowledgement line.

## 2. Corrections from the user

**Fixed pages**
- Template pages must be exact ("ye wala page exact bana"). Fill a copy; never recreate them.
- Bonafide fits on one page. Signature block centred in its box. Each name has its designation, with the same line count in each column.
- Footer is the page number only, no rule line.
- A page border only if asked.

**Names and labels**
- The label is GUIDE / CO-GUIDE, not SUPERVISOR. The senior designation is GUIDE and goes first (Associate before Assistant).
- Get faculty spelling and designation right, and confirm them with the user. All faculty are named everywhere.
- Faculty Advisors are named individually, in the same order as the students.
- Names are bold in the acknowledgements.
- The template's Vice-Chancellor, Dean and Chairperson names and its "his/her" may be out of date, or wrong for the person now in the post. Ask the user to confirm them.

**Voice**
- "We", never "I", everywhere, including the declaration and bonafide ("under our supervision", "best of our knowledge"). User stories in a backlog stay in the first person.

**Course and degree**
- The course code from the user or title sheet beats the template (21CSE734P vs 21CSP302L), and must be identical on cover and bonafide.
- Core CSE drops the "with specialization in" line. Never print `[SPECIALIZATION NAME]`; a signed report did, so don't copy mistakes from examples.
- Replace degree and department everywhere the same way. The template says Computational Intelligence, which may not be the user's department.

**Tables and contents**
- No table crosses a page ("pages mei ek bhi table katna nahi"). White space at a page end is acceptable.
- TOC filled with real page numbers. Chapter lines bold, sub-topics not.

**Appendices**
- Remove placeholder appendices when asked ("appendix b hata de"). With no publications, drop Conference and Journal and re-letter the rest.
- Code listings longer than a page are split at logical boundaries into A.2(a), (b), (c). Keep each to about 55 lines; 74 was still too long.

**Truth**
- "saare rules updated hai na? real hai na sab?" Everything must be real. See `content.md`.

**Images**
- Don't crop preparation screenshots ("crop nahi kar").
- A photo goes in only after checking it matches the session and date.
- Take screenshots from a separate fixture-mode instance, never the user's live backend or quota.
- Figures: white background, Times New Roman, print-friendly.

**Before printing**
- Print from the PDF, not the .docx.
- Checklist for the user: Turnitin report, HoD name, month on the cover, hand signatures.
- Compare against a signed report if the user has one.

## 3. Failures already hit, and what prevents them

| Symptom | Cause | Prevention |
|---|---|---|
| Word says the file is corrupt | XML children out of schema order | `schema_order()` runs in `save()` |
| Corrupt file after copying a code box | Copied after trimming it | Keep a pristine deepcopy first |
| A word jumps to the next page | Cloned paragraph carried a `w:sectPr` | `clone_after` strips it |
| "DepartmentAssistant Professor," | Stale runs left behind after editing one run | `set_text` removes other runs |
| Three page numbers in one footer | Looping over a linked footer | Unlink and clear the footer first |
| TOC lists cover lines | Cover uses Heading styles, and a `\o` TOC picks them up | Custom styles and `TOC \t` |
| Bold TOC chapter lines also bold the List of Figures | Lists share toc 1 | Lists use toc 4 |
| Title printed twice on the cover | A broad `replace()` hit the already filled cover | Edit the one paragraph, not the whole document |
| Logo missing from the cover | Logo sat in a note paragraph that got deleted | Check for images first; the checker now tests this |
| Blank page before Abstract | Empty lines before the cut spilled over | `cut_from` trims them |
| Bonafide spills to a second page | Note lines kept as blanks | Delete them; the checker now tests this |
| One name alone on a page | The author list overflowed the acknowledgements | Trim blank lines above it; the checker now flags stray-line pages |
| Acknowledgements starts on the bonafide page | Template reaches the next page with blank lines alone | `start_new_page()` gives a real break; the checker tests main headings |
| Paragraphs on one page have different widths | Removed `w:ind`; the page depends on left and right indents | `tidy_paragraphs()` removes only first-line offsets |
| Checker reports results for the previous build | Word pass failed, and the old PDF was still there | `word_pass.ps1` deletes the old PDF first |
| "Device or resource busy", `~$` file | The user has the file open in Word | Work on a scratch copy; never kill Word |
| cp1252 errors in the console | Windows default encoding | Run Python with `-X utf8` |
| Shell errors on long inline scripts | Quoting in heredocs | Write the script to a file first |
| Can't read a source PDF through a web fetch | Binary file | Save it and read it with PyMuPDF |
| Word windows popping up | Visible automation | Keep `Visible=false`; never kill by window title |
