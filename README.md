# srm-report

A Claude Code skill that builds a print-ready SRM Institute report (.docx and .pdf) from the template you attach. It fills a copy of the template, so fixed pages stay as issued, and it generates the numbered chapters, contents, lists and references.

Install by placing this folder at `~/.claude/skills/srm-report`. Needs Windows with Microsoft Word, Python with `python-docx`, `PyMuPDF` and `Pillow`.

- `SKILL.md`: the workflow
- `references/`: rules, report families, content rules
- `scripts/`: template inspector, builder library, Word pass, page renderer, checker
- `assets/examples/`: full recipe for the 21CSE734P / 21CSP302L project report
