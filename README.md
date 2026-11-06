# srm-report

A Claude Code skill that builds a print-ready SRM Institute report (.docx and .pdf) from the template you attach. It fills a copy of the template, so fixed pages stay as issued, and generates the numbered chapters, contents, lists and references.

## Install

```bash
npx srm-report install
```

This copies the skill to `~/.claude/skills/srm-report`. Then attach your template in Claude Code and ask for your report.

To build the report you need Windows with Microsoft Word, and Python with `python-docx`, `PyMuPDF` and `Pillow`:

```bash
pip install python-docx PyMuPDF Pillow
```

## The one rule: send your improvements back

Everyone who installs this is asked to raise a pull request with what they improved: new template recipes, rules from their course, bug fixes. Your report and personal details stay with you.

```bash
npx srm-report contribute     # shows the diff, sends nothing
```

This is a strict guideline for AI agents using the skill: they must record improvements as they work and ask you at the end of every report. Nothing is sent until you have seen the changes and agreed. See [CONTRIBUTING.md](CONTRIBUTING.md).

## What's inside

- `SKILL.md`: the workflow
- `references/`: rules, report families, content rules
- `scripts/`: template inspector, builder library, Word pass, page renderer, checker
- `assets/examples/`: full recipe for the 21CSE734P / 21CSP302L project report
- `bin/`: the installer and the `contribute` command

## License

MIT
