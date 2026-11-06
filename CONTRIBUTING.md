# Sending your improvements back

This skill gets better from real reports. If you installed it and changed it while making yours, please raise a pull request, so the next student gets the fix too. It's the one thing we ask of everyone who installs this package.

## What to send

- A recipe for a template the skill doesn't cover yet, as a new file in `assets/examples/`
- A rule your course follows, or a correction your faculty made, in `references/rules.md`
- A bug fix in `scripts/`, or a new check in `check_report.py`
- A failure you hit and what prevented it, in the table in `references/rules.md`

## What never goes in

Your report and anything from it: names, registration numbers, faculty details, project titles, results, photos, signed pages, and the template files themselves. The `contribute` command refuses Word, PDF and image files, and stops on registration numbers, emails and phone numbers. It can't recognise names, so read the diff yourself.

## How

```bash
npx srm-report contribute                    # shows the diff, sends nothing
npx srm-report contribute --yes --title "Add recipe for 21CSC203P case study" --body "What changed and why"
```

This needs the [GitHub CLI](https://cli.github.com), signed in. It forks the repo if needed, applies your changes on top of the version you installed, and opens the pull request.

If you use Claude Code, it will offer this at the end of a report. It only sends after you've seen the diff and said yes.

## How pull requests are handled

The maintainer reads each one and merges what's useful for other reports, reworking it where needed. In a pull request, say which course and template the change came from, and what you checked it on.
