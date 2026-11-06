# Writing the content

The report goes to faculty who sign it, and the user asks directly whether everything is real. So honesty rules come first.

## Everything real

- **Results:** re-run the tests, evaluations and benchmarks, then quote those numbers with their denominators ("0 false passes out of 23"). Don't quote from old notes or memory.
- **Not done means pending.** Write "pending" or "planned", and tell the user. Never write a finished-sounding sentence about work that hasn't happened.
- **Baselines:** if a comparison number is an estimate, say so in the sentence.
- **Screenshots** show the real system on fixture data, from a separate instance.
- **Publications:** never imply a submission or acceptance that hasn't happened.

## References

The template asks for at least 15, but every one must be real and actually relevant.

1. Search for the paper or standard. Don't write a reference from memory.
2. Open the primary source: publisher page, arXiv, DOI, the official PDF. Check authors, title, venue, year and pages.
3. Check that the sentence citing it matches what the source says. For laws and regulations, quote word for word from the official PDF, and keep the date and notification number.
4. List web sources with an accessed date.
5. If you can't verify one, remove it and tell the user.

`Report.prepare()` numbers citations by first appearance and stops the build if a reference is never cited or a citation has no reference. So add references as `key: text` and cite them as `[@key]`, and never type numbers by hand.

## Writing style

- Plain, specific prose in "we". Name the real component, the real number and the real file.
- Keep paragraphs under 12 lines, the template's limit; the checker warns on long ones.
- Introduce every figure and table in the text before it appears, using `{fig:id}` and `{tab:id}`.
- Use tables for lists of requirements, comparisons and results. Keep each short enough for one page; split a long one into two numbered tables.
- Use the `humanizer` skill's patterns as a scan on the prose before handing over: inflated claims, staged contrasts, stock AI words.
- Don't pad to reach the page minimum. If it's short, add real substance: sprint documents, test tables, failure analysis.

## Figures

White background, Times New Roman, print-friendly colours that survive greyscale, and no chart titles repeating the caption. Check images by looking at them after placement. Give every figure a short caption that says what it shows.

## Where things come from

Before asking the user, look through the repo, its README and PRD, test output and earlier reports. Use real source code for appendix listings, with the parts that carry the project's main claims chosen first. Put a note before them saying where the full code lives.
