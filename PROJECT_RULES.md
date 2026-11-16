# blacklist_gap - project rules

This file holds the STANDING rules this project followed throughout.

## A. After every task: log it in PROGRESS_LOG.md

After finishing any task (a buyer script, a bug fix, a design decision,
anything), add ONE entry at the bottom of `PROGRESS_LOG.md` section 5, using
the template already there. Every entry includes:

- **Asked**: what was requested
- **Done**: step by step, in the order it was actually done
- **Code changed**: each changed file, with its `.before`/`.diff` (or
  `.created`/`.deleted`) file names from `code_history\` (see rule B)
- **Results with numbers**: row counts, match counts, whatever is countable
- **Problems and source issues**: anything logged to `output/source_issues.csv`,
  anything that didn't work, anything skipped
- **Open items**: what is left, what needs a decision

**Never edit or delete an older entry.** If something earlier turns out to
be wrong, add a NEW entry that says so and corrects it - the log is a
history, not a live document.

## B. Before every code edit: back it up and diff it

Before changing ANY `.py` file, `PROJECT_RULES.md`, or `README.txt`:

1. Copy the file as it currently is to
   `code_history\<YYYY-MM-DD_HHMM>_<filename>.before`
2. Make the edit.
3. Write the exact difference to
   `code_history\<YYYY-MM-DD_HHMM>_<filename>.diff`
   using Python's `difflib.unified_diff` (standard library - no installs).
4. In the log entry (rule A), list the file with what changed, why, and the
   `.before`/`.diff` file names.

For a **new** file: save a copy as `<timestamp>_<filename>.created` instead
of `.before`/`.diff`.
For a **deleted** file: save the last copy as `<timestamp>_<filename>.deleted`
before removing it.

This applies to temporary/throwaway scripts too - if it's a `.py` file in
this project, it gets backed up before it's touched. This means ALL
scripts, including one-off utility/report builders that are not part of
the numbered buyer/pipeline chain (e.g. `build_pdfs_to_read.py`,
`build_network_firms_buyers.py`, `build_rate_contract_sources.py`) -
these scripts decide what counts as "confirmed" data to act
on, so their own history must be traceable the same way the pipeline
scripts' is. There is no "too small to back up" exception.

Use `scripts\record_change.py` to do steps 1 and 3 consistently (see its own
docstring for usage) - don't hand-roll this differently each time.

## C. Keep the log's reference sections current

Sections 1, 3, and 4 of `PROGRESS_LOG.md` are REFERENCE sections, not a
history - keep them edited in place to reflect the CURRENT state (unlike
section 5, which is append-only):

- **Section 1**: status table (one row per buyer: rows, format, done/blocked)
- **Section 3**: script list, output files, ID rules per buyer
- **Section 4**: known issues / open source_issues.csv items, current
  design decisions

When something changes (a new buyer, a fixed bug, a new ID rule), update
these sections directly so they always describe "how it is now" - the
step-by-step history of HOW it got that way lives in section 5 instead.

## D. Standing engineering rules for this project

- **Respect robots.txt.** If a site blocks a path, give the exact link and
  file name so it can be saved by hand - never bypass it.
- **No CAPTCHA solving, no stealth/fingerprint evasion, ever** - even for a
  single page, even with a plausible excuse. If a site shows a CAPTCHA or an
  explicit block page to an automated request, stop and log it.
- **Certificate verification stays ON, always** - never `verify=False`, even
  to work around a server's own TLS misconfiguration (use a custom
  `SSLContext` with a narrow fix instead, as done for MPPHSCL's legacy
  renegotiation requirement).
- **Every Excel output is plain typed values** - no formulas, no merged
  cells, nothing locked. Arial font, via `common.write_sheet()`.
- **Entry IDs must be stable across reruns** - built from the source (a
  serial, an order number, a date+code), never from a running counter. If
  two rows still collide, use `-b`/`-c`/... and print a warning - never
  silently drop or silently rename without saying so.
- **Log source problems instead of fixing them silently** - anything odd in
  the data (a missing date, a corrupted field, a name that doesn't split
  right) goes in `output/source_issues.csv` with enough detail to act on
  later, rather than being guessed at or quietly patched.
- **Nothing is ever auto-merged.** Fuzzy/candidate matches are reports for a
  human to decide on (an empty "My decision" column) - never applied
  automatically.

## E. Keep the manual checklist current

After every task, review the manual checklist (section 6 of
`PROGRESS_LOG.md`) and list which items are still open - separately from
the task's own "Open items" (rule A). If every checklist item is done,
say so explicitly.

## F. OCR handling rules

Once OCR is available, any text it produces is lower-confidence than a
PDF's own embedded text and must be handled differently everywhere:

- Save OCR output as `data\text\ocr\<file>_p<page>.txt`, next to the
  source PDF it came from - never only in memory or only in the final
  spreadsheet.
- Any value that came from OCR gets `Status` = `OCR - check` and the
  exact page number recorded in `Page / row in source` - never silently
  folded in as if it were normal extracted text.
- Dates and numbers read via OCR are NOT trusted until they have been
  checked them against the PDF by eye. Never base a "Possible breach",
  "Confirmed", or similarly asserted finding on an unchecked OCR value -
  such a finding stays provisional (e.g. "Possible breach - pending OCR
  check") until it is cleared by hand.
- If OCR confidence for a page is low, log it to `output/source_issues.csv`
  rather than silently accepting a low-confidence read.
